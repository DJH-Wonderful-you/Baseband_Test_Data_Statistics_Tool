from __future__ import annotations

import re
import time
from dataclasses import dataclass
from datetime import timedelta
from pathlib import Path
from typing import Any, Callable

from .charge_parser import parse_charge_workbook
from .charge_protection_render import render_charge_protection_workbook
from .charge_statistics_service import apply_tail_fill_check, resolve_output_path
from .errors import AppError
from .file_collect import collect_statistics_excel_files
from .models import (
    BatchResult,
    ChargeDataset,
    ChargeProtectionDataset,
    ChargeProtectionMetrics,
    ProcessItemResult,
)


Logger = Callable[[str, str], None] | None
HIGH_ONLY_MODE = "high"
LOW_ONLY_MODE = "low"
HIGH_LOW_MODE = "high_low"


@dataclass(slots=True)
class ChargeProtectionInterval:
    start_index: int
    end_index: int | None
    category: str | None = None


def _emit(logger: Logger, level: str, message: str) -> None:
    if logger is not None:
        logger(level, message)


def _format_wait_seconds(value: float) -> str:
    if float(value).is_integer():
        return str(int(value))
    return f"{value:.2f}".rstrip("0").rstrip(".")


def _is_empty(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    return False


def _safe_str(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _to_float(value: Any) -> float | None:
    if _is_empty(value) or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = _safe_str(value).replace(",", "")
    match = re.search(r"[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?", text)
    if match is None:
        return None
    return float(match.group(0))


def _normalize_header(text: str) -> str:
    normalized = text.strip().lower()
    normalized = normalized.replace("（", "(").replace("）", ")")
    normalized = normalized.replace(" ", "")
    return normalized


def _series_has_data(values: list[float | None]) -> bool:
    return any(value is not None for value in values)


def _normalize_series_length(values: list[float | None], target_length: int) -> list[float | None]:
    if len(values) == target_length:
        return values
    if len(values) > target_length:
        return values[:target_length]
    return values + [None] * (target_length - len(values))


def _extract_extra_series(
    dataset: ChargeDataset,
    *,
    required_keywords: tuple[str, ...],
) -> list[float | None] | None:
    for header in dataset.extra_headers_order:
        normalized = _normalize_header(header)
        if not all(keyword in normalized for keyword in required_keywords):
            continue
        values = dataset.extras.get(header, [])
        converted = [_to_float(value) for value in values]
        normalized_series = _normalize_series_length(converted, dataset.row_count())
        if _series_has_data(normalized_series):
            return normalized_series
    return None


def build_charge_protection_dataset(dataset: ChargeDataset) -> ChargeProtectionDataset:
    if not dataset.has_current_data():
        raise AppError("MISSING_CURRENT", "未检测到“电流”相关列或有效数据", detail=dataset.source_path.name)
    if not dataset.has_voltage_data():
        raise AppError("MISSING_VOLTAGE", "未检测到“电压”相关列或有效数据", detail=dataset.source_path.name)

    env_temps_c = _extract_extra_series(dataset, required_keywords=("环境", "温度(°c)")) or []
    if not _series_has_data(env_temps_c):
        env_temps_c = _normalize_series_length(dataset.env_temps_c.copy(), dataset.row_count())
    if not _series_has_data(env_temps_c):
        raise AppError("MISSING_ENV_TEMP", "未检测到“环境温度”相关列或有效数据", detail=dataset.source_path.name)

    cell_temps_c = _extract_extra_series(dataset, required_keywords=("电芯", "温度(°c)"))
    if cell_temps_c is None or not _series_has_data(cell_temps_c):
        raise AppError("MISSING_CELL_TEMP", "未检测到“电芯温度”相关列或有效数据", detail=dataset.source_path.name)

    return ChargeProtectionDataset(
        source_path=dataset.source_path,
        stem=dataset.stem,
        index_values=dataset.index_values.copy(),
        datetimes=dataset.datetimes.copy(),
        date_strings=dataset.date_strings.copy(),
        time_strings=dataset.time_strings.copy(),
        currents_ma=dataset.currents_ma.copy(),
        voltages_v=dataset.voltages_v.copy(),
        env_temps_c=env_temps_c,
        cell_temps_c=cell_temps_c,
        warnings=dataset.warnings.copy(),
    )


def _interval_end_index(dataset: ChargeProtectionDataset, interval: ChargeProtectionInterval) -> int:
    if interval.end_index is not None:
        return interval.end_index
    return dataset.row_count() - 1


def _interval_duration(dataset: ChargeProtectionDataset, interval: ChargeProtectionInterval) -> timedelta:
    end_index = _interval_end_index(dataset, interval)
    return dataset.datetimes[end_index] - dataset.datetimes[interval.start_index]


def _find_charge_protection_intervals(dataset: ChargeProtectionDataset) -> list[ChargeProtectionInterval]:
    intervals: list[ChargeProtectionInterval] = []
    open_start_index: int | None = None
    for idx in range(dataset.row_count() - 1):
        current_value = dataset.currents_ma[idx]
        next_value = dataset.currents_ma[idx + 1]
        if current_value is None or next_value is None:
            continue
        if open_start_index is None:
            if current_value > 1 and next_value < 1:
                open_start_index = idx
            continue
        if current_value < 1 and next_value > 1:
            intervals.append(ChargeProtectionInterval(start_index=open_start_index, end_index=idx))
            open_start_index = None

    if open_start_index is not None:
        dataset.warnings.append("检测到停充区间起点但未检测到终点，已按开放区间继续统计")
        intervals.append(ChargeProtectionInterval(start_index=open_start_index, end_index=None))

    if not intervals:
        raise AppError(
            "PROTECTION_START_NOT_FOUND",
            "未检测到任何停充区间起点，无法计算充电保护温度",
            detail=dataset.source_path.name,
        )
    return intervals


def _classify_interval(
    dataset: ChargeProtectionDataset,
    interval: ChargeProtectionInterval,
) -> str | None:
    end_index = _interval_end_index(dataset, interval)
    interval_temps = [
        value for value in dataset.cell_temps_c[interval.start_index : end_index + 1] if value is not None
    ]
    if not interval_temps:
        raise AppError(
            "CELL_TEMP_EMPTY",
            "停充区间内缺少有效的电芯温度数据",
            detail=f"{dataset.source_path.name} start={interval.start_index + 1} end={end_index + 1}",
        )

    has_high = any(value > 45 for value in interval_temps)
    has_low = any(value < 0 for value in interval_temps)
    if has_high and has_low:
        raise AppError(
            "PROTECTION_INTERVAL_AMBIGUOUS",
            "同一停充区间同时满足高温和低温判定条件，请检查测试数据",
            detail=f"{dataset.source_path.name} start={interval.start_index + 1} end={end_index + 1}",
        )
    if has_high:
        return HIGH_ONLY_MODE
    if has_low:
        return LOW_ONLY_MODE

    dataset.warnings.append(
        "检测到虚假停充区间，区间内未满足高温或低温判定条件，已忽略"
        f"（start={interval.start_index + 1}, end={end_index + 1}）"
    )
    return None


def _cell_temp_at_index(
    dataset: ChargeProtectionDataset,
    index: int,
    *,
    purpose: str,
) -> float | None:
    value = dataset.cell_temps_c[index]
    if value is None:
        dataset.warnings.append(f"{purpose}对应时间点的电芯温度为空，已记为未检测到符合要求的数据，请人工查看")
    return value


def compute_charge_protection_metrics(dataset: ChargeProtectionDataset) -> ChargeProtectionMetrics:
    intervals = _find_charge_protection_intervals(dataset)
    high_intervals: list[ChargeProtectionInterval] = []
    low_intervals: list[ChargeProtectionInterval] = []
    for interval in intervals:
        category = _classify_interval(dataset, interval)
        if category is None:
            continue
        interval.category = category
        if category == HIGH_ONLY_MODE:
            high_intervals.append(interval)
        else:
            low_intervals.append(interval)

    if not high_intervals and not low_intervals:
        raise AppError(
            "PROTECTION_CLASS_NOT_FOUND",
            "未检测到符合高温或低温判定条件的停充区间",
            detail=dataset.source_path.name,
        )

    if "高低温" in dataset.stem and (not high_intervals or not low_intervals):
        raise AppError(
            "PROTECTION_HIGH_LOW_MISSING",
            "文件名包含“高低温”，但未同时检测到高温区间和低温区间",
            detail=dataset.source_path.name,
        )

    mode = HIGH_LOW_MODE if high_intervals and low_intervals else HIGH_ONLY_MODE if high_intervals else LOW_ONLY_MODE

    high_protect_temp_c: float | None = None
    high_resume_temp_c: float | None = None
    if high_intervals:
        longest_high = max(high_intervals, key=lambda item: _interval_duration(dataset, item))
        high_protect_temp_c = _cell_temp_at_index(dataset, longest_high.start_index, purpose="高温充电保护温度")
        last_high = high_intervals[-1]
        if last_high.end_index is not None:
            high_resume_temp_c = _cell_temp_at_index(
                dataset,
                last_high.end_index,
                purpose="高温充电复充温度",
            )

    low_protect_temp_c: float | None = None
    low_resume_temp_c: float | None = None
    if low_intervals:
        longest_low = max(low_intervals, key=lambda item: _interval_duration(dataset, item))
        low_protect_temp_c = _cell_temp_at_index(dataset, longest_low.start_index, purpose="低温充电保护温度")
        last_low = low_intervals[-1]
        if last_low.end_index is not None:
            low_resume_temp_c = _cell_temp_at_index(
                dataset,
                last_low.end_index,
                purpose="低温充电复充温度",
            )

    return ChargeProtectionMetrics(
        mode=mode,
        high_protect_temp_c=high_protect_temp_c,
        high_resume_temp_c=high_resume_temp_c,
        low_protect_temp_c=low_protect_temp_c,
        low_resume_temp_c=low_resume_temp_c,
    )


def process_charge_protection_statistics(
    inputs: list[Path],
    output_dir: Path,
    logger: Logger = None,
    chunk_size: int | None = None,
    wait_seconds: float = 0.0,
) -> BatchResult:
    output_dir.mkdir(parents=True, exist_ok=True)
    excel_files, collection_warnings = collect_statistics_excel_files(inputs)
    for warning in collection_warnings:
        _emit(logger, "WARN", warning)

    if not excel_files:
        warning_message = "未检测到可处理的 Excel 文件（.xlsx/.xls）"
        _emit(logger, "WARN", warning_message)
        return BatchResult(total=0, success=0, failed=0, items=[], warnings=[warning_message])

    results: list[ProcessItemResult] = []
    total_files = len(excel_files)
    use_pacing = chunk_size is not None and chunk_size > 0 and wait_seconds > 0
    effective_chunk_size = chunk_size if use_pacing else 0
    for index, excel_file in enumerate(excel_files, start=1):
        _emit(logger, "INFO", f"开始处理（充电保护测试）：{excel_file}")
        try:
            parsed_dataset = parse_charge_workbook(excel_file, require_voltage=True)
            apply_tail_fill_check(parsed_dataset)
            protection_dataset = build_charge_protection_dataset(parsed_dataset)
            protection_metrics = compute_charge_protection_metrics(protection_dataset)
            output_path = resolve_output_path(output_dir, protection_dataset.stem)
            render_charge_protection_workbook(protection_dataset, protection_metrics, output_path)
            results.append(
                ProcessItemResult(
                    name=excel_file.name,
                    status="success",
                    output_path=output_path,
                    warnings=protection_dataset.warnings.copy(),
                )
            )
            _emit(logger, "INFO", f"[成功] 输出成功：{output_path}")
            for warning in protection_dataset.warnings:
                _emit(logger, "WARN", f"{excel_file.name}：{warning}")
        except AppError as exc:
            results.append(
                ProcessItemResult(
                    name=excel_file.name,
                    status="failed",
                    error=str(exc),
                )
            )
            _emit(logger, "ERROR", f"[失败] {excel_file.name} 处理失败：{exc}")
        except Exception as exc:  # pragma: no cover
            results.append(
                ProcessItemResult(
                    name=excel_file.name,
                    status="failed",
                    error=f"[UNEXPECTED] {exc}",
                )
            )
            _emit(logger, "ERROR", f"[失败] {excel_file.name} 处理失败：[UNEXPECTED] {exc}")

        if use_pacing and index % effective_chunk_size == 0 and index < total_files:
            remaining = total_files - index
            _emit(
                logger,
                "INFO",
                f"已处理 {index}/{total_files} 个文件，等待 {_format_wait_seconds(wait_seconds)} 秒后继续（剩余 {remaining} 个）",
            )
            time.sleep(wait_seconds)

    success_count = sum(1 for item in results if item.status == "success")
    failed_count = len(results) - success_count
    return BatchResult(
        total=len(results),
        success=success_count,
        failed=failed_count,
        items=results,
    )
