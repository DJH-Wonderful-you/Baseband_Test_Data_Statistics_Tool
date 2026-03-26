from __future__ import annotations

import re
import time
from pathlib import Path
from typing import Any, Callable

from .charge_parser import normalize_series_polarity, parse_charge_workbook, parse_time_value_csv
from .charge_statistics_service import (
    apply_tail_fill_check,
    compute_charge_metrics,
    compute_temperature_metrics,
    resolve_output_path,
)
from .errors import AppError
from .excel_render import render_charge_workbook
from .file_collect import collect_merge_groups
from .models import BatchResult, ChargeDataset, ProcessItemResult


Logger = Callable[[str, str], None] | None


def _emit(logger: Logger, level: str, message: str) -> None:
    if logger is not None:
        logger(level, message)


def _format_wait_seconds(value: float) -> str:
    if float(value).is_integer():
        return str(int(value))
    return f"{value:.2f}".rstrip("0").rstrip(".")


def _stem_from_pair_error(message: str) -> str:
    match = re.match(r"\[(.+?)\]", message)
    if match:
        return match.group(1)
    return message


def _filter_dataset_by_indices(dataset: ChargeDataset, keep_indices: list[int]) -> None:
    dataset.index_values = [dataset.index_values[idx] for idx in keep_indices]
    dataset.datetimes = [dataset.datetimes[idx] for idx in keep_indices]
    dataset.date_strings = [dataset.date_strings[idx] for idx in keep_indices]
    dataset.time_strings = [dataset.time_strings[idx] for idx in keep_indices]
    dataset.currents_ma = [dataset.currents_ma[idx] for idx in keep_indices]
    dataset.voltages_v = [dataset.voltages_v[idx] for idx in keep_indices]
    dataset.pen_temps_c = [dataset.pen_temps_c[idx] for idx in keep_indices]
    dataset.env_temps_c = [dataset.env_temps_c[idx] for idx in keep_indices]
    dataset.extras = {
        header: [values[idx] for idx in keep_indices] for header, values in dataset.extras.items()
    }


def _normalize_header(text: str) -> str:
    return text.strip().lower().replace("（", "(").replace("）", ")").replace(" ", "")


def _to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    match = re.search(r"[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?", text)
    if not match:
        return None
    return float(match.group(0))


def _contains_ol_text(value: Any) -> bool:
    if not isinstance(value, str):
        return False
    return "o.l" in value.strip().lower()


def _replace_ol_with_next_value(values: list[Any]) -> int:
    if len(values) < 2:
        return 0
    replaced_count = 0
    for idx in range(len(values) - 2, -1, -1):
        current_value = values[idx]
        if _to_float(current_value) is not None:
            continue
        if not _contains_ol_text(current_value):
            continue
        values[idx] = values[idx + 1]
        replaced_count += 1
    return replaced_count


def _resolve_csv_unit(unit_value: str | None) -> tuple[str, float]:
    normalized_unit = _normalize_header(unit_value or "")
    if normalized_unit == "v":
        return "voltage", 1.0
    if normalized_unit == "ma":
        return "current", 1.0
    if normalized_unit == "a":
        return "current", 1000.0
    raise AppError(
        "CSV_UNIT_INVALID",
        "CSV Unit 列首个有效值仅支持 V / mA / A",
        detail=f"unit={unit_value}",
    )


def _parse_numeric_series(values: list[Any], file_name: str) -> list[float]:
    parsed_values: list[float] = []
    for idx, raw_value in enumerate(values, start=1):
        numeric_value = _to_float(raw_value)
        if numeric_value is None:
            raise AppError(
                "CSV_VALUE_INVALID",
                "CSV Value 列存在无法解析的数据",
                detail=f"{file_name} index={idx} value={raw_value}",
            )
        parsed_values.append(numeric_value)
    return parsed_values


def process_charge_merge(
    inputs: list[Path],
    output_dir: Path,
    logger: Logger = None,
    chunk_size: int | None = None,
    wait_seconds: float = 0.0,
) -> BatchResult:
    output_dir.mkdir(parents=True, exist_ok=True)
    groups, pair_errors, collection_warnings = collect_merge_groups(inputs)
    for warning in collection_warnings:
        _emit(logger, "WARN", warning)

    results: list[ProcessItemResult] = []
    for pair_error in pair_errors:
        stem = _stem_from_pair_error(pair_error)
        results.append(ProcessItemResult(name=stem, status="failed", error=pair_error))
        _emit(logger, "ERROR", f"[失败] {pair_error}")

    if not groups and not pair_errors:
        warning_message = "未检测到可处理的 Excel+csv 配对文件组（Excel: .xlsx/.xls）"
        _emit(logger, "WARN", warning_message)
        return BatchResult(total=0, success=0, failed=0, items=[], warnings=[warning_message])

    total_groups = len(groups)
    use_pacing = chunk_size is not None and chunk_size > 0 and wait_seconds > 0
    effective_chunk_size = chunk_size if use_pacing else 0
    for index, group in enumerate(groups, start=1):
        _emit(logger, "INFO", f"开始处理（合并后统计数据）：{group.stem}")
        try:
            dataset = parse_charge_workbook(
                group.excel_path,
                require_voltage=False,
                normalize_duplicate_seconds=True,
                require_current=False,
            )
            csv_mapping, csv_warnings, csv_unit = parse_time_value_csv(group.csv_path)
            dataset.warnings.extend(csv_warnings)

            excel_has_current = any(value is not None for value in dataset.currents_ma)
            excel_has_voltage = any(value is not None for value in dataset.voltages_v)
            if excel_has_current == excel_has_voltage:
                raise AppError(
                    "MERGE_SOURCE_AMBIGUOUS",
                    "Excel 需仅包含“电流”或“电压”其中一列，CSV 再提供另一列进行合并",
                    detail=group.stem,
                )

            matched_indices: list[int] = []
            matched_csv_raw_values: list[Any] = []
            for idx, dt_value in enumerate(dataset.datetimes):
                csv_value = csv_mapping.get(dt_value)
                if csv_value is None:
                    continue
                matched_indices.append(idx)
                matched_csv_raw_values.append(csv_value)
            if not matched_indices:
                raise AppError(
                    "MERGE_TIME_NO_INTERSECTION",
                    "Excel 与 CSV 的时间点没有交集，无法合并",
                    detail=group.stem,
                )

            dropped_count = dataset.row_count() - len(matched_indices)
            if dropped_count > 0:
                dataset.warnings.append(
                    f"Excel 与 CSV 时间点已按交集匹配：保留 {len(matched_indices)} 条，过滤 {dropped_count} 条"
                )
            csv_ol_replaced = _replace_ol_with_next_value(matched_csv_raw_values)
            if csv_ol_replaced > 0:
                dataset.warnings.append(f"CSV Value 列检测到 {csv_ol_replaced} 个 O.L，已使用后一个值替换")
            matched_csv_values = _parse_numeric_series(matched_csv_raw_values, group.csv_path.name)

            _filter_dataset_by_indices(dataset, matched_indices)
            csv_role, current_factor = _resolve_csv_unit(csv_unit)
            if csv_role == "voltage":
                if not excel_has_current:
                    raise AppError(
                        "MERGE_SOURCE_CONFLICT",
                        "CSV Unit=V 时，Excel 需提供电流列，CSV 提供电压列",
                        detail=group.stem,
                    )
                dataset.voltages_v, voltage_flipped = normalize_series_polarity(matched_csv_values)
                if voltage_flipped:
                    dataset.warnings.append("检测到电压极性与预期相反，已将整列电压取相反数后再进行后续统计")
            else:
                if not excel_has_voltage:
                    raise AppError(
                        "MERGE_SOURCE_CONFLICT",
                        "CSV Unit=mA/A 时，Excel 需提供电压列，CSV 提供电流列",
                        detail=group.stem,
                    )
                dataset.currents_ma = [
                    (value * current_factor if value is not None else None) for value in matched_csv_values
                ]
                dataset.currents_ma, current_flipped = normalize_series_polarity(dataset.currents_ma)
                if current_flipped:
                    dataset.warnings.append("检测到电流方向与预期相反，已将整列电流取相反数后再进行后续统计")
            apply_tail_fill_check(dataset)
            metrics = compute_charge_metrics(dataset)
            temp_metrics = compute_temperature_metrics(dataset)
            output_path = resolve_output_path(output_dir, group.stem)
            render_charge_workbook(dataset, metrics, temp_metrics, output_path)
            results.append(
                ProcessItemResult(
                    name=group.stem,
                    status="success",
                    output_path=output_path,
                    warnings=dataset.warnings.copy(),
                )
            )
            _emit(logger, "INFO", f"[成功] 输出成功：{output_path}")
            for warning in dataset.warnings:
                _emit(logger, "WARN", f"{group.stem}：{warning}")
        except AppError as exc:
            results.append(
                ProcessItemResult(
                    name=group.stem,
                    status="failed",
                    error=str(exc),
                )
            )
            _emit(logger, "ERROR", f"[失败] {group.stem} 处理失败：{exc}")
        except Exception as exc:  # pragma: no cover
            results.append(
                ProcessItemResult(
                    name=group.stem,
                    status="failed",
                    error=f"[UNEXPECTED] {exc}",
                )
            )
            _emit(logger, "ERROR", f"[失败] {group.stem} 处理失败：[UNEXPECTED] {exc}")
        if use_pacing and index % effective_chunk_size == 0 and index < total_groups:
            remaining = total_groups - index
            _emit(
                logger,
                "INFO",
                f"已处理 {index}/{total_groups} 组文件，等待 {_format_wait_seconds(wait_seconds)} 秒后继续（剩余 {remaining} 组）",
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
