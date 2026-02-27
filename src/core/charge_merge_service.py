from __future__ import annotations

import re
from pathlib import Path
from typing import Callable

from .charge_parser import parse_charge_workbook, parse_voltage_csv
from .charge_statistics_service import (
    apply_tail_fill_check,
    compute_charge_metrics,
    compute_temperature_metrics,
    resolve_output_path,
)
from .errors import AppError
from .excel_render import render_charge_workbook
from .file_collect import collect_merge_groups
from .models import BatchResult, ProcessItemResult


Logger = Callable[[str, str], None] | None


def _emit(logger: Logger, level: str, message: str) -> None:
    if logger is not None:
        logger(level, message)


def _stem_from_pair_error(message: str) -> str:
    match = re.match(r"\[(.+?)\]", message)
    if match:
        return match.group(1)
    return message


def process_charge_merge(inputs: list[Path], output_dir: Path, logger: Logger = None) -> BatchResult:
    output_dir.mkdir(parents=True, exist_ok=True)
    groups, pair_errors, collection_warnings = collect_merge_groups(inputs)
    for warning in collection_warnings:
        _emit(logger, "WARN", warning)

    results: list[ProcessItemResult] = []
    for pair_error in pair_errors:
        stem = _stem_from_pair_error(pair_error)
        results.append(ProcessItemResult(name=stem, status="failed", error=pair_error))
        _emit(logger, "ERROR", pair_error)

    if not groups and not pair_errors:
        warning_message = "未检测到可处理的 xlsx+csv 配对文件组"
        _emit(logger, "WARN", warning_message)
        return BatchResult(total=0, success=0, failed=0, items=[], warnings=[warning_message])

    for group in groups:
        _emit(logger, "INFO", f"开始处理（合并后统计数据）：{group.stem}")
        try:
            dataset = parse_charge_workbook(group.xlsx_path, require_voltage=False)
            csv_mapping, csv_warnings = parse_voltage_csv(group.csv_path)
            dataset.warnings.extend(csv_warnings)

            merged_voltages: list[float | None] = []
            missing_timestamps: list[str] = []
            for dt_value in dataset.datetimes:
                voltage = csv_mapping.get(dt_value)
                if voltage is None:
                    missing_timestamps.append(dt_value.strftime("%Y-%m-%d %H:%M:%S"))
                merged_voltages.append(voltage)
            if missing_timestamps:
                preview = ", ".join(missing_timestamps[:3])
                raise AppError(
                    "MERGE_TIME_MISMATCH",
                    "存在 Excel 时间点无法在 CSV 中匹配",
                    detail=f"{group.stem}: {preview}",
                )

            dataset.voltages_v = merged_voltages
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
            _emit(logger, "INFO", f"输出成功：{output_path}")
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
            _emit(logger, "ERROR", f"{group.stem} 处理失败：{exc}")
        except Exception as exc:  # pragma: no cover
            results.append(
                ProcessItemResult(
                    name=group.stem,
                    status="failed",
                    error=f"[UNEXPECTED] {exc}",
                )
            )
            _emit(logger, "ERROR", f"{group.stem} 处理失败：[UNEXPECTED] {exc}")

    success_count = sum(1 for item in results if item.status == "success")
    failed_count = len(results) - success_count
    return BatchResult(
        total=len(results),
        success=success_count,
        failed=failed_count,
        items=results,
    )
