from __future__ import annotations

import time
from pathlib import Path
from typing import Callable

from .charge_parser import parse_voltage_sampling_workbook
from .charge_statistics_service import (
    apply_tail_fill_check,
    compute_charge_metrics,
    compute_temperature_metrics,
    resolve_output_path,
)
from .errors import AppError
from .excel_render import render_charge_workbook
from .file_collect import collect_statistics_excel_files
from .models import BatchResult, ProcessItemResult


Logger = Callable[[str, str], None] | None


def _emit(logger: Logger, level: str, message: str) -> None:
    if logger is not None:
        logger(level, message)


def _format_wait_seconds(value: float) -> str:
    if float(value).is_integer():
        return str(int(value))
    return f"{value:.2f}".rstrip("0").rstrip(".")


def process_voltage_sampling_statistics(
    inputs: list[Path],
    output_dir: Path,
    sampling_resistance_milliohm: float,
    logger: Logger = None,
    chunk_size: int | None = None,
    wait_seconds: float = 0.0,
) -> BatchResult:
    if sampling_resistance_milliohm <= 0:
        raise AppError(
            "INVALID_SAMPLING_RESISTANCE",
            "采样电阻阻值必须大于 0 mΩ",
            detail=str(sampling_resistance_milliohm),
        )

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
    sampling_resistance_ohm = sampling_resistance_milliohm / 1000.0

    for index, excel_file in enumerate(excel_files, start=1):
        _emit(logger, "INFO", f"开始处理（分压采集测试）：{excel_file}")
        try:
            dataset = parse_voltage_sampling_workbook(excel_file, sampling_resistance_ohm)
            apply_tail_fill_check(dataset)
            charge_metrics = compute_charge_metrics(dataset)
            temperature_metrics = compute_temperature_metrics(dataset)
            output_path = resolve_output_path(output_dir, dataset.stem)
            render_charge_workbook(dataset, charge_metrics, temperature_metrics, output_path)
            results.append(
                ProcessItemResult(
                    name=excel_file.name,
                    status="success",
                    output_path=output_path,
                    warnings=dataset.warnings.copy(),
                )
            )
            _emit(logger, "INFO", f"[成功] 输出成功：{output_path}")
            for warning in dataset.warnings:
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
