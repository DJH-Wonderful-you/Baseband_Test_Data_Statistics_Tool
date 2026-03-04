from __future__ import annotations

import re
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable

from .charge_parser import parse_charge_workbook
from .charge_statistics_service import apply_tail_fill_check, resolve_output_path
from .endurance_excel_render import (
    IndicatorLogRow,
    render_endurance_duration_workbook,
    render_endurance_indicator_workbook,
)
from .endurance_parser import TimedBatteryEvent, parse_battery_log_file
from .errors import AppError
from .file_collect import collect_endurance_indicator_groups, collect_statistics_excel_files
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


def compute_endurance_duration(dataset: ChargeDataset) -> timedelta:
    indexed_voltages = [(idx, value) for idx, value in enumerate(dataset.voltages_v) if value is not None]
    if not indexed_voltages:
        raise AppError("MISSING_VOLTAGE_VALUE", "电压列无可用数据，无法计算续航时长", detail=dataset.source_path.name)
    min_voltage = min(value for _, value in indexed_voltages)
    end_index = max(idx for idx, value in indexed_voltages if value == min_voltage)
    start_time = dataset.datetimes[0]
    end_time = dataset.datetimes[end_index]
    return end_time - start_time


def _row_from_dataset(dataset: ChargeDataset, row_index: int, level: int) -> IndicatorLogRow:
    return IndicatorLogRow(
        log_datetime=dataset.datetimes[row_index],
        level=level,
        mapped_voltage_v=dataset.voltages_v[row_index],
        log_date=dataset.date_strings[row_index],
    )


def _build_indicator_rows_from_timed_events(
    dataset: ChargeDataset,
    events: list[TimedBatteryEvent],
) -> list[IndicatorLogRow]:
    if not events:
        raise AppError("TEXT_NO_EVENT", "文本中未提取到可用电量数据", detail=dataset.source_path.name)

    matched_indices: list[int] = []
    cursor = 0
    for event in events:
        matched_index: int | None = None
        for idx in range(cursor, dataset.row_count()):
            if dataset.datetimes[idx].time() == event.time_value:
                matched_index = idx
                break
        if matched_index is None:
            raise AppError(
                "TEXT_TIME_NOT_MATCHED",
                "文本中的时间未能在 Excel “时间 (s)”列中匹配到对应项",
                detail=(
                    f"{dataset.source_path.name} time={event.time_value.strftime('%H:%M:%S')} level={event.level}"
                ),
            )
        matched_indices.append(matched_index)
        cursor = matched_index

    rows: list[IndicatorLogRow] = [_row_from_dataset(dataset, matched_indices[0], events[0].level)]
    for idx in range(1, len(events)):
        prev_index = matched_indices[idx - 1]
        curr_index = matched_indices[idx]
        prev_level = events[idx - 1].level
        curr_level = events[idx].level
        for row_index in range(prev_index + 1, curr_index + 1):
            rows.append(_row_from_dataset(dataset, row_index, prev_level))
        rows.append(_row_from_dataset(dataset, curr_index, curr_level))
    return rows


def _build_indicator_rows_from_special_points(
    levels_and_voltages: list[tuple[int, float]],
) -> list[IndicatorLogRow]:
    if not levels_and_voltages:
        raise AppError("TEXT_NO_EVENT", "文本中未提取到可用电量数据")
    start_dt = datetime(1900, 1, 1, 0, 0, 0)
    rows: list[IndicatorLogRow] = []
    for index, (level, voltage_v) in enumerate(levels_and_voltages):
        rows.append(
            IndicatorLogRow(
                log_datetime=start_dt + timedelta(seconds=index),
                level=level,
                mapped_voltage_v=voltage_v,
                log_date=None,
            )
        )
    return rows


def process_endurance_duration_statistics(
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
        _emit(logger, "INFO", f"开始处理（续航时长统计）：{excel_file}")
        try:
            dataset = parse_charge_workbook(excel_file, require_voltage=True)
            apply_tail_fill_check(dataset)
            endurance_duration = compute_endurance_duration(dataset)
            output_path = resolve_output_path(output_dir, dataset.stem)
            render_endurance_duration_workbook(dataset, endurance_duration, output_path)
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
            results.append(ProcessItemResult(name=excel_file.name, status="failed", error=str(exc)))
            _emit(logger, "ERROR", f"[失败] {excel_file.name} 处理失败：{exc}")
        except Exception as exc:  # pragma: no cover
            results.append(
                ProcessItemResult(name=excel_file.name, status="failed", error=f"[UNEXPECTED] {exc}")
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


def process_endurance_indicator_statistics(
    inputs: list[Path],
    output_dir: Path,
    logger: Logger = None,
    chunk_size: int | None = None,
    wait_seconds: float = 0.0,
) -> BatchResult:
    output_dir.mkdir(parents=True, exist_ok=True)
    groups, pair_errors, collection_warnings = collect_endurance_indicator_groups(inputs)
    for warning in collection_warnings:
        _emit(logger, "WARN", warning)

    results: list[ProcessItemResult] = []
    for pair_error in pair_errors:
        stem = _stem_from_pair_error(pair_error)
        results.append(ProcessItemResult(name=stem, status="failed", error=pair_error))
        _emit(logger, "ERROR", f"[失败] {pair_error}")

    if not groups and not pair_errors:
        warning_message = "未检测到可处理的 Excel+文本 配对文件组（Excel: .xlsx/.xls，文本: .txt/.log）"
        _emit(logger, "WARN", warning_message)
        return BatchResult(total=0, success=0, failed=0, items=[], warnings=[warning_message])

    total_groups = len(groups)
    use_pacing = chunk_size is not None and chunk_size > 0 and wait_seconds > 0
    effective_chunk_size = chunk_size if use_pacing else 0
    for index, group in enumerate(groups, start=1):
        _emit(logger, "INFO", f"开始处理（电量指示统计）：{group.stem}")
        try:
            dataset = parse_charge_workbook(group.excel_path, require_voltage=True)
            apply_tail_fill_check(dataset)

            log_parse_result = parse_battery_log_file(group.text_path)
            dataset.warnings.extend(log_parse_result.warnings)
            if log_parse_result.mode == "timed":
                log_rows = _build_indicator_rows_from_timed_events(dataset, log_parse_result.timed_events)
                include_log_date = True
            else:
                special_values = [
                    (point.level, point.voltage_v) for point in log_parse_result.special_points
                ]
                log_rows = _build_indicator_rows_from_special_points(special_values)
                include_log_date = False

            output_path = resolve_output_path(output_dir, group.stem)
            render_endurance_indicator_workbook(
                dataset=dataset,
                log_rows=log_rows,
                include_log_date=include_log_date,
                output_path=output_path,
            )
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
            results.append(ProcessItemResult(name=group.stem, status="failed", error=str(exc)))
            _emit(logger, "ERROR", f"[失败] {group.stem} 处理失败：{exc}")
        except Exception as exc:  # pragma: no cover
            results.append(
                ProcessItemResult(name=group.stem, status="failed", error=f"[UNEXPECTED] {exc}")
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
