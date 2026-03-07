from __future__ import annotations

import re
import time
from datetime import date, datetime, time as dt_time, timedelta
from pathlib import Path
from typing import Callable

from .charge_parser import parse_charge_workbook
from .charge_statistics_service import apply_tail_fill_check, resolve_output_path
from .endurance_excel_render import (
    IndicatorLogRow,
    render_endurance_duration_workbook,
    render_endurance_indicator_workbook,
    render_endurance_single_log_workbook,
)
from .endurance_parser import SpecialBatteryPoint, TimedBatteryEvent, parse_battery_log_file
from .errors import AppError
from .file_collect import collect_endurance_indicator_groups, collect_files, collect_statistics_excel_files
from .models import BatchResult, ChargeDataset, ProcessItemResult


Logger = Callable[[str, str], None] | None
_DATE_IN_NAME_PATTERN = re.compile(r"(?P<year>20\d{2})[-_./](?P<month>\d{1,2})[-_./](?P<day>\d{1,2})")


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


def _build_date_value(year_text: str, month_text: str, day_text: str) -> date | None:
    try:
        return date(year=int(year_text), month=int(month_text), day=int(day_text))
    except ValueError:
        return None


def _infer_start_date_from_path(path: Path) -> date:
    for candidate in (path.stem, path.name):
        match = _DATE_IN_NAME_PATTERN.search(candidate)
        if match is None:
            continue
        parsed = _build_date_value(
            match.group("year"),
            match.group("month"),
            match.group("day"),
        )
        if parsed is not None:
            return parsed
    try:
        return datetime.fromtimestamp(path.stat().st_mtime).date()
    except OSError:
        return datetime.now().date()


def _infer_start_date_from_events(path: Path, events: list[TimedBatteryEvent]) -> date:
    for event in events:
        if event.date_value is not None:
            return event.date_value
    return _infer_start_date_from_path(path)


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


def _build_single_log_rows_from_timed_events(
    text_path: Path,
    events: list[TimedBatteryEvent],
) -> list[IndicatorLogRow]:
    if not events:
        raise AppError("TEXT_NO_EVENT", "文本中未提取到可用电量数据", detail=text_path.name)

    current_date = _infer_start_date_from_events(text_path, events)
    previous_time: dt_time | None = None
    transition_rows: list[IndicatorLogRow] = []
    for event in events:
        if event.date_value is not None:
            current_date = event.date_value
        elif previous_time is not None and event.time_value < previous_time:
            current_date += timedelta(days=1)

        log_datetime = datetime.combine(current_date, event.time_value)
        transition_rows.append(
            IndicatorLogRow(
                log_datetime=log_datetime,
                level=event.level,
                mapped_voltage_v=None,
                log_date=current_date.isoformat(),
            )
        )
        previous_time = event.time_value
    return _expand_rows_by_second(transition_rows)


def _build_single_log_rows_from_special_points(
    text_path: Path,
    special_points: list[SpecialBatteryPoint],
) -> list[IndicatorLogRow]:
    if not special_points:
        raise AppError("TEXT_NO_EVENT", "文本中未提取到可用电量数据", detail=text_path.name)

    # 特殊格式无稳定时间戳：按旧规则从 0:00:00 起按 1 秒步进填充。
    start_datetime = datetime.combine(_infer_start_date_from_path(text_path), dt_time(0, 0, 0))
    raw_rows: list[IndicatorLogRow] = []
    for index, point in enumerate(special_points):
        log_datetime = start_datetime + timedelta(seconds=index)
        raw_rows.append(
            IndicatorLogRow(
                log_datetime=log_datetime,
                level=point.level,
                mapped_voltage_v=point.voltage_v,
                log_date=log_datetime.date().isoformat(),
            )
        )
    return raw_rows


def _expand_rows_by_second(rows: list[IndicatorLogRow]) -> list[IndicatorLogRow]:
    if not rows:
        return rows
    expanded_rows: list[IndicatorLogRow] = [rows[0]]
    previous = rows[0]
    for current in rows[1:]:
        delta_seconds = int((current.log_datetime - previous.log_datetime).total_seconds())
        if delta_seconds < 0:
            raise AppError(
                "TEXT_TIME_REVERSED",
                "日志时间顺序异常，存在倒序时间点，请检查文本数据",
            )
        for offset in range(1, delta_seconds):
            intermediate_dt = previous.log_datetime + timedelta(seconds=offset)
            expanded_rows.append(
                IndicatorLogRow(
                    log_datetime=intermediate_dt,
                    level=previous.level,
                    mapped_voltage_v=previous.mapped_voltage_v,
                    log_date=intermediate_dt.date().isoformat(),
                )
            )
        expanded_rows.append(current)
        previous = current
    return expanded_rows


def compute_single_log_endurance_duration(log_rows: list[IndicatorLogRow]) -> timedelta:
    if not log_rows:
        raise AppError("TEXT_NO_EVENT", "文本中未提取到可用电量数据")
    return log_rows[-1].log_datetime - log_rows[0].log_datetime


def compute_indicator_endurance_duration(
    dataset: ChargeDataset,
    log_rows: list[IndicatorLogRow],
) -> timedelta:
    if not log_rows:
        raise AppError("TEXT_NO_EVENT", "文本中未提取到可用电量数据", detail=dataset.source_path.name)

    start_time = dataset.datetimes[0]
    end_row = log_rows[-1]
    end_time = end_row.log_datetime
    if end_row.log_date:
        try:
            end_date = datetime.strptime(end_row.log_date, "%Y-%m-%d").date()
            end_time = datetime.combine(end_date, end_row.log_datetime.time())
        except ValueError:
            pass
    return end_time - start_time


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


def process_endurance_single_log_statistics(
    inputs: list[Path],
    output_dir: Path,
    logger: Logger = None,
    chunk_size: int | None = None,
    wait_seconds: float = 0.0,
) -> BatchResult:
    output_dir.mkdir(parents=True, exist_ok=True)
    files, collection_warnings = collect_files(inputs, {".xlsx", ".xls", ".txt", ".log"})
    for warning in collection_warnings:
        _emit(logger, "WARN", warning)
    text_files = [file for file in files if file.suffix.lower() in {".txt", ".log"}]
    excel_files = [file for file in files if file.suffix.lower() in {".xlsx", ".xls"}]

    if excel_files:
        error_message = (
            f"单log续航时长统计仅支持文本文件（.txt/.log），检测到 {len(excel_files)} 个 Excel 文件"
        )
        results = [
            ProcessItemResult(name=excel_file.name, status="failed", error=error_message)
            for excel_file in excel_files
        ]
        for result in results:
            _emit(logger, "ERROR", f"[失败] {result.name} 处理失败：{error_message}")
        return BatchResult(
            total=len(results),
            success=0,
            failed=len(results),
            items=results,
            warnings=[error_message],
        )

    if not text_files:
        warning_message = "未检测到可处理的文本文件（.txt/.log）"
        _emit(logger, "WARN", warning_message)
        return BatchResult(total=0, success=0, failed=0, items=[], warnings=[warning_message])

    results: list[ProcessItemResult] = []
    total_files = len(text_files)
    use_pacing = chunk_size is not None and chunk_size > 0 and wait_seconds > 0
    effective_chunk_size = chunk_size if use_pacing else 0
    for index, text_file in enumerate(text_files, start=1):
        _emit(logger, "INFO", f"开始处理（单log续航时长统计）：{text_file}")
        try:
            log_parse_result = parse_battery_log_file(text_file)
            if log_parse_result.mode == "timed":
                log_rows = _build_single_log_rows_from_timed_events(text_file, log_parse_result.timed_events)
                include_voltage = False
            else:
                log_rows = _build_single_log_rows_from_special_points(text_file, log_parse_result.special_points)
                include_voltage = True

            endurance_duration = compute_single_log_endurance_duration(log_rows)
            output_path = resolve_output_path(output_dir, text_file.stem)
            render_endurance_single_log_workbook(
                file_stem=text_file.stem,
                log_rows=log_rows,
                include_voltage=include_voltage,
                endurance_duration=endurance_duration,
                output_path=output_path,
            )
            result = ProcessItemResult(
                name=text_file.name,
                status="success",
                output_path=output_path,
                warnings=log_parse_result.warnings.copy(),
            )
            results.append(result)
            _emit(logger, "INFO", f"[成功] 输出成功：{output_path}")
            for warning in log_parse_result.warnings:
                _emit(logger, "WARN", f"{text_file.name}：{warning}")
        except AppError as exc:
            result = ProcessItemResult(name=text_file.name, status="failed", error=str(exc))
            results.append(result)
            _emit(logger, "ERROR", f"[失败] {text_file.name} 处理失败：{exc}")
        except Exception as exc:  # pragma: no cover
            result = ProcessItemResult(name=text_file.name, status="failed", error=f"[UNEXPECTED] {exc}")
            results.append(result)
            _emit(logger, "ERROR", f"[失败] {text_file.name} 处理失败：[UNEXPECTED] {exc}")

        if use_pacing and index % effective_chunk_size == 0 and index < total_files:
            remaining = total_files - index
            _emit(
                logger,
                "INFO",
                f"已处理 {index}/{total_files} 个文本文件，等待 {_format_wait_seconds(wait_seconds)} 秒后继续（剩余 {remaining} 个）",
            )
            time.sleep(wait_seconds)

    success_count = sum(1 for item in results if item.status == "success")
    return BatchResult(
        total=len(results),
        success=success_count,
        failed=len(results) - success_count,
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
            if log_parse_result.mode != "timed":
                raise AppError(
                    "TEXT_SPECIAL_MOVED_TO_SINGLE_LOG",
                    "电量指示统计不再支持该特殊文本格式，请改用“执行单log续航时长统计”",
                    detail=group.text_path.name,
                )
            log_rows = _build_indicator_rows_from_timed_events(dataset, log_parse_result.timed_events)

            endurance_duration = compute_indicator_endurance_duration(
                dataset=dataset,
                log_rows=log_rows,
            )
            output_path = resolve_output_path(output_dir, group.stem)
            render_endurance_indicator_workbook(
                dataset=dataset,
                log_rows=log_rows,
                include_log_date=True,
                endurance_duration=endurance_duration,
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
