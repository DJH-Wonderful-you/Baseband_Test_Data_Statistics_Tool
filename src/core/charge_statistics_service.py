from __future__ import annotations

from pathlib import Path
from typing import Callable

from .charge_parser import parse_charge_workbook
from .errors import AppError
from .excel_render import render_charge_workbook
from .file_collect import collect_statistics_excel_files
from .models import BatchResult, ChargeDataset, ChargeMetrics, ProcessItemResult, TemperatureMetrics


Logger = Callable[[str, str], None] | None


def _emit(logger: Logger, level: str, message: str) -> None:
    if logger is not None:
        logger(level, message)


def resolve_output_path(output_dir: Path, stem: str) -> Path:
    candidate = output_dir / f"{stem}.xlsx"
    if not candidate.exists():
        return candidate
    suffix = 1
    while True:
        candidate = output_dir / f"{stem}({suffix}).xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def apply_tail_fill_check(dataset: ChargeDataset) -> None:
    if dataset.row_count() < 2:
        raise AppError("ROW_NOT_ENOUGH", "数据行不足，无法进行尾行校验", detail=dataset.source_path.name)
    last_idx = dataset.row_count() - 1
    prev_idx = last_idx - 1

    current_last = dataset.currents_ma[last_idx]
    current_prev = dataset.currents_ma[prev_idx]
    voltage_last = dataset.voltages_v[last_idx]
    voltage_prev = dataset.voltages_v[prev_idx]

    if current_last is None or voltage_last is None:
        if current_prev is None or voltage_prev is None:
            raise AppError(
                "TAIL_DATA_INVALID",
                "末尾倒数第二个时间点的电流或电压为空，请检查数据",
                detail=dataset.source_path.name,
            )
        if current_last is None:
            dataset.currents_ma[last_idx] = current_prev
            dataset.warnings.append("最后一个时间点电流为空，已使用倒数第二个数据补齐")
        if voltage_last is None:
            dataset.voltages_v[last_idx] = voltage_prev
            dataset.warnings.append("最后一个时间点电压为空，已使用倒数第二个数据补齐")


def _within_ten_percent(value: float, previous: float) -> bool:
    if previous == 0:
        return False
    return abs(value - previous) <= abs(previous) * 0.1


def compute_charge_metrics(dataset: ChargeDataset) -> ChargeMetrics:
    currents = dataset.currents_ma
    voltages = dataset.voltages_v
    times = dataset.datetimes
    row_count = dataset.row_count()

    precharge: float | None = None
    cutoff: float | None = None
    cutoff_index: int | None = None
    for idx in range(1, row_count - 1):
        value = currents[idx]
        previous = currents[idx - 1]
        following = currents[idx + 1]
        if value is None or previous is None or following is None:
            continue

        common = value > 1 and _within_ten_percent(value, previous)
        if precharge is None and common and following >= value * 1.5:
            precharge = value

        cutoff_condition = common and (value >= following * 1.5 or following < 0)
        if cutoff_condition:
            cutoff = value
            cutoff_index = idx

    valid_currents = [value for value in currents if value is not None]
    const_current = max(valid_currents) if valid_currents else None

    valid_voltages = [value for value in voltages if value is not None]
    full_voltage = max(valid_voltages) if valid_voltages else None

    duration = None
    if const_current is not None and cutoff_index is not None:
        start_index = next(
            (
                idx
                for idx, value in enumerate(currents)
                if value is not None and value >= const_current * 0.9
            ),
            None,
        )
        if start_index is not None:
            duration = times[cutoff_index] - times[start_index]

    return ChargeMetrics(
        precharge_current_ma=precharge,
        const_current_ma=const_current,
        cutoff_current_ma=cutoff,
        full_voltage_v=full_voltage,
        duration=duration,
    )


def compute_temperature_metrics(dataset: ChargeDataset) -> TemperatureMetrics | None:
    if not dataset.has_temperature_data:
        return None
    indexed_pen = [
        (idx, value) for idx, value in enumerate(dataset.pen_temps_c) if value is not None
    ]
    if not indexed_pen:
        return TemperatureMetrics(None, None, None)
    max_index, max_pen = max(indexed_pen, key=lambda item: item[1])
    env_at_max = dataset.env_temps_c[max_index]
    hotspot = (max_pen - env_at_max) if env_at_max is not None else None
    return TemperatureMetrics(
        max_pen_temp_c=max_pen,
        env_temp_at_max_pen_c=env_at_max,
        hotspot_rise_c=hotspot,
    )


def process_charge_statistics(
    inputs: list[Path], output_dir: Path, logger: Logger = None
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
    for excel_file in excel_files:
        _emit(logger, "INFO", f"开始处理（统计数据）：{excel_file}")
        try:
            dataset = parse_charge_workbook(excel_file, require_voltage=True)
            apply_tail_fill_check(dataset)
            metrics = compute_charge_metrics(dataset)
            temp_metrics = compute_temperature_metrics(dataset)
            output_path = resolve_output_path(output_dir, dataset.stem)
            render_charge_workbook(dataset, metrics, temp_metrics, output_path)
            results.append(
                ProcessItemResult(
                    name=excel_file.name,
                    status="success",
                    output_path=output_path,
                    warnings=dataset.warnings.copy(),
                )
            )
            _emit(logger, "INFO", f"输出成功：{output_path}")
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
            _emit(logger, "ERROR", f"{excel_file.name} 处理失败：{exc}")
        except Exception as exc:  # pragma: no cover
            results.append(
                ProcessItemResult(
                    name=excel_file.name,
                    status="failed",
                    error=f"[UNEXPECTED] {exc}",
                )
            )
            _emit(logger, "ERROR", f"{excel_file.name} 处理失败：[UNEXPECTED] {exc}")

    success_count = sum(1 for item in results if item.status == "success")
    failed_count = len(results) - success_count
    return BatchResult(
        total=len(results),
        success=success_count,
        failed=failed_count,
        items=results,
    )
