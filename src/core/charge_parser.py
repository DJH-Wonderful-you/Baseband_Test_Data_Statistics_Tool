from __future__ import annotations

import csv
import re
from datetime import date, datetime, time
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

from .errors import AppError
from .models import ChargeDataset


def _normalize_header(text: str) -> str:
    normalized = text.strip().lower()
    normalized = normalized.replace("（", "(").replace("）", ")")
    normalized = normalized.replace(" ", "")
    return normalized


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
    if _is_empty(value):
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = _safe_str(value).replace(",", "")
    match = re.search(r"[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?", text)
    if not match:
        return None
    return float(match.group(0))


def _parse_date_only(value: Any, epoch: datetime) -> date | None:
    if _is_empty(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        converted = from_excel(value, epoch=epoch)
        if isinstance(converted, datetime):
            return converted.date()
        if isinstance(converted, date):
            return converted
        return None
    text = _safe_str(value)
    match = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", text)
    if not match:
        return None
    year, month, day = map(int, match.groups())
    return date(year, month, day)


def _parse_datetime_with_truncation(
    value: Any, epoch: datetime, fallback_date: date | None = None
) -> datetime | None:
    if _is_empty(value):
        return None
    if isinstance(value, datetime):
        return value.replace(microsecond=0)
    if isinstance(value, date):
        return datetime.combine(value, time.min)
    if isinstance(value, (int, float)):
        converted = from_excel(value, epoch=epoch)
        if isinstance(converted, datetime):
            return converted.replace(microsecond=0)
        if isinstance(converted, date):
            return datetime.combine(converted, time.min)
        if isinstance(converted, time):
            if fallback_date is None:
                return None
            return datetime.combine(fallback_date, converted).replace(microsecond=0)
        return None

    text = _safe_str(value)
    full_match = re.search(
        r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})\D+(\d{1,2})[:：](\d{1,2})[:：](\d{1,2})",
        text,
    )
    if full_match:
        year, month, day, hour, minute, second = map(int, full_match.groups())
        return datetime(year, month, day, hour, minute, second)

    time_match = re.search(r"(\d{1,2})[:：](\d{1,2})[:：](\d{1,2})", text)
    if time_match and fallback_date is not None:
        hour, minute, second = map(int, time_match.groups())
        return datetime.combine(fallback_date, time(hour, minute, second))
    return None


def _find_header_row(ws: Any) -> int:
    keyword_set = ("时间", "电流", "电压", "索引", "笔壳", "环境")
    best_row = 1
    best_score = -1
    for row_idx in range(1, min(ws.max_row, 30) + 1):
        score = 0
        non_empty = 0
        for col_idx in range(1, ws.max_column + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if _is_empty(value):
                continue
            non_empty += 1
            text = _safe_str(value)
            if any(keyword in text for keyword in keyword_set):
                score += 1
        if non_empty == 0:
            continue
        if score > best_score:
            best_score = score
            best_row = row_idx
    return best_row


def _read_headers(ws: Any, header_row: int) -> dict[int, str]:
    headers: dict[int, str] = {}
    for col_idx in range(1, ws.max_column + 1):
        value = ws.cell(row=header_row, column=col_idx).value
        if _is_empty(value):
            continue
        headers[col_idx] = _safe_str(value)
    return headers


def _find_column(headers: dict[int, str], predicate: Any) -> int | None:
    for col_idx, text in headers.items():
        if predicate(text):
            return col_idx
    return None


def _build_extra_headers(headers: dict[int, str], excluded_columns: set[int]) -> list[tuple[int, str]]:
    extras: list[tuple[int, str]] = []
    existing: set[str] = set()
    for col_idx, text in headers.items():
        if col_idx in excluded_columns:
            continue
        candidate = text
        if candidate in existing:
            candidate = f"{candidate}_{col_idx}"
        existing.add(candidate)
        extras.append((col_idx, candidate))
    return extras


def parse_charge_workbook(path: Path, require_voltage: bool) -> ChargeDataset:
    workbook = load_workbook(path, data_only=True)
    sheet = workbook[workbook.sheetnames[0]]
    header_row = _find_header_row(sheet)
    headers = _read_headers(sheet, header_row)
    if not headers:
        raise AppError("EMPTY_HEADER", "未检测到有效表头", detail=str(path))

    time_col = _find_column(
        headers,
        lambda h: "时间" in h and "log" not in _normalize_header(h),
    )
    current_col = _find_column(
        headers,
        lambda h: "电流" in h and "log" not in _normalize_header(h),
    )
    voltage_col = _find_column(
        headers,
        lambda h: ("电压" in h or "vbat" in _normalize_header(h))
        and "映射" not in h
        and "log" not in _normalize_header(h),
    )
    index_col = _find_column(headers, lambda h: "索引" in h)
    date_col = _find_column(headers, lambda h: "日期" in h and "log" not in _normalize_header(h))
    pen_col = _find_column(headers, lambda h: "笔壳" in h)
    env_col = _find_column(headers, lambda h: "环境" in h)
    parser_warnings: list[str] = []
    if pen_col is not None and env_col is None:
        temperature_candidates = [
            col_idx
            for col_idx, header_text in headers.items()
            if "温度" in header_text and col_idx != pen_col
        ]
        if temperature_candidates:
            env_col = temperature_candidates[0]
            parser_warnings.append("未直接匹配到“环境”关键词，已使用另一列温度数据作为环境温度")

    if time_col is None:
        raise AppError("MISSING_TIME", "未检测到“时间 (s)”相关列", detail=str(path))
    if current_col is None:
        raise AppError("MISSING_CURRENT", "未检测到“电流”相关列", detail=str(path))
    if require_voltage and voltage_col is None:
        raise AppError("MISSING_VOLTAGE", "未检测到“电压”相关列", detail=str(path))

    excluded_columns = {time_col, current_col}
    if index_col is not None:
        excluded_columns.add(index_col)
    if date_col is not None:
        excluded_columns.add(date_col)
    if voltage_col is not None:
        excluded_columns.add(voltage_col)
    if pen_col is not None:
        excluded_columns.add(pen_col)
    if env_col is not None:
        excluded_columns.add(env_col)
    extra_headers = _build_extra_headers(headers, excluded_columns)

    index_values: list[int | None] = []
    datetimes: list[datetime] = []
    date_strings: list[str] = []
    time_strings: list[str] = []
    currents_raw: list[float | None] = []
    voltages_v: list[float | None] = []
    pen_temps_c: list[float | None] = []
    env_temps_c: list[float | None] = []
    extras: dict[str, list[Any]] = {header: [] for _, header in extra_headers}
    warnings: list[str] = parser_warnings.copy()

    for row_idx in range(header_row + 1, sheet.max_row + 1):
        time_value = sheet.cell(row=row_idx, column=time_col).value
        if _is_empty(time_value):
            continue
        fallback_date = None
        if date_col is not None:
            date_value = sheet.cell(row=row_idx, column=date_col).value
            fallback_date = _parse_date_only(date_value, workbook.epoch)
        dt_value = _parse_datetime_with_truncation(time_value, workbook.epoch, fallback_date)
        if dt_value is None:
            raise AppError(
                "INVALID_DATETIME",
                "时间列存在无法解析的数据",
                detail=f"{path.name} row={row_idx} value={time_value}",
            )

        raw_index_value = sheet.cell(row=row_idx, column=index_col).value if index_col else None
        parsed_index = int(raw_index_value) if isinstance(raw_index_value, int) else _to_float(raw_index_value)
        index_values.append(int(parsed_index) if parsed_index is not None else None)

        datetimes.append(dt_value)
        date_strings.append(dt_value.strftime("%Y-%m-%d"))
        time_strings.append(dt_value.strftime("%H:%M:%S"))
        currents_raw.append(_to_float(sheet.cell(row=row_idx, column=current_col).value))
        if voltage_col is not None:
            voltages_v.append(_to_float(sheet.cell(row=row_idx, column=voltage_col).value))
        else:
            voltages_v.append(None)
        pen_temps_c.append(_to_float(sheet.cell(row=row_idx, column=pen_col).value) if pen_col else None)
        env_temps_c.append(_to_float(sheet.cell(row=row_idx, column=env_col).value) if env_col else None)
        for extra_col, extra_name in extra_headers:
            extras[extra_name].append(sheet.cell(row=row_idx, column=extra_col).value)

    if not datetimes:
        raise AppError("NO_DATA", "未检测到有效数据行", detail=str(path))

    current_header_text = headers[current_col]
    current_header_norm = _normalize_header(current_header_text)
    current_factor = 1000.0
    if "ma" in current_header_norm:
        current_factor = 1.0
    elif "(a)" in current_header_norm or "电流(a)" in current_header_norm:
        current_factor = 1000.0

    currents_ma = [
        (value * current_factor if value is not None else None) for value in currents_raw
    ]

    if current_factor == 1.0:
        valid_abs = [abs(v) for v in currents_ma if v is not None]
        if valid_abs and 0 < max(valid_abs) <= 5:
            currents_ma = [(v * 1000 if v is not None else None) for v in currents_ma]
            warnings.append(
                "检测到电流列标注为 mA 但量级偏小，已按 A->mA 进行自动换算（乘以1000）"
            )

    valid_current_values = [value for value in currents_ma if value is not None]
    if valid_current_values and max(valid_current_values) < 0:
        currents_ma = [(-value if value is not None else None) for value in currents_ma]

    has_temperature_data = pen_col is not None and env_col is not None
    return ChargeDataset(
        source_path=path,
        stem=path.stem,
        index_values=index_values,
        datetimes=datetimes,
        date_strings=date_strings,
        time_strings=time_strings,
        currents_ma=currents_ma,
        voltages_v=voltages_v,
        pen_temps_c=pen_temps_c,
        env_temps_c=env_temps_c,
        extras=extras,
        extra_headers_order=[header for _, header in extra_headers],
        has_temperature_data=has_temperature_data,
        warnings=warnings,
    )


def parse_voltage_csv(path: Path) -> tuple[dict[datetime, float], list[str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        if not reader.fieldnames:
            raise AppError("CSV_EMPTY", "CSV 文件为空或缺少表头", detail=str(path))

        field_map: dict[str, str] = {}
        for field in reader.fieldnames:
            normalized = _normalize_header(field)
            field_map[field] = normalized

        datetime_field = next(
            (
                field
                for field, normalized in field_map.items()
                if "date/time" in normalized or ("date" in normalized and "time" in normalized)
            ),
            None,
        )
        value_field = next(
            (
                field
                for field, normalized in field_map.items()
                if normalized == "value" or normalized.endswith("value")
            ),
            None,
        )
        if datetime_field is None or value_field is None:
            raise AppError(
                "CSV_HEADER_INVALID",
                "CSV 文件缺少 Date/Time 或 Value 列",
                detail=str(path),
            )

        mapping: dict[datetime, float] = {}
        duplicate_counter = 0
        for row_index, row in enumerate(reader, start=2):
            dt_raw = _safe_str(row.get(datetime_field))
            value_raw = row.get(value_field)
            if not dt_raw:
                continue
            try:
                dt_value = datetime.strptime(dt_raw, "%Y-%m-%d %H:%M:%S")
            except ValueError:
                match = re.search(
                    r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})\D+(\d{1,2})[:：](\d{1,2})[:：](\d{1,2})",
                    dt_raw,
                )
                if not match:
                    raise AppError(
                        "CSV_DATETIME_INVALID",
                        "CSV 时间格式无法解析",
                        detail=f"{path.name} row={row_index} value={dt_raw}",
                    )
                dt_value = datetime(*map(int, match.groups()))

            numeric_value = _to_float(value_raw)
            if numeric_value is None:
                raise AppError(
                    "CSV_VALUE_INVALID",
                    "CSV Value 列存在无法解析的数据",
                    detail=f"{path.name} row={row_index} value={value_raw}",
                )
            if dt_value in mapping:
                duplicate_counter += 1
                continue
            mapping[dt_value] = numeric_value

    if not mapping:
        raise AppError("CSV_NO_DATA", "CSV 文件没有可用数据", detail=str(path))
    warnings: list[str] = []
    if duplicate_counter > 0:
        warnings.append(f"CSV 存在 {duplicate_counter} 个重复秒级时间点，已保留首次出现的数据")
    return mapping, warnings
