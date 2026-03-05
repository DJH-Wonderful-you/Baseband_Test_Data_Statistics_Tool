from __future__ import annotations

from dataclasses import dataclass, field
from datetime import time
from pathlib import Path
import re

from .errors import AppError


_TIMED_LEVEL_PATTERN = re.compile(
    r"\[(?P<hms>\d{1,2}:\d{2}:\d{2})\.\d+\][^\r\n]*?level:(?P<level>-?\d+)\(%\)",
    re.IGNORECASE,
)
_TAB_LEVEL_PATTERN = re.compile(
    r"(?:^|[\t ])(?P<hms>\d{1,2}:\d{2}:\d{2})\.\d+[\t ].*?\"(?P<level>-?\d+)%\" received",
    re.IGNORECASE,
)
_L_FIELD_PATTERN = re.compile(
    r"\[(?P<hms>\d{1,2}:\d{2}:\d{2})\.\d+\][^\r\n]*?\bL=(?P<level>-?\d+)",
    re.IGNORECASE,
)
_SPECIAL_PATTERN = re.compile(
    r"vol:(?P<mv>-?\d+)\(mv\)\s+level:(?P<level>-?\d+)\(%\)",
    re.IGNORECASE,
)


@dataclass(slots=True)
class TimedBatteryEvent:
    time_value: time
    level: int


@dataclass(slots=True)
class SpecialBatteryPoint:
    level: int
    voltage_v: float


@dataclass(slots=True)
class BatteryLogParseResult:
    mode: str
    timed_events: list[TimedBatteryEvent] = field(default_factory=list)
    special_points: list[SpecialBatteryPoint] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def _parse_hms(hms_text: str) -> time:
    parts = hms_text.split(":")
    if len(parts) != 3:
        raise ValueError(hms_text)
    hour, minute, second = (int(part) for part in parts)
    return time(hour=hour, minute=minute, second=second)


def _timed_event_from_line(line: str) -> TimedBatteryEvent | None:
    for pattern in (_TIMED_LEVEL_PATTERN, _TAB_LEVEL_PATTERN, _L_FIELD_PATTERN):
        match = pattern.search(line)
        if match is None:
            continue
        hms_text = match.group("hms")
        level = int(match.group("level"))
        try:
            return TimedBatteryEvent(time_value=_parse_hms(hms_text), level=level)
        except ValueError:
            continue
    return None


def _special_point_from_line(line: str) -> SpecialBatteryPoint | None:
    match = _SPECIAL_PATTERN.search(line)
    if match is None:
        return None
    mv_value = int(match.group("mv"))
    level = int(match.group("level"))
    return SpecialBatteryPoint(level=level, voltage_v=mv_value / 1000.0)


def _score_decoded_text(text: str) -> int:
    return sum(
        len(pattern.findall(text))
        for pattern in (_TIMED_LEVEL_PATTERN, _TAB_LEVEL_PATTERN, _L_FIELD_PATTERN, _SPECIAL_PATTERN)
    )


def _read_text_file(path: Path) -> tuple[str, str]:
    raw = path.read_bytes()
    if not raw:
        return "", "utf-8-sig"

    candidates = [
        "utf-8-sig",
        "utf-8",
        "gb18030",
        "utf-16",
        "utf-16-le",
        "utf-16-be",
    ]

    best_text = ""
    best_encoding = candidates[0]
    best_score = -1
    for encoding in candidates:
        try:
            decoded = raw.decode(encoding)
        except UnicodeDecodeError:
            continue
        score = _score_decoded_text(decoded)
        if score > best_score:
            best_score = score
            best_text = decoded
            best_encoding = encoding
            if score > 0 and encoding in {"utf-8-sig", "utf-8", "gb18030"}:
                break

    if best_score >= 0:
        return best_text, best_encoding

    return raw.decode("utf-8", errors="ignore"), "utf-8"


def _collapse_repeated_levels_timed(events: list[TimedBatteryEvent]) -> tuple[list[TimedBatteryEvent], int]:
    collapsed: list[TimedBatteryEvent] = []
    skipped = 0
    for event in events:
        if collapsed and collapsed[-1].level == event.level:
            skipped += 1
            continue
        collapsed.append(event)
    return collapsed, skipped


def _collapse_repeated_levels_special(
    points: list[SpecialBatteryPoint],
) -> tuple[list[SpecialBatteryPoint], int]:
    collapsed: list[SpecialBatteryPoint] = []
    skipped = 0
    for point in points:
        if collapsed and collapsed[-1].level == point.level:
            skipped += 1
            continue
        collapsed.append(point)
    return collapsed, skipped


def _build_step_warning(level_values: list[int]) -> str | None:
    if len(level_values) <= 1:
        return None
    mismatches: list[tuple[int, int]] = []
    for prev, curr in zip(level_values, level_values[1:]):
        if curr != prev - 1:
            mismatches.append((prev, curr))
    if not mismatches:
        return None
    preview = "，".join(f"{prev}->{curr}" for prev, curr in mismatches[:5])
    if len(mismatches) > 5:
        preview += f"（共 {len(mismatches)} 处）"
    return f"电量序列未严格按 1 递减（示例：{preview}），请检查原始数据。"


def parse_battery_log_file(path: Path) -> BatteryLogParseResult:
    text, encoding = _read_text_file(path)
    if not text.strip():
        raise AppError("TEXT_EMPTY", "文本文件为空", detail=path.name)

    timed_events_raw: list[TimedBatteryEvent] = []
    special_points_raw: list[SpecialBatteryPoint] = []
    for line in text.splitlines():
        if not line.strip():
            continue
        timed_event = _timed_event_from_line(line)
        if timed_event is not None:
            timed_events_raw.append(timed_event)
            continue
        special_point = _special_point_from_line(line)
        if special_point is not None:
            special_points_raw.append(special_point)

    warnings: list[str] = []
    if encoding.lower() not in {"utf-8", "utf-8-sig"}:
        warnings.append(f"文本文件按 {encoding} 编码读取")

    if timed_events_raw:
        timed_events, skipped = _collapse_repeated_levels_timed(timed_events_raw)
        if skipped > 0:
            warnings.append(f"文本中存在连续重复电量记录 {skipped} 条，已仅保留首次出现值")
        step_warning = _build_step_warning([event.level for event in timed_events])
        if step_warning is not None:
            warnings.append(step_warning)
        return BatteryLogParseResult(mode="timed", timed_events=timed_events, warnings=warnings)

    if special_points_raw:
        # 无时间戳特殊格式按 1 秒采样，需保留全部电量与电压记录，不做“首次电量去重”。
        transition_points, _ = _collapse_repeated_levels_special(special_points_raw)
        step_warning = _build_step_warning([point.level for point in transition_points])
        if step_warning is not None:
            warnings.append(step_warning)
        return BatteryLogParseResult(mode="special", special_points=special_points_raw, warnings=warnings)

    raise AppError("TEXT_PARSE_FAILED", "文本文件未识别到可用的电量日志格式", detail=path.name)
