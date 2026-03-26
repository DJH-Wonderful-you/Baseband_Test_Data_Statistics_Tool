from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any


@dataclass(slots=True)
class ProcessItemResult:
    name: str
    status: str
    output_path: Path | None = None
    error: str | None = None
    warnings: list[str] = field(default_factory=list)


@dataclass(slots=True)
class BatchResult:
    total: int
    success: int
    failed: int
    items: list[ProcessItemResult]
    warnings: list[str] = field(default_factory=list)


@dataclass(slots=True)
class MergeInputGroup:
    stem: str
    excel_path: Path
    csv_path: Path


@dataclass(slots=True)
class EnduranceInputGroup:
    stem: str
    excel_path: Path
    text_path: Path


@dataclass(slots=True)
class ChargeMetrics:
    precharge_current_ma: float | None
    const_current_ma: float | None
    cutoff_current_ma: float | None
    full_voltage_v: float | None
    duration: timedelta | None


@dataclass(slots=True)
class TemperatureMetrics:
    max_pen_temp_c: float | None
    env_temp_at_max_pen_c: float | None
    hotspot_rise_c: float | None


@dataclass(slots=True)
class ChargeDataset:
    source_path: Path
    stem: str
    index_values: list[int | None]
    datetimes: list[datetime]
    date_strings: list[str]
    time_strings: list[str]
    currents_ma: list[float | None]
    voltages_v: list[float | None]
    pen_temps_c: list[float | None]
    env_temps_c: list[float | None]
    extras: dict[str, list[Any]]
    extra_headers_order: list[str]
    has_temperature_data: bool
    warnings: list[str] = field(default_factory=list)

    def row_count(self) -> int:
        return len(self.datetimes)

    def has_current_data(self) -> bool:
        return any(value is not None for value in self.currents_ma)

    def has_voltage_data(self) -> bool:
        return any(value is not None for value in self.voltages_v)

    def has_charge_curve_data(self) -> bool:
        return self.has_current_data() and self.has_voltage_data()

    def has_temperature_measurements(self) -> bool:
        return (
            self.has_temperature_data
            and any(value is not None for value in self.pen_temps_c)
            and any(value is not None for value in self.env_temps_c)
        )


@dataclass(slots=True)
class ChargeProtectionDataset:
    source_path: Path
    stem: str
    index_values: list[int | None]
    datetimes: list[datetime]
    date_strings: list[str]
    time_strings: list[str]
    currents_ma: list[float | None]
    voltages_v: list[float | None]
    env_temps_c: list[float | None]
    cell_temps_c: list[float | None]
    warnings: list[str] = field(default_factory=list)

    def row_count(self) -> int:
        return len(self.datetimes)


@dataclass(slots=True)
class ChargeProtectionMetrics:
    mode: str
    high_protect_temp_c: float | None
    high_resume_temp_c: float | None
    low_protect_temp_c: float | None
    low_resume_temp_c: float | None
