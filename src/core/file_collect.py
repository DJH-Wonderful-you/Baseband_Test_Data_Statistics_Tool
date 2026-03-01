from __future__ import annotations

from collections import defaultdict
from pathlib import Path

from .models import MergeInputGroup


def _iter_files_from_path(path: Path, allowed_extensions: set[str]) -> list[Path]:
    if not path.exists():
        return []
    if path.is_file():
        return [path] if path.suffix.lower() in allowed_extensions else []
    files: list[Path] = []
    for candidate in path.rglob("*"):
        if candidate.is_file() and candidate.suffix.lower() in allowed_extensions:
            files.append(candidate)
    return files


def collect_files(
    input_paths: list[Path], allowed_extensions: set[str]
) -> tuple[list[Path], list[str]]:
    warnings: list[str] = []
    deduplicated: dict[Path, None] = {}
    for raw_path in input_paths:
        path = raw_path.expanduser()
        if not path.exists():
            warnings.append(f"路径不存在，已忽略：{path}")
            continue
        for file_path in _iter_files_from_path(path, allowed_extensions):
            deduplicated[file_path.resolve()] = None
    return sorted(deduplicated.keys()), warnings


def collect_statistics_excel_files(input_paths: list[Path]) -> tuple[list[Path], list[str]]:
    return collect_files(input_paths, {".xlsx", ".xls"})


def collect_merge_groups(
    input_paths: list[Path],
) -> tuple[list[MergeInputGroup], list[str], list[str]]:
    files, warnings = collect_files(input_paths, {".xlsx", ".xls", ".csv"})
    excel_by_stem: defaultdict[str, list[Path]] = defaultdict(list)
    csv_by_stem: defaultdict[str, list[Path]] = defaultdict(list)
    for file_path in files:
        stem = file_path.stem
        if file_path.suffix.lower() in {".xlsx", ".xls"}:
            excel_by_stem[stem].append(file_path)
        elif file_path.suffix.lower() == ".csv":
            csv_by_stem[stem].append(file_path)

    groups: list[MergeInputGroup] = []
    errors: list[str] = []
    all_stems = sorted(set(excel_by_stem.keys()) | set(csv_by_stem.keys()))
    for stem in all_stems:
        excel_candidates = excel_by_stem.get(stem, [])
        csv_candidates = csv_by_stem.get(stem, [])
        if len(excel_candidates) == 1 and len(csv_candidates) == 1:
            groups.append(
                MergeInputGroup(
                    stem=stem,
                    excel_path=excel_candidates[0],
                    csv_path=csv_candidates[0],
                )
            )
            continue
        if not excel_candidates:
            errors.append(f"[{stem}] 缺少匹配的 Excel 文件（.xlsx/.xls）")
            continue
        if not csv_candidates:
            errors.append(f"[{stem}] 缺少匹配的 .csv 文件")
            continue
        errors.append(
            f"[{stem}] 检测到多个同名文件（excel={len(excel_candidates)}, csv={len(csv_candidates)}），无法唯一配对"
        )
    return groups, errors, warnings
