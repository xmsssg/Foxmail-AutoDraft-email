from __future__ import annotations

from datetime import datetime
from fnmatch import fnmatch
from pathlib import Path
from typing import Iterable


EXCEL_EXTENSIONS = {".xlsx", ".xls"}


def is_candidate_excel(path: Path) -> bool:
    if not path.is_file():
        return False
    if path.name.startswith("~$"):
        return False
    return path.suffix.lower() in EXCEL_EXTENSIONS


def matches_include_patterns(path: Path, customer: dict) -> bool:
    patterns = customer.get("include_patterns") or ["*.xlsx", "*.xls"]
    return any(fnmatch(path.name, pattern) for pattern in patterns)


def parse_start_mtime(value: str | None) -> float | None:
    if not value:
        return None
    text = value.strip()
    if not text:
        return None
    formats = (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%Y/%m/%d",
    )
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).timestamp()
        except ValueError:
            continue
    raise ValueError(f"起始时间格式不正确: {value}")


def matches_start_mtime(path: Path, customer: dict) -> bool:
    start_timestamp = parse_start_mtime(customer.get("start_mtime"))
    if start_timestamp is None:
        return True
    return path.stat().st_mtime > start_timestamp


def scan_customer_files(customer: dict) -> Iterable[Path]:
    watch_dir = Path(customer["watch_dir"])
    if not watch_dir.exists():
        return []
    recursive = bool(customer.get("recursive", True))
    paths = watch_dir.rglob("*") if recursive else watch_dir.iterdir()
    return sorted(
        path
        for path in paths
        if is_candidate_excel(path)
        and matches_include_patterns(path, customer)
        and matches_start_mtime(path, customer)
    )
