from __future__ import annotations

import time
from pathlib import Path


def is_file_stable(path: Path, checks: int = 3, interval_seconds: float = 1.0) -> bool:
    last_signature: tuple[int, float] | None = None
    stable_count = 0

    for _ in range(checks):
        if not path.exists() or not path.is_file():
            return False
        stat = path.stat()
        signature = (stat.st_size, stat.st_mtime)
        if signature == last_signature:
            stable_count += 1
        else:
            stable_count = 1
            last_signature = signature
        time.sleep(interval_seconds)

    return stable_count >= checks
