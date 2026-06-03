from __future__ import annotations

import ctypes
import time
from pathlib import Path


if hasattr(ctypes, "WinDLL"):
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateFileW.restype = ctypes.c_void_p
else:
    kernel32 = None


def is_file_available(path: Path) -> bool:
    if kernel32 is None:
        try:
            with path.open("rb"):
                return True
        except OSError:
            return False

    generic_read = 0x80000000
    open_existing = 3
    file_attribute_normal = 0x00000080
    invalid_handle_value = ctypes.c_void_p(-1).value

    handle = kernel32.CreateFileW(
        str(path),
        generic_read,
        0,
        None,
        open_existing,
        file_attribute_normal,
        None,
    )
    if handle == invalid_handle_value:
        return False
    kernel32.CloseHandle(handle)
    return True


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


def is_file_ready(path: Path, checks: int = 3, interval_seconds: float = 1.0) -> bool:
    return is_file_stable(path, checks, interval_seconds) and is_file_available(path)
