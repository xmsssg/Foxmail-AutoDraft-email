from __future__ import annotations

import ctypes
import msvcrt
import os
import subprocess
import time
from dataclasses import dataclass
from pathlib import Path

from app.core.config_loader import expand_path


user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002
kernel32.GlobalAlloc.restype = ctypes.c_void_p
kernel32.GlobalLock.restype = ctypes.c_void_p
user32.SetClipboardData.restype = ctypes.c_void_p


EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


@dataclass
class FoxmailMapiImportResult:
    success: bool
    message: str
    imported_at: str | None = None
    foxmail_msg_id: int | None = None
    foxmail_mail_path: str | None = None


def _window_text(hwnd: int) -> str:
    buf = ctypes.create_unicode_buffer(512)
    user32.GetWindowTextW(hwnd, buf, 512)
    return buf.value


def _class_name(hwnd: int) -> str:
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buf, 256)
    return buf.value


def _window_rect(hwnd: int) -> tuple[int, int, int, int]:
    rect = RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    return rect.left, rect.top, rect.right, rect.bottom


def _enum_windows() -> list[int]:
    handles: list[int] = []

    def callback(hwnd: int, _lparam: int) -> bool:
        handles.append(hwnd)
        return True

    user32.EnumWindows(EnumWindowsProc(callback), 0)
    return handles


def _compose_windows() -> set[int]:
    handles: set[int] = set()
    for hwnd in _enum_windows():
        if not user32.IsWindow(hwnd):
            continue
        class_name = _class_name(hwnd)
        if not class_name.startswith("TFoxComposeForm"):
            continue
        left, top, right, bottom = _window_rect(hwnd)
        if right - left <= 100 or bottom - top <= 100:
            continue
        handles.add(hwnd)
    return handles


def _find_compose_window(
    subject: str,
    timeout_seconds: int,
    existing_handles: set[int] | None = None,
) -> int | None:
    existing_handles = existing_handles or set()
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        compose_handles = _compose_windows()
        new_handles = [hwnd for hwnd in compose_handles if hwnd not in existing_handles]
        preferred_handles = new_handles or list(compose_handles)
        for hwnd in preferred_handles:
            if hwnd in existing_handles and new_handles:
                continue
            title = _window_text(hwnd)
            if subject in title or "写邮件" in title:
                return hwnd
        time.sleep(0.5)
    return None


def _activate_window(hwnd: int) -> None:
    hwnd_topmost = ctypes.c_void_p(-1)
    hwnd_notopmost = ctypes.c_void_p(-2)
    swp_showwindow = 0x0040
    user32.ShowWindow(hwnd, 9)
    user32.SetWindowPos(hwnd, hwnd_topmost, 120, 70, 1200, 820, swp_showwindow)
    time.sleep(0.3)
    user32.SetForegroundWindow(hwnd)
    time.sleep(0.3)
    user32.SetWindowPos(hwnd, hwnd_notopmost, 120, 70, 1200, 820, swp_showwindow)


def _send_ctrl_s() -> None:
    # Use keybd_event to keep the MVP free of extra GUI automation packages.
    vk_control = 0x11
    vk_s = 0x53
    keyeventf_keyup = 0x0002
    user32.keybd_event(vk_control, 0, 0, 0)
    user32.keybd_event(vk_s, 0, 0, 0)
    time.sleep(0.1)
    user32.keybd_event(vk_s, 0, keyeventf_keyup, 0)
    user32.keybd_event(vk_control, 0, keyeventf_keyup, 0)


def _send_ctrl_a() -> None:
    vk_control = 0x11
    vk_a = 0x41
    keyeventf_keyup = 0x0002
    user32.keybd_event(vk_control, 0, 0, 0)
    user32.keybd_event(vk_a, 0, 0, 0)
    time.sleep(0.05)
    user32.keybd_event(vk_a, 0, keyeventf_keyup, 0)
    user32.keybd_event(vk_control, 0, keyeventf_keyup, 0)


def _send_ctrl_v() -> None:
    vk_control = 0x11
    vk_v = 0x56
    keyeventf_keyup = 0x0002
    user32.keybd_event(vk_control, 0, 0, 0)
    user32.keybd_event(vk_v, 0, 0, 0)
    time.sleep(0.05)
    user32.keybd_event(vk_v, 0, keyeventf_keyup, 0)
    user32.keybd_event(vk_control, 0, keyeventf_keyup, 0)


def _set_clipboard_text(text: str) -> None:
    data = text.encode("utf-16-le") + b"\x00\x00"
    if not user32.OpenClipboard(None):
        raise OSError(ctypes.get_last_error(), "无法打开剪贴板")
    try:
        if not user32.EmptyClipboard():
            raise OSError(ctypes.get_last_error(), "无法清空剪贴板")
        handle = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data))
        if not handle:
            raise OSError(ctypes.get_last_error(), "无法分配剪贴板内存")
        locked = kernel32.GlobalLock(handle)
        if not locked:
            raise OSError(ctypes.get_last_error(), "无法锁定剪贴板内存")
        try:
            ctypes.memmove(locked, data, len(data))
        finally:
            kernel32.GlobalUnlock(handle)
        if not user32.SetClipboardData(CF_UNICODETEXT, handle):
            raise OSError(ctypes.get_last_error(), "无法写入剪贴板")
    finally:
        user32.CloseClipboard()


def _click_compose_body(hwnd: int) -> None:
    left, top, right, bottom = _window_rect(hwnd)
    width = right - left
    height = bottom - top
    x = left + max(240, int(width * 0.28))
    y = top + max(330, int(height * 0.46))
    user32.SetCursorPos(x, y)
    user32.mouse_event(0x0002, 0, 0, 0, 0)
    time.sleep(0.05)
    user32.mouse_event(0x0004, 0, 0, 0, 0)
    time.sleep(0.2)


def _replace_compose_body(hwnd: int, body: str) -> None:
    if not body:
        return
    _set_clipboard_text(body)
    _click_compose_body(hwnd)
    _send_ctrl_a()
    time.sleep(0.1)
    _send_ctrl_v()
    time.sleep(0.3)


def _read_bytes_shared(path: Path) -> bytes:
    generic_read = 0x80000000
    file_share_read = 0x00000001
    file_share_write = 0x00000002
    file_share_delete = 0x00000004
    open_existing = 3
    file_attribute_normal = 0x00000080
    invalid_handle_value = ctypes.c_void_p(-1).value

    handle = kernel32.CreateFileW(
        str(path),
        generic_read,
        file_share_read | file_share_write | file_share_delete,
        None,
        open_existing,
        file_attribute_normal,
        None,
    )
    if handle == invalid_handle_value:
        raise OSError(ctypes.get_last_error(), f"无法共享读取文件: {path}")

    fd = msvcrt.open_osfhandle(handle, os.O_RDONLY)
    with os.fdopen(fd, "rb") as f:
        return f.read()


def _u32be(data: bytes, offset: int) -> int:
    return (
        (data[offset] << 24)
        | (data[offset + 1] << 16)
        | (data[offset + 2] << 8)
        | data[offset + 3]
    )


def _read_box_ids(path: Path) -> list[int]:
    data = _read_bytes_shared(path)
    ids: list[int] = []
    for offset in range(64, len(data) - 3, 4):
        value = int.from_bytes(data[offset : offset + 4], "little")
        if value:
            ids.append(value)
    return ids


def _find_new_draft(
    account_storage: Path,
    before_ids: set[int],
) -> tuple[int | None, Path | None]:
    boxes_dir = account_storage / "Boxes"
    draft_ids = _read_box_ids(boxes_dir / "draft.box")
    new_ids = [msg_id for msg_id in draft_ids if msg_id not in before_ids]
    if not new_ids:
        return None, None

    msg_id = max(new_ids)
    map_data = _read_bytes_shared(boxes_dir / "mId_bId.map")
    box_id = None
    for offset in range(32, len(map_data) - 11, 12):
        current_msg = _u32be(map_data, offset + 4)
        if current_msg == msg_id:
            box_id = _u32be(map_data, offset + 8)
            break
    if box_id != 4:
        return None, None

    mail_path = next(
        (p for p in (account_storage / "Mails").rglob(str(msg_id)) if p.is_file()),
        None,
    )
    return msg_id, mail_path


class FoxmailMapiDraftImporter:
    def __init__(self, foxmail_settings: dict) -> None:
        self.foxmail_exe = Path(expand_path(foxmail_settings["foxmail_exe"]))
        self.account_storage = Path(expand_path(foxmail_settings["account_storage"]))
        self.wait_window_seconds = int(foxmail_settings.get("wait_window_seconds", 15))
        self.wait_after_save_seconds = int(foxmail_settings.get("wait_after_save_seconds", 6))

    def import_xml(self, xml_path: Path, subject: str, body: str = "") -> FoxmailMapiImportResult:
        if not self.foxmail_exe.exists():
            return FoxmailMapiImportResult(False, f"Foxmail 不存在: {self.foxmail_exe}")
        if not xml_path.exists():
            return FoxmailMapiImportResult(False, f"FMMapi.xml 不存在: {xml_path}")

        draft_box = self.account_storage / "Boxes" / "draft.box"
        before_ids = set(_read_box_ids(draft_box)) if draft_box.exists() else set()

        existing_handles = _compose_windows()
        subprocess.Popen([str(self.foxmail_exe), "MAPI:", str(xml_path)])
        hwnd = _find_compose_window(subject, self.wait_window_seconds, existing_handles)
        if hwnd is None:
            return FoxmailMapiImportResult(False, "未找到 Foxmail 写信窗口")

        _activate_window(hwnd)
        _replace_compose_body(hwnd, body)
        _send_ctrl_s()
        time.sleep(self.wait_after_save_seconds)

        msg_id, mail_path = _find_new_draft(self.account_storage, before_ids)
        if msg_id is None:
            return FoxmailMapiImportResult(False, "未检测到新增 Foxmail 草稿")

        return FoxmailMapiImportResult(
            True,
            "Foxmail 草稿保存成功",
            imported_at=time.strftime("%Y-%m-%dT%H:%M:%S"),
            foxmail_msg_id=msg_id,
            foxmail_mail_path=str(mail_path) if mail_path else None,
        )
