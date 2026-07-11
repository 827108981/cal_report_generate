from __future__ import annotations

import ctypes
import queue
import sys
import threading
import time
from ctypes import wintypes


WM_HOTKEY = 0x0312
PM_REMOVE = 0x0001
MOD_NOREPEAT = 0x4000
VK_F8 = 0x77
VK_F9 = 0x78
HOTKEY_REGION = 0xB501
HOTKEY_FULLSCREEN = 0xB502


class _Point(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


class _Message(ctypes.Structure):
    _fields_ = [
        ("hwnd", wintypes.HWND),
        ("message", wintypes.UINT),
        ("wParam", wintypes.WPARAM),
        ("lParam", wintypes.LPARAM),
        ("time", wintypes.DWORD),
        ("pt", _Point),
        ("lPrivate", wintypes.DWORD),
    ]


class GlobalHotkeyListener:
    """Windows-native F8/F9 listener that hands events back to Tk through a queue."""

    def __init__(self) -> None:
        self.events: queue.SimpleQueue[str] = queue.SimpleQueue()
        self._stop = threading.Event()
        self._ready = threading.Event()
        self._thread: threading.Thread | None = None
        self.available = {"region": False, "fullscreen": False}

    def start(self) -> dict[str, bool]:
        if sys.platform != "win32":
            return self.available.copy()
        if self._thread and self._thread.is_alive():
            return self.available.copy()
        self._stop.clear()
        self._ready.clear()
        self._thread = threading.Thread(target=self._run, name="global-screenshot-hotkeys", daemon=True)
        self._thread.start()
        self._ready.wait(timeout=1.0)
        return self.available.copy()

    def stop(self) -> None:
        self._stop.set()
        if self._thread:
            self._thread.join(timeout=0.5)
        self._thread = None

    def drain(self) -> list[str]:
        result: list[str] = []
        while True:
            try:
                result.append(self.events.get_nowait())
            except queue.Empty:
                return result

    def _run(self) -> None:
        user32 = ctypes.windll.user32
        msg = _Message()
        # PeekMessage creates this worker's message queue before RegisterHotKey.
        user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 0)
        self.available["region"] = bool(user32.RegisterHotKey(None, HOTKEY_REGION, MOD_NOREPEAT, VK_F8))
        self.available["fullscreen"] = bool(user32.RegisterHotKey(None, HOTKEY_FULLSCREEN, MOD_NOREPEAT, VK_F9))
        self._ready.set()
        try:
            while not self._stop.is_set():
                while user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, PM_REMOVE):
                    if msg.message == WM_HOTKEY:
                        if msg.wParam == HOTKEY_REGION:
                            self.events.put("region")
                        elif msg.wParam == HOTKEY_FULLSCREEN:
                            self.events.put("fullscreen")
                time.sleep(0.03)
        finally:
            if self.available["region"]:
                user32.UnregisterHotKey(None, HOTKEY_REGION)
            if self.available["fullscreen"]:
                user32.UnregisterHotKey(None, HOTKEY_FULLSCREEN)
