"""EndNote session handling for the MCP server's job worker. Windows only, standard library only.

EndNote runs as one process per Windows user, and Cite While You Write talks to that single
process, so every job shares it. The behaviour below was measured on EndNote 21.5:

- ``EndNote.exe <library.enl>`` against a running instance opens that library in an additional
  window beside the libraries already open; it never closes the user's libraries.
- WM_CLOSE on one library window closes only that library. Closing the last one ends EndNote.
- EndNote shows news notices at start-up ("EndNote Password Expiration Policy Change",
  "EndNote 2025 is here!"), each a dialog with Learn More..., Cancel and Ignore This. While one
  is visible EndNote handles no request: a library forwarded meanwhile is lost and CWYW waits.
  Cancel dismisses a notice for this run only and changes no preference.
- A dismissed dialog stays behind as a hidden window, so only visible dialogs block anything.
- Ending the hidden Word while CWYW waits on its Select Matching Reference chooser leaves EndNote
  unable to exit: it ignores WM_CLOSE and File > Exit for minutes. Cancel on the chooser instead
  lets CWYW give up the citation, Word close normally and EndNote exit.

Nothing here terminates a process. Windows are closed with WM_CLOSE, the same request a click on
the close button sends, and the only button ever pressed is Cancel, on a news notice or on the
record chooser.
"""

from __future__ import annotations

import ctypes
import glob
import logging
import os
import subprocess
import time
import winreg
from ctypes import wintypes
from dataclasses import dataclass
from pathlib import Path

LOGGER = logging.getLogger(__name__)

DIALOG_CLASS = "#32770"
WORD_DIALOG_CLASS_PREFIX = "bosa_sdm"
NOTICE_TITLE_PREFIX = "EndNote"
CANCEL_TEXT = "Cancel"
NOTICE_BUTTONS = frozenset({"Learn More...", CANCEL_TEXT, "Ignore This"})
CANCEL_CONTROL_ID = 2
WM_CLOSE = 0x0010
BM_CLICK = 0x00F5
GW_OWNER = 4
WINDOW_TEXT_CHARACTERS = 1024
CLASS_NAME_CHARACTERS = 256
LIBRARY_EXTENSION = ".enl"
TITLE_SEPARATOR = " - "
LIBRARY_WINDOW_CLASS_PREFIX = "Afx:"

RECENT_LIBRARIES_KEY = r"Software\ISI ResearchSoft\EndNote\Recent Libraries"
RECENT_LIBRARY_VALUE_PREFIX = "File"
APP_PATHS_KEYS = (
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\EndNote.exe"),
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EndNote.exe"),
)
PROGRAM_FILES_VARIABLES = ("ProgramFiles(x86)", "ProgramFiles")
INSTALL_GLOB = os.path.join("EndNote*", "EndNote.exe")

POLL_SECONDS = 0.25
# A library forwarded while a notice was up is lost, so the open request is repeated this often.
REFORWARD_SECONDS = 10

_user32 = ctypes.WinDLL("user32", use_last_error=True)
_ENUM_PROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
_user32.EnumWindows.argtypes = (_ENUM_PROC, wintypes.LPARAM)
_user32.EnumChildWindows.argtypes = (wintypes.HWND, _ENUM_PROC, wintypes.LPARAM)
_user32.GetWindowThreadProcessId.argtypes = (wintypes.HWND, ctypes.POINTER(wintypes.DWORD))
_user32.GetWindowTextW.argtypes = (wintypes.HWND, wintypes.LPWSTR, ctypes.c_int)
_user32.GetClassNameW.argtypes = (wintypes.HWND, wintypes.LPWSTR, ctypes.c_int)
_user32.IsWindowVisible.argtypes = (wintypes.HWND,)
_user32.IsWindow.argtypes = (wintypes.HWND,)
_user32.GetWindow.argtypes = (wintypes.HWND, wintypes.UINT)
_user32.GetWindow.restype = wintypes.HWND
_user32.GetDlgCtrlID.argtypes = (wintypes.HWND,)
_user32.PostMessageW.argtypes = (wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
_user32.SendMessageW.argtypes = (wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)


@dataclass(frozen=True)
class Window:
    handle: int
    process_id: int
    class_name: str
    title: str
    visible: bool
    owned: bool


def _text(handle: int) -> str:
    buffer = ctypes.create_unicode_buffer(WINDOW_TEXT_CHARACTERS)
    _user32.GetWindowTextW(handle, buffer, WINDOW_TEXT_CHARACTERS)
    return buffer.value


def _class_name(handle: int) -> str:
    buffer = ctypes.create_unicode_buffer(CLASS_NAME_CHARACTERS)
    _user32.GetClassNameW(handle, buffer, CLASS_NAME_CHARACTERS)
    return buffer.value


def windows_of(process_ids: set[int]) -> list[Window]:
    """Every top-level window owned by one of process_ids, visible or not."""
    found: list[Window] = []

    def visit(handle, _):
        owner = wintypes.DWORD()
        _user32.GetWindowThreadProcessId(handle, ctypes.byref(owner))
        if owner.value in process_ids:
            found.append(Window(
                int(handle), owner.value, _class_name(handle), _text(handle),
                bool(_user32.IsWindowVisible(handle)), bool(_user32.GetWindow(handle, GW_OWNER)),
            ))
        return True

    _user32.EnumWindows(_ENUM_PROC(visit), 0)
    return found


def buttons_of(handle: int) -> list[tuple[int, str, int]]:
    """(control id, text, handle) of the push buttons inside a dialog."""
    found: list[tuple[int, str, int]] = []

    def visit(child, _):
        if _class_name(child) == "Button":
            found.append((int(_user32.GetDlgCtrlID(child)), _text(child), int(child)))
        return True

    _user32.EnumChildWindows(handle, _ENUM_PROC(visit), 0)
    return found


def is_notice(title: str, button_texts: set[str]) -> bool:
    """The EndNote news-notice template: an EndNote title with exactly these three buttons."""
    return title.startswith(NOTICE_TITLE_PREFIX) and button_texts == NOTICE_BUTTONS


def is_dialog(window: Window) -> bool:
    return window.class_name == DIALOG_CLASS or window.class_name.startswith(WORD_DIALOG_CLASS_PREFIX)


def press_cancel(dialog_handle: int) -> bool:
    """Click the Cancel button of a dialog; False when it has none."""
    for control_id, text, handle in buttons_of(dialog_handle):
        if control_id == CANCEL_CONTROL_ID or text == CANCEL_TEXT:
            _user32.SendMessageW(handle, BM_CLICK, 0, 0)
            return True
    return False


def dismiss_notices(endnote_process_ids: set[int]) -> list[str]:
    """Press Cancel on every visible EndNote news notice and return their titles."""
    dismissed: list[str] = []
    for window in windows_of(endnote_process_ids):
        if not (window.visible and window.class_name == DIALOG_CLASS):
            continue
        if not is_notice(window.title, {text for _, text, _ in buttons_of(window.handle)}):
            continue
        if press_cancel(window.handle):
            dismissed.append(window.title)
            LOGGER.info("Dismissed the EndNote notice %r with Cancel.", window.title)
    return dismissed


def blocking_dialogs(process_ids: set[int]) -> list[dict[str, object]]:
    """Visible dialogs other than news notices; reported, never clicked."""
    found: list[dict[str, object]] = []
    for window in windows_of(process_ids):
        if not (window.visible and is_dialog(window)):
            continue
        texts = [text for _, text, _ in buttons_of(window.handle)]
        if not is_notice(window.title, set(texts)):
            found.append({"handle": window.handle, "process_id": window.process_id, "title": window.title, "buttons": texts})
    return found


def library_title_matches(title: str, library: Path) -> bool:
    """EndNote titles a library window '<product> - <file name>', and later just '<file name>'
    (measured: the prefix is gone once a CWYW job has finished)."""
    folded, name = title.casefold(), library.name.casefold()
    return folded == name or folded.endswith((TITLE_SEPARATOR + library.name).casefold())


def library_windows(process_ids: set[int]) -> list[Window]:
    return [
        window for window in windows_of(process_ids)
        if window.class_name.startswith(LIBRARY_WINDOW_CLASS_PREFIX) and window.title.casefold().endswith(LIBRARY_EXTENSION)
    ]


def library_window(process_ids: set[int], library: Path) -> Window | None:
    return next((window for window in library_windows(process_ids) if library_title_matches(window.title, library)), None)


def close_window(handle: int) -> None:
    _user32.PostMessageW(handle, WM_CLOSE, 0, 0)


def window_exists(handle: int) -> bool:
    return bool(_user32.IsWindow(handle))


def recent_libraries() -> dict[str, str]:
    """EndNote's recent-library list, which decides the library it opens on the next start."""
    values: dict[str, str] = {}
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RECENT_LIBRARIES_KEY) as key:
            index = 0
            while True:
                try:
                    name, value, _ = winreg.EnumValue(key, index)
                except OSError:
                    break
                if name.startswith(RECENT_LIBRARY_VALUE_PREFIX):
                    values[name] = str(value)
                index += 1
    except FileNotFoundError:
        pass
    return values


def restore_recent_libraries(snapshot: dict[str, str]) -> None:
    """Put back the list a job found, so opening a document's library does not reorder it."""
    current = recent_libraries()
    if current == snapshot:
        return
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, RECENT_LIBRARIES_KEY) as key:
        for name, value in snapshot.items():
            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
        for name in set(current) - set(snapshot):
            winreg.DeleteValue(key, name)


def endnote_executable() -> str | None:
    """The newest EndNote.exe from App Paths, then from Program Files."""
    for hive, path in APP_PATHS_KEYS:
        try:
            with winreg.OpenKey(hive, path) as key:
                value, _ = winreg.QueryValueEx(key, "")
            if value and Path(str(value).strip('"')).is_file():
                return str(value).strip('"')
        except OSError:
            continue
    candidates: list[str] = []
    for variable in PROGRAM_FILES_VARIABLES:
        root = os.environ.get(variable)
        if root:
            candidates.extend(glob.glob(os.path.join(root, INSTALL_GLOB)))
    return sorted(candidates)[-1] if candidates else None


def launch_endnote(executable: str, library: Path, hidden: bool) -> subprocess.Popen:
    """Start EndNote with a library, or forward the library to the running instance."""
    startup = subprocess.STARTUPINFO()
    if hidden:
        startup.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startup.wShowWindow = 0  # SW_HIDE, as the bridge's Start-Process -WindowStyle Hidden
    return subprocess.Popen([executable, str(library)], startupinfo=startup, stdin=subprocess.DEVNULL,
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def frame_windows(process_ids: set[int]) -> list[Window]:
    """EndNote's main frame; closing it is the fallback when no library window is open."""
    return [window for window in windows_of(process_ids) if window.class_name.endswith(" Frame")]


def wait_for_library(endnote_pids, library: Path, timeout_seconds: float, reforward) -> tuple[Window | None, list[str]]:
    """Wait for the library window, dismissing notices and repeating the open request meanwhile.
    Returns the window, or None at the deadline, and the titles of the notices dismissed."""
    deadline = time.monotonic() + timeout_seconds
    next_forward = time.monotonic() + REFORWARD_SECONDS
    dismissed: list[str] = []
    while time.monotonic() < deadline:
        process_ids = endnote_pids()
        dismissed.extend(dismiss_notices(process_ids))
        window = library_window(process_ids, library)
        if window is not None:
            return window, dismissed
        if time.monotonic() >= next_forward:
            reforward()
            next_forward = time.monotonic() + REFORWARD_SECONDS
        time.sleep(POLL_SECONDS)
    return None, dismissed
