"""리스트에서 지정한 Excel 시트·셀로 이동한다."""

from __future__ import annotations

import base64
import ctypes
import os
import re
import subprocess
import threading
from ctypes import wintypes
from typing import Callable, Dict, Optional, Tuple

from models import ExtractedImage, InspectionItem

JumpTarget = Tuple[str, str, str]
ErrorCallback = Callable[[str], None]

_CELL_COLUMN_TO_SIDE = {3: "ref", 4: "a", 5: "b", 6: "c"}
_SIDE_TO_COLUMN = {"ref": 3, "a": 4, "b": 5, "c": 6}
_CELL_PATTERN = re.compile(r"^[A-Z]{1,3}\d{1,7}$", re.IGNORECASE)


def side_for_list_column(column: int) -> Optional[str]:
    return _CELL_COLUMN_TO_SIDE.get(int(column))


def jump_target_for_list_cell(
    item: InspectionItem,
    column: int,
    workbook_paths: Dict[str, str],
) -> Optional[JumpTarget]:
    """더블클릭한 칸이 Excel로 이동할 대상이면 (경로, 시트, 셀)을 반환한다."""
    side = side_for_list_column(column)
    if side is None:
        return None
    extracted: Optional[ExtractedImage] = {
        "ref": item.image_ref,
        "a": item.image_a,
        "b": item.image_b,
        "c": item.image_c,
    }.get(side)
    if extracted is None or extracted.is_null:
        return None
    cell = str(extracted.cell_address or "").strip()
    if not cell or cell == "-" or _CELL_PATTERN.match(cell) is None:
        return None
    path = str(workbook_paths.get(side, "") or "").strip()
    if not path:
        return None
    return (os.path.abspath(os.path.expanduser(path)), extracted.sheet_name, cell.upper())


def jump_target_for_side(
    item: InspectionItem,
    side: str,
    workbook_paths: Dict[str, str],
) -> Optional[JumpTarget]:
    """미리보기 칸(ref/a/b/c)에 해당하는 Excel 이동 대상을 반환한다."""
    column = _SIDE_TO_COLUMN.get(str(side))
    if column is None:
        return None
    return jump_target_for_list_cell(item, column, workbook_paths)


def jump_to_excel_cell(
    path: str,
    sheet_name: str,
    cell_address: str,
    *,
    error_cb: Optional[ErrorCallback] = None,
) -> None:
    """Excel을 열고 시트·셀로 이동한다. GUI 스레드를 막지 않는다."""
    _allow_next_foreground()
    worker = threading.Thread(
        target=_jump_worker,
        args=(path, sheet_name, cell_address, error_cb),
        daemon=True,
        name="excel-jump",
    )
    worker.start()


def _jump_worker(
    path: str,
    sheet_name: str,
    cell_address: str,
    error_cb: Optional[ErrorCallback],
) -> None:
    try:
        if not os.path.isfile(path):
            raise FileNotFoundError(f"Excel 파일을 찾을 수 없습니다: {path}")
        if _jump_with_win32com(path, sheet_name, cell_address):
            return
        _jump_with_powershell(path, sheet_name, cell_address)
    except Exception as exc:
        if error_cb is not None:
            error_cb(str(exc))


def _jump_with_win32com(path: str, sheet_name: str, cell_address: str) -> bool:
    try:
        import win32com.client  # type: ignore
    except ImportError:
        return False
    excel = None
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    workbook = _find_open_workbook(excel, path)
    if workbook is None:
        alerts = excel.DisplayAlerts
        excel.DisplayAlerts = False
        try:
            workbook = excel.Workbooks.Open(path, UpdateLinks=0)
        finally:
            excel.DisplayAlerts = alerts
    worksheet = workbook.Worksheets(sheet_name)
    workbook.Activate()
    worksheet.Activate()
    excel.Goto(worksheet.Range(cell_address), True)
    _bring_excel_to_front(excel, workbook)
    return True


def _find_open_workbook(excel, path: str):
    target = os.path.normcase(os.path.abspath(path))
    for workbook in excel.Workbooks:
        try:
            full_name = str(workbook.FullName or "")
        except Exception:
            continue
        if full_name and os.path.normcase(os.path.abspath(full_name)) == target:
            return workbook
    return None


def _excel_window_hwnd(excel, workbook) -> Optional[int]:
    """해당 통합문서 창을 우선하고, 없으면 Excel 앱 창을 쓴다."""
    getters = (
        lambda: workbook.Windows(1).Hwnd,
        lambda: excel.ActiveWindow.Hwnd,
        lambda: excel.Hwnd,
    )
    for getter in getters:
        try:
            hwnd = int(getter())
        except Exception:
            continue
        if hwnd:
            return hwnd
    return None


def _allow_next_foreground() -> None:
    """더블클릭 직후, 우리 창이 포커스를 가진 동안 Excel이 앞으로 오게 허용한다."""
    try:
        _user32().AllowSetForegroundWindow(-1)
    except Exception:
        return


def _user32():
    return ctypes.windll.user32


def _kernel32():
    return ctypes.windll.kernel32


def _bring_hwnd_to_front(hwnd: int) -> None:
    """이미 떠 있는 Excel 창을 맨 앞으로 가져온다."""
    handle = int(hwnd)
    if handle <= 0:
        return
    user32 = _user32()
    if user32.IsIconic(handle):
        user32.ShowWindow(handle, 9)
    else:
        user32.ShowWindow(handle, 5)
    user32.BringWindowToTop(handle)
    if user32.SetForegroundWindow(handle):
        return

    kernel32 = _kernel32()
    foreground = int(user32.GetForegroundWindow() or 0)
    current_tid = int(kernel32.GetCurrentThreadId())
    pid = wintypes.DWORD()
    foreground_tid = int(
        user32.GetWindowThreadProcessId(foreground, ctypes.byref(pid)) or 0
    )
    target_tid = int(user32.GetWindowThreadProcessId(handle, ctypes.byref(pid)) or 0)
    attached_foreground = False
    attached_target = False
    try:
        if foreground_tid and foreground_tid != current_tid:
            attached_foreground = bool(
                user32.AttachThreadInput(current_tid, foreground_tid, True)
            )
        if target_tid and target_tid != current_tid and target_tid != foreground_tid:
            attached_target = bool(
                user32.AttachThreadInput(current_tid, target_tid, True)
            )
        user32.BringWindowToTop(handle)
        user32.SetForegroundWindow(handle)
    finally:
        if attached_target:
            user32.AttachThreadInput(current_tid, target_tid, False)
        if attached_foreground:
            user32.AttachThreadInput(current_tid, foreground_tid, False)

    if int(user32.GetForegroundWindow() or 0) != handle:
        _app_activate_hwnd(handle)


def _app_activate_hwnd(hwnd: int) -> None:
    try:
        import win32com.client  # type: ignore

        win32com.client.Dispatch("WScript.Shell").AppActivate(int(hwnd))
    except Exception:
        return


def _bring_excel_to_front(excel, workbook) -> None:
    try:
        excel.Visible = True
        try:
            if int(excel.WindowState) == -4140:
                excel.WindowState = -4143
        except Exception:
            pass
        try:
            excel.Activate()
        except Exception:
            pass
        hwnd = _excel_window_hwnd(excel, workbook)
        if hwnd is not None:
            _bring_hwnd_to_front(hwnd)
    except Exception:
        return


def _ps_literal(value: str) -> str:
    return "'" + str(value).replace("'", "''") + "'"


def _jump_with_powershell(path: str, sheet_name: str, cell_address: str) -> None:
    script = f"""
$path = {_ps_literal(path)}
$sheet = {_ps_literal(sheet_name)}
$cell = {_ps_literal(cell_address)}
$excel = $null
try {{
  $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
}} catch {{
  $excel = New-Object -ComObject Excel.Application
}}
$excel.Visible = $true
$target = [IO.Path]::GetFullPath($path)
$workbook = $null
foreach ($wb in @($excel.Workbooks)) {{
  if ($wb.FullName -and ([IO.Path]::GetFullPath($wb.FullName) -ieq $target)) {{
    $workbook = $wb
    break
  }}
}}
if (-not $workbook) {{
  $alerts = $excel.DisplayAlerts
  $excel.DisplayAlerts = $false
  try {{
    $workbook = $excel.Workbooks.Open($target, 0)
  }} finally {{
    $excel.DisplayAlerts = $alerts
  }}
}}
$worksheet = $workbook.Worksheets.Item($sheet)
$workbook.Activate()
$worksheet.Activate()
$excel.Goto($worksheet.Range($cell), $true)
try {{ $excel.Activate() }} catch {{}}
$hwnd = $null
try {{ $hwnd = $workbook.Windows.Item(1).Hwnd }} catch {{}}
if (-not $hwnd) {{ try {{ $hwnd = $excel.ActiveWindow.Hwnd }} catch {{}} }}
if (-not $hwnd) {{ $hwnd = $excel.Hwnd }}
if (-not ([System.Management.Automation.PSTypeName]'ExcelFront').Type) {{
  Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class ExcelFront {{
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
}}
"@
}}
$ptr = [IntPtr]$hwnd
if ([ExcelFront]::IsIconic($ptr)) {{
  [void][ExcelFront]::ShowWindow($ptr, 9)
}} else {{
  [void][ExcelFront]::ShowWindow($ptr, 5)
}}
[void][ExcelFront]::BringWindowToTop($ptr)
[void][ExcelFront]::SetForegroundWindow($ptr)
try {{
  $wshell = New-Object -ComObject WScript.Shell
  [void]$wshell.AppActivate($hwnd)
}} catch {{}}
"""
    encoded = base64.b64encode(script.encode("utf-16le")).decode("ascii")
    completed = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-Sta",
            "-ExecutionPolicy",
            "Bypass",
            "-EncodedCommand",
            encoded,
        ],
        capture_output=True,
        text=True,
        timeout=45,
        check=False,
    )
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout or "Excel 이동에 실패했습니다.").strip()
        raise RuntimeError(detail)
