"""리스트 칸에서 Excel 이동 대상을 고르는 순수 로직."""

import time
from types import SimpleNamespace

from models import ExtractedImage, InspectionItem
from excel_jump import (
    _bring_excel_to_front,
    _bring_hwnd_to_front,
    _excel_window_hwnd,
    jump_target_for_list_cell,
    jump_to_excel_cell,
    side_for_list_column,
)


def _image(cell: str, *, null: bool = False) -> ExtractedImage:
    return ExtractedImage(
        sheet_index=1,
        sheet_name="MAIN",
        cell_address=cell,
        merged_range="",
        anchor_row=0,
        anchor_col=0,
        is_null=null,
    )


def _item(*, a_null: bool = False) -> InspectionItem:
    return InspectionItem(
        sheet_index=1,
        index=1,
        image_ref=_image("D10"),
        image_a=_image("E11", null=a_null),
        image_b=_image("F12"),
        image_c=_image("G13"),
    )


def test_only_ref_compare_columns_are_hyperlink_targets():
    assert side_for_list_column(0) is None
    assert side_for_list_column(1) is None
    assert side_for_list_column(2) is None
    assert side_for_list_column(3) == "ref"
    assert side_for_list_column(4) == "a"
    assert side_for_list_column(5) == "b"
    assert side_for_list_column(6) == "c"


def test_jump_target_uses_matching_workbook_sheet_and_cell(tmp_path):
    paths = {
        "ref": str(tmp_path / "ref.xlsx"),
        "a": str(tmp_path / "a.xlsx"),
        "b": str(tmp_path / "b.xlsx"),
        "c": str(tmp_path / "c.xlsx"),
    }
    item = _item()
    target = jump_target_for_list_cell(item, 3, paths)
    assert target is not None
    assert target[0].endswith("ref.xlsx")
    assert target[1] == "MAIN"
    assert target[2] == "D10"
    assert jump_target_for_list_cell(item, 4, paths)[2] == "E11"
    assert jump_target_for_list_cell(item, 5, paths)[2] == "F12"
    assert jump_target_for_list_cell(item, 6, paths)[2] == "G13"


def test_jump_target_skips_index_columns_and_missing_images(tmp_path):
    paths = {"ref": str(tmp_path / "ref.xlsx"), "a": str(tmp_path / "a.xlsx")}
    item = _item(a_null=True)
    assert jump_target_for_list_cell(item, 0, paths) is None
    assert jump_target_for_list_cell(item, 1, paths) is None
    assert jump_target_for_list_cell(item, 2, paths) is None
    assert jump_target_for_list_cell(item, 4, paths) is None
    assert jump_target_for_list_cell(item, 3, paths) is not None
    assert jump_target_for_list_cell(item, 3, {"ref": ""}) is None


def test_jump_target_rejects_invalid_or_placeholder_cells(tmp_path):
    paths = {"ref": str(tmp_path / "ref.xlsx")}

    def item_with_cell(cell: str) -> InspectionItem:
        return InspectionItem(
            sheet_index=1,
            index=1,
            image_ref=_image(cell),
            image_a=_image("E11"),
        )

    assert jump_target_for_list_cell(item_with_cell(""), 3, paths) is None
    assert jump_target_for_list_cell(item_with_cell("-"), 3, paths) is None
    assert jump_target_for_list_cell(item_with_cell("이미지 없음"), 3, paths) is None
    assert jump_target_for_list_cell(item_with_cell("A1:B2"), 3, paths) is None
    target = jump_target_for_list_cell(item_with_cell("d10"), 3, paths)
    assert target is not None
    assert target[2] == "D10"


def test_jump_to_excel_cell_returns_without_waiting(monkeypatch):
    started = []
    allowed = []

    def fake_worker(*args):
        started.append(args)
        time.sleep(0.4)

    monkeypatch.setattr("excel_jump._jump_worker", fake_worker)
    monkeypatch.setattr(
        "excel_jump._allow_next_foreground",
        lambda: allowed.append(True),
    )
    began = time.perf_counter()
    jump_to_excel_cell(r"C:\missing.xlsx", "MAIN", "A1")
    elapsed = time.perf_counter() - began
    assert elapsed < 0.2
    assert allowed == [True]
    deadline = time.perf_counter() + 1.0
    while not started and time.perf_counter() < deadline:
        time.sleep(0.01)
    assert started


class _FakeUser32:
    def __init__(self, *, iconic=False, set_foreground_ok=True):
        self.iconic = iconic
        self.set_foreground_ok = set_foreground_ok
        self.foreground = 7
        self.calls = []

    def IsIconic(self, hwnd):
        return self.iconic

    def ShowWindow(self, hwnd, command):
        self.calls.append(("show", int(hwnd), int(command)))
        return True

    def BringWindowToTop(self, hwnd):
        self.calls.append(("top", int(hwnd)))
        return True

    def SetForegroundWindow(self, hwnd):
        self.calls.append(("foreground", int(hwnd)))
        if self.set_foreground_ok:
            self.foreground = int(hwnd)
        return self.set_foreground_ok

    def GetForegroundWindow(self):
        return self.foreground

    def GetWindowThreadProcessId(self, hwnd, _pid):
        return 20 if hwnd else 0

    def AttachThreadInput(self, current, other, attach):
        self.calls.append(("attach", int(current), int(other), bool(attach)))
        return True


class _FakeKernel32:
    def GetCurrentThreadId(self):
        return 11


def test_excel_window_hwnd_prefers_workbook_window():
    workbook = SimpleNamespace(Windows=lambda _index: SimpleNamespace(Hwnd=4242))
    excel = SimpleNamespace(Hwnd=1111, ActiveWindow=SimpleNamespace(Hwnd=2222))
    assert _excel_window_hwnd(excel, workbook) == 4242


def test_excel_window_hwnd_falls_back_to_excel_hwnd():
    workbook = SimpleNamespace()
    excel = SimpleNamespace(Hwnd=1111)
    assert _excel_window_hwnd(excel, workbook) == 1111


def test_bring_hwnd_to_front_restores_and_sets_foreground(monkeypatch):
    user32 = _FakeUser32(iconic=True, set_foreground_ok=True)
    monkeypatch.setattr("excel_jump._user32", lambda: user32)
    monkeypatch.setattr("excel_jump._kernel32", lambda: _FakeKernel32())
    activated = []
    monkeypatch.setattr("excel_jump._app_activate_hwnd", activated.append)

    _bring_hwnd_to_front(4242)

    assert ("show", 4242, 9) in user32.calls
    assert ("foreground", 4242) in user32.calls
    assert activated == []


def test_bring_hwnd_to_front_retries_when_windows_blocks_focus(monkeypatch):
    user32 = _FakeUser32(iconic=False, set_foreground_ok=False)
    monkeypatch.setattr("excel_jump._user32", lambda: user32)
    monkeypatch.setattr("excel_jump._kernel32", lambda: _FakeKernel32())
    activated = []
    monkeypatch.setattr("excel_jump._app_activate_hwnd", activated.append)

    _bring_hwnd_to_front(4242)

    assert ("show", 4242, 5) in user32.calls
    assert any(call[0] == "attach" for call in user32.calls)
    assert activated == [4242]


def test_allow_next_foreground_asks_windows_to_let_excel_steal_focus(monkeypatch):
    calls = []

    class _AllowUser32:
        def AllowSetForegroundWindow(self, pid):
            calls.append(int(pid))
            return True

    monkeypatch.setattr("excel_jump._user32", lambda: _AllowUser32())
    from excel_jump import _allow_next_foreground

    _allow_next_foreground()
    assert calls == [-1]


def test_bring_excel_to_front_restores_minimized_app_then_window(monkeypatch):
    brought = []
    monkeypatch.setattr("excel_jump._bring_hwnd_to_front", brought.append)
    excel = SimpleNamespace(
        Visible=False,
        WindowState=-4140,
        Hwnd=1111,
        Activate=lambda: setattr(excel, "activated", True),
    )
    workbook = SimpleNamespace(Windows=lambda _index: SimpleNamespace(Hwnd=4242))

    _bring_excel_to_front(excel, workbook)

    assert excel.Visible is True
    assert excel.WindowState == -4143
    assert excel.activated is True
    assert brought == [4242]
