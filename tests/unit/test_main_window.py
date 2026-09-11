"""Double/Triple 육안 검사 메인 화면의 탐색 smoke test."""

from pathlib import Path

from PIL import Image
from PyQt6.QtCore import QSettings

import ui.main_window as main_window_module
from models import MISSING_SHEET, ExtractedImage, InspectionItem, SheetInfo


def _extracted(
    path: Path | None,
    *,
    sheet_index: int,
    sheet_name: str,
    cell: str,
) -> ExtractedImage:
    return ExtractedImage(
        sheet_index=sheet_index,
        sheet_name=sheet_name,
        cell_address=cell,
        merged_range="",
        anchor_row=0,
        anchor_col=0,
        is_null=path is None,
        source_path=str(path) if path is not None else None,
        preview_path=str(path) if path is not None else None,
    )


def _png(path: Path, color: tuple[int, int, int]) -> Path:
    Image.new("RGB", (80, 50), color).save(path)
    return path


def _window(tmp_path, monkeypatch):
    settings_path = tmp_path / "settings.ini"
    monkeypatch.setattr(
        main_window_module,
        "QSettings",
        lambda *_args: QSettings(
            str(settings_path), QSettings.Format.IniFormat
        ),
    )
    return main_window_module.MainWindow()


def test_double_navigation_shows_two_images_without_comparison_controls(
    qapp, tmp_path, monkeypatch
):
    ref_path = _png(tmp_path / "ref.png", (200, 20, 20))
    a_path = _png(tmp_path / "a.png", (20, 200, 20))
    first = InspectionItem(
        1,
        1,
        _extracted(ref_path, sheet_index=1, sheet_name="MAIN", cell="A1"),
        _extracted(a_path, sheet_index=1, sheet_name="MAIN", cell="A1"),
    )
    second = InspectionItem(
        1,
        2,
        _extracted(ref_path, sheet_index=1, sheet_name="MAIN", cell="D10"),
        _extracted(None, sheet_index=1, sheet_name="MAIN", cell="D10"),
    )
    window = _window(tmp_path, monkeypatch)
    try:
        window.show()
        window.state.set_items(
            {1: [first, second]},
            mode="double",
            workbook_paths={"ref": "ref.xlsx", "a": "a.xlsx"},
        )
        window._populate_sheet_combo()
        window._render_current()
        qapp.processEvents()

        assert window.current_mode_label.text() == "현재 결과: Double"
        assert window.position_label.text() == "전체 1 / 2"
        assert window.ref_image.pixmap() is not None
        assert window.a_image.pixmap() is not None
        assert window.b_container.isHidden()
        assert not hasattr(window, "threshold_spin")
        assert not hasattr(window, "summary_table")
        assert not hasattr(window, "export_button")

        window._move(1)
        assert window.position_label.text() == "전체 2 / 2"
        assert "이미지가 없습니다" in window.a_image.text()
        assert window.previous_button.isEnabled()
        assert not window.next_button.isEnabled()
    finally:
        window.close()
        window.deleteLater()


def test_triple_shows_third_image_and_marks_input_mode_change(
    qapp, tmp_path, monkeypatch
):
    ref_path = _png(tmp_path / "ref.png", (200, 20, 20))
    a_path = _png(tmp_path / "a.png", (20, 200, 20))
    b_path = _png(tmp_path / "b.png", (20, 20, 200))
    item = InspectionItem(
        1,
        1,
        _extracted(ref_path, sheet_index=1, sheet_name="MAIN", cell="A1"),
        _extracted(a_path, sheet_index=1, sheet_name="MAIN", cell="A1"),
        _extracted(b_path, sheet_index=1, sheet_name="MAIN", cell="A1"),
    )
    window = _window(tmp_path, monkeypatch)
    try:
        window.show()
        window.triple_check.setChecked(True)
        window.state.set_items(
            {1: [item]},
            mode="triple",
            workbook_paths={
                "ref": "ref.xlsx",
                "a": "a.xlsx",
                "b": "b.xlsx",
            },
        )
        window._populate_sheet_combo()
        window._render_current()
        qapp.processEvents()

        assert window.current_mode_label.text() == "현재 결과: Triple"
        assert not window.b_container.isHidden()
        assert window.b_image.pixmap() is not None

        window.double_check.setChecked(True)
        assert "현재 화면은 이전 Triple 결과" in window.status_label.text()
        assert not window.b_container.isHidden()
    finally:
        window.close()
        window.deleteLater()


def test_main_shows_sheet_role_status_and_sheet_missing_message(
    qapp, tmp_path, monkeypatch
):
    ref_path = _png(tmp_path / "ref-only.png", (200, 20, 20))
    info = SheetInfo(
        sheet_index=1,
        display_name="REF_ONLY",
        selected_roles=("ref", "a"),
        role_sheet_names={"ref": "REF_ONLY", "a": None},
        role_sheet_ids={"ref": "ref-sheet", "a": None},
    )
    ref_image = _extracted(
        ref_path, sheet_index=1, sheet_name="REF_ONLY", cell="A1"
    )
    missing_a = ExtractedImage(
        sheet_index=1,
        sheet_name="REF_ONLY",
        cell_address="A1",
        merged_range="",
        anchor_row=0,
        anchor_col=0,
        is_null=True,
        source_role="a",
        missing_reason=MISSING_SHEET,
    )
    item = InspectionItem(
        sheet_index=1,
        index=1,
        image_ref=ref_image,
        image_a=missing_a,
        sheet_info=info,
    )
    window = _window(tmp_path, monkeypatch)
    try:
        window.show()
        window.state.set_items(
            {1: [item]},
            mode="double",
            workbook_paths={"ref": "ref.xlsx", "a": "a.xlsx"},
            sheet_infos={1: info},
        )
        window._populate_sheet_combo()
        window._render_current()
        qapp.processEvents()

        assert "REF_ONLY [REF ONLY]" in window.sheet_combo.itemText(0)
        assert "REF_ONLY [REF ONLY]" in window.location_label.text()
        assert "시트 없음" in window.a_meta.text()
        assert window.a_image.text() == "현재 Excel에는 이 시트가 없습니다."
    finally:
        window.close()
        window.deleteLater()


def test_quadra_shows_fourth_image(qapp, tmp_path, monkeypatch):
    ref_path = _png(tmp_path / "ref.png", (200, 20, 20))
    a_path = _png(tmp_path / "a.png", (20, 200, 20))
    b_path = _png(tmp_path / "b.png", (20, 20, 200))
    c_path = _png(tmp_path / "c.png", (200, 200, 20))
    item = InspectionItem(
        1,
        1,
        _extracted(ref_path, sheet_index=1, sheet_name="MAIN", cell="A1"),
        _extracted(a_path, sheet_index=1, sheet_name="MAIN", cell="A1"),
        _extracted(b_path, sheet_index=1, sheet_name="MAIN", cell="A1"),
        _extracted(c_path, sheet_index=1, sheet_name="MAIN", cell="A1"),
    )
    window = _window(tmp_path, monkeypatch)
    try:
        window.show()
        window.quadra_check.setChecked(True)
        window.state.set_items(
            {1: [item]},
            mode="quadra",
            workbook_paths={
                "ref": "ref.xlsx",
                "a": "a.xlsx",
                "b": "b.xlsx",
                "c": "c.xlsx",
            },
        )
        window._populate_sheet_combo()
        window._render_current()
        qapp.processEvents()

        assert window.current_mode_label.text() == "현재 결과: Quadra"
        assert not window.b_container.isHidden()
        assert not window.c_container.isHidden()
        assert window.c_image.pixmap() is not None
        assert not window.b_row.isHidden()
        assert not window.c_row.isHidden()
    finally:
        window.close()
        window.deleteLater()


def test_mode_checkboxes_are_exclusive(qapp, tmp_path, monkeypatch):
    window = _window(tmp_path, monkeypatch)
    try:
        window.show()
        assert window.double_check.isChecked()
        assert not window.triple_check.isChecked()
        assert not window.quadra_check.isChecked()
        assert window.b_row.isHidden()
        assert window.c_row.isHidden()

        window.triple_check.setChecked(True)
        assert window.triple_check.isChecked()
        assert not window.double_check.isChecked()
        assert not window.quadra_check.isChecked()
        assert not window.b_row.isHidden()
        assert window.c_row.isHidden()

        window.quadra_check.setChecked(True)
        assert window.quadra_check.isChecked()
        assert not window.double_check.isChecked()
        assert not window.triple_check.isChecked()
        assert not window.b_row.isHidden()
        assert not window.c_row.isHidden()

        window.quadra_check.setChecked(False)
        assert window.quadra_check.isChecked()
        assert not window.double_check.isChecked()
        assert not window.triple_check.isChecked()
        assert window._selected_mode() == "quadra"
    finally:
        window.close()
        window.deleteLater()


def test_usage_help_covers_modes_list_and_hyperlink(qapp):
    from ui.dialogs import USAGE_HELP_TEXT, UsageHelpDialog

    assert "Double (Excel 2개)" in USAGE_HELP_TEXT
    assert "Triple (Excel 3개)" in USAGE_HELP_TEXT
    assert "Quadra (Excel 4개)" in USAGE_HELP_TEXT
    assert "리스트실행" in USAGE_HELP_TEXT
    assert "하이퍼링크 모드" in USAGE_HELP_TEXT
    assert "엑셀 파형 바로가기" in USAGE_HELP_TEXT
    assert "오른쪽 끝" in USAGE_HELP_TEXT
    assert "PASS/FAIL" in USAGE_HELP_TEXT
    assert "Ctrl+Enter" in USAGE_HELP_TEXT
    dialog = UsageHelpDialog()
    try:
        assert dialog.windowTitle() == "사용 방법"
        assert "하이퍼링크 모드" in dialog.text.toPlainText()
    finally:
        dialog.close()
        dialog.deleteLater()


def test_main_hyperlink_checkbox_and_preview_jump_buttons(
    qapp, tmp_path, monkeypatch
):
    import ui.dialogs as dialogs_module

    jumps = []
    monkeypatch.setattr(
        dialogs_module,
        "jump_to_excel_cell",
        lambda path, sheet, cell, error_cb=None: jumps.append((path, sheet, cell)),
    )
    ref_xlsx = tmp_path / "ref.xlsx"
    a_xlsx = tmp_path / "a.xlsx"
    ref_xlsx.write_bytes(b"xlsx")
    a_xlsx.write_bytes(b"xlsx")
    ref_path = _png(tmp_path / "ref.png", (200, 20, 20))
    a_path = _png(tmp_path / "a.png", (20, 200, 20))
    first = InspectionItem(
        1,
        1,
        _extracted(ref_path, sheet_index=1, sheet_name="MAIN", cell="D10"),
        _extracted(a_path, sheet_index=1, sheet_name="MAIN", cell="E11"),
    )
    second = InspectionItem(
        1,
        2,
        _extracted(ref_path, sheet_index=1, sheet_name="MAIN", cell="F12"),
        _extracted(None, sheet_index=1, sheet_name="MAIN", cell="F12"),
    )
    window = _window(tmp_path, monkeypatch)
    try:
        window.show()
        window.state.set_items(
            {1: [first, second]},
            mode="double",
            workbook_paths={"ref": str(ref_xlsx), "a": str(a_xlsx)},
        )
        window._populate_sheet_combo()
        window._render_current()
        qapp.processEvents()

        assert window.hyperlink_check.text() == "하이퍼링크 모드"
        assert window.hyperlink_check.isChecked() is False
        assert window.ref_jump_button.text() == "엑셀 파형 바로가기"
        assert window.ref_jump_button.isEnabled() is False
        assert window.a_jump_button.isEnabled() is False
        assert window.b_jump_button.isEnabled() is False
        assert window.c_jump_button.isEnabled() is False
        window.ref_jump_button.click()
        assert jumps == []

        window.hyperlink_check.setChecked(True)
        qapp.processEvents()
        assert window.ref_jump_button.isEnabled() is True
        assert window.a_jump_button.isEnabled() is True
        assert window.b_jump_button.isEnabled() is False
        window.ref_jump_button.click()
        assert len(jumps) == 1
        assert jumps[0][0].endswith("ref.xlsx")
        assert jumps[0][1] == "MAIN"
        assert jumps[0][2] == "D10"
        window.a_jump_button.click()
        assert jumps[1][2] == "E11"

        window._move(1)
        qapp.processEvents()
        assert window.ref_jump_button.isEnabled() is True
        assert window.a_jump_button.isEnabled() is False
        window.a_jump_button.click()
        assert len(jumps) == 2
    finally:
        window.close()
        window.deleteLater()


def test_main_enlarge_jump_follows_hyperlink_mode(qapp, tmp_path, monkeypatch):
    import ui.dialogs as dialogs_module

    jumps = []
    captured = {}

    class FakeViewer:
        def __init__(self, *args, **kwargs):
            captured["kwargs"] = kwargs
            captured["callback"] = kwargs.get("jump_callback")

        def exec(self):
            return 0

    monkeypatch.setattr(main_window_module, "ImageViewerDialog", FakeViewer)
    monkeypatch.setattr(
        dialogs_module,
        "jump_to_excel_cell",
        lambda path, sheet, cell, error_cb=None: jumps.append((path, sheet, cell)),
    )
    ref_xlsx = tmp_path / "ref.xlsx"
    a_xlsx = tmp_path / "a.xlsx"
    ref_xlsx.write_bytes(b"xlsx")
    a_xlsx.write_bytes(b"xlsx")
    ref_path = _png(tmp_path / "ref.png", (200, 20, 20))
    a_path = _png(tmp_path / "a.png", (20, 200, 20))
    item = InspectionItem(
        1,
        1,
        _extracted(ref_path, sheet_index=1, sheet_name="MAIN", cell="A1"),
        _extracted(a_path, sheet_index=1, sheet_name="MAIN", cell="B2"),
    )
    window = _window(tmp_path, monkeypatch)
    try:
        window.show()
        window.state.set_items(
            {1: [item]},
            mode="double",
            workbook_paths={"ref": str(ref_xlsx), "a": str(a_xlsx)},
        )
        window._populate_sheet_combo()
        window._render_current()
        window._enlarge("ref")
        assert captured["kwargs"]["jump_enabled"] is False
        captured["callback"]()
        assert jumps == []

        window.hyperlink_check.setChecked(True)
        window._enlarge("a")
        assert captured["kwargs"]["jump_enabled"] is True
        captured["callback"]()
        assert len(jumps) == 1
        assert jumps[0][0].endswith("a.xlsx")
        assert jumps[0][2] == "B2"
    finally:
        window.close()
        window.deleteLater()


def test_main_and_list_hyperlink_checkboxes_stay_in_sync(
    qapp, tmp_path, monkeypatch
):
    import ui.dialogs as dialogs_module

    settings_path = tmp_path / "settings.ini"
    monkeypatch.setattr(
        main_window_module,
        "QSettings",
        lambda *_args: QSettings(str(settings_path), QSettings.Format.IniFormat),
    )
    monkeypatch.setattr(
        dialogs_module,
        "QSettings",
        lambda *_args: QSettings(str(settings_path), QSettings.Format.IniFormat),
    )
    ref_path = _png(tmp_path / "ref.png", (200, 20, 20))
    a_path = _png(tmp_path / "a.png", (20, 200, 20))
    item = InspectionItem(
        1,
        1,
        _extracted(ref_path, sheet_index=1, sheet_name="MAIN", cell="A1"),
        _extracted(a_path, sheet_index=1, sheet_name="MAIN", cell="A1"),
    )
    window = main_window_module.MainWindow()
    try:
        window.show()
        window.state.set_items(
            {1: [item]},
            mode="double",
            workbook_paths={"ref": "ref.xlsx", "a": "a.xlsx"},
        )
        window._populate_sheet_combo()
        window._render_current()
        window._open_image_list_window()
        qapp.processEvents()
        list_window = window.image_list_window
        assert list_window is not None
        assert window.hyperlink_check.isChecked() is False
        assert list_window.hyperlink_check.isChecked() is False

        window.hyperlink_check.setChecked(True)
        qapp.processEvents()
        assert list_window.hyperlink_check.isChecked() is True
        assert list_window.ref_jump_button.isEnabled() is True
        assert window.ref_jump_button.isEnabled() is True

        list_window.hyperlink_check.setChecked(False)
        qapp.processEvents()
        assert window.hyperlink_check.isChecked() is False
        assert window.ref_jump_button.isEnabled() is False
        assert list_window.ref_jump_button.isEnabled() is False
    finally:
        window._close_image_list_window()
        qapp.processEvents()
        window.close()
        window.deleteLater()


def test_main_hyperlink_mode_persists(qapp, tmp_path, monkeypatch):
    window = _window(tmp_path, monkeypatch)
    try:
        window.show()
        assert window.hyperlink_check.isChecked() is False
        window.hyperlink_check.setChecked(True)
        qapp.processEvents()
        window.settings.sync()
    finally:
        window.close()
        window.deleteLater()

    second = _window(tmp_path, monkeypatch)
    try:
        assert second.hyperlink_check.isChecked() is True
    finally:
        second.close()
        second.deleteLater()


def test_quadra_main_preview_jump_buttons(qapp, tmp_path, monkeypatch):
    import ui.dialogs as dialogs_module

    jumps = []
    monkeypatch.setattr(
        dialogs_module,
        "jump_to_excel_cell",
        lambda path, sheet, cell, error_cb=None: jumps.append((path, sheet, cell)),
    )
    paths = {
        "ref": tmp_path / "ref.xlsx",
        "a": tmp_path / "a.xlsx",
        "b": tmp_path / "b.xlsx",
        "c": tmp_path / "c.xlsx",
    }
    for path in paths.values():
        path.write_bytes(b"xlsx")
    item = InspectionItem(
        1,
        1,
        _extracted(_png(tmp_path / "ref.png", (200, 20, 20)), sheet_index=1, sheet_name="MAIN", cell="A1"),
        _extracted(_png(tmp_path / "a.png", (20, 200, 20)), sheet_index=1, sheet_name="MAIN", cell="B1"),
        _extracted(_png(tmp_path / "b.png", (20, 20, 200)), sheet_index=1, sheet_name="MAIN", cell="C1"),
        _extracted(_png(tmp_path / "c.png", (200, 200, 20)), sheet_index=1, sheet_name="MAIN", cell="D1"),
    )
    window = _window(tmp_path, monkeypatch)
    try:
        window.show()
        window.quadra_check.setChecked(True)
        window.state.set_items(
            {1: [item]},
            mode="quadra",
            workbook_paths={key: str(path) for key, path in paths.items()},
        )
        window._populate_sheet_combo()
        window._render_current()
        window.hyperlink_check.setChecked(True)
        qapp.processEvents()
        assert window.b_jump_button.isEnabled() is True
        assert window.c_jump_button.isEnabled() is True
        window.b_jump_button.click()
        window.c_jump_button.click()
        assert [item[2] for item in jumps] == ["C1", "D1"]
        assert jumps[0][0].endswith("b.xlsx")
        assert jumps[1][0].endswith("c.xlsx")
    finally:
        window.close()
        window.deleteLater()


def test_main_preview_header_keeps_jump_button_at_right(
    qapp, tmp_path, monkeypatch
):
    from tests.unit.test_preview_pane import _assert_jump_button_at_right_edge

    window = _window(tmp_path, monkeypatch)
    try:
        window.resize(1100, 800)
        window.show()
        qapp.processEvents()
        window.ref_meta.setText("SC CLKx 64 case · D137 · 시트 내 137/156")
        _assert_jump_button_at_right_edge(
            window.ref_container,
            window.ref_meta,
            window.ref_jump_button,
            qapp,
        )
    finally:
        window.close()
        window.deleteLater()
