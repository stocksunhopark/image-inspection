"""Double/Triple 육안 검사 메인 화면의 탐색 smoke test."""

from pathlib import Path

from PIL import Image
from PyQt6.QtCore import QSettings
from PyQt6.QtWidgets import QMessageBox

import ui.main_window as main_window_module
from models import MISSING_SHEET, ExtractedImage, InspectionItem, SheetInfo
from ui.dialogs import LoadSummaryDialog


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


def test_main_window_closes_only_after_exit_confirmation(
    qapp, tmp_path, monkeypatch
):
    window = _window(tmp_path, monkeypatch)
    calls = []

    def answer_no(parent, title, message, buttons, default_button):
        calls.append((parent, title, message, buttons, default_button))
        return QMessageBox.StandardButton.No

    monkeypatch.setattr(main_window_module.QMessageBox, "question", answer_no)
    try:
        window.show()
        qapp.processEvents()

        assert window.close() is False
        assert window.isVisible()
        assert len(calls) == 1
        assert calls[0][0] is window
        assert calls[0][1] == "프로그램 종료"
        assert calls[0][2] == "프로그램을 종료 하시겠습니까? (즐거운 하루 되세요)"
        assert calls[0][3] == (
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        assert calls[0][4] == QMessageBox.StandardButton.No

        monkeypatch.setattr(
            main_window_module.QMessageBox,
            "question",
            lambda *_args, **_kwargs: QMessageBox.StandardButton.Yes,
        )
        assert window.close() is True
        assert not window.isVisible()
    finally:
        window._close_confirmed = True
        window.close()
        window.deleteLater()


def test_load_summary_dialog_uses_fixed_scrollable_text_area(qapp):
    message = "\n".join(f"확인 위치 {index}" for index in range(200))
    dialog = LoadSummaryDialog(message, has_warnings=True)
    try:
        dialog.show()
        qapp.processEvents()

        assert dialog.windowTitle() == "불러오기 완료 · 확인 필요"
        assert dialog.height() <= 480
        assert dialog.text.isReadOnly()
        assert dialog.text.verticalScrollBar().maximum() > 0
    finally:
        dialog.close()
        dialog.deleteLater()


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
        assert all(not widget.isHidden() for widget in window._file_input_rows["b"])
        assert all(not widget.isHidden() for widget in window._file_input_rows["c"])
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
        assert all(widget.isHidden() for widget in window._file_input_rows["b"])
        assert all(widget.isHidden() for widget in window._file_input_rows["c"])

        window.triple_check.setChecked(True)
        assert window.triple_check.isChecked()
        assert not window.double_check.isChecked()
        assert not window.quadra_check.isChecked()
        assert all(not widget.isHidden() for widget in window._file_input_rows["b"])
        assert all(widget.isHidden() for widget in window._file_input_rows["c"])

        window.quadra_check.setChecked(True)
        assert window.quadra_check.isChecked()
        assert not window.double_check.isChecked()
        assert not window.triple_check.isChecked()
        assert all(not widget.isHidden() for widget in window._file_input_rows["b"])
        assert all(not widget.isHidden() for widget in window._file_input_rows["c"])

        window.quadra_check.setChecked(False)
        assert window.quadra_check.isChecked()
        assert not window.double_check.isChecked()
        assert not window.triple_check.isChecked()
        assert window._selected_mode() == "quadra"
    finally:
        window.close()
        window.deleteLater()


def test_each_excel_row_has_unclipped_path_reset_before_file_select(
    qapp, tmp_path, monkeypatch
):
    window = _window(tmp_path, monkeypatch)
    try:
        window.resize(900, 650)
        window.quadra_check.setChecked(True)
        window.show()
        qapp.processEvents()

        rows = (
            (
                window.file_ref_label,
                window.file_ref_edit,
                window.file_ref_reset_button,
                window.file_ref_button,
            ),
            (
                window.file_a_label,
                window.file_a_edit,
                window.file_a_reset_button,
                window.file_a_button,
            ),
            (
                window.file_b_label,
                window.file_b_edit,
                window.file_b_reset_button,
                window.file_b_button,
            ),
            (
                window.file_c_label,
                window.file_c_edit,
                window.file_c_reset_button,
                window.file_c_button,
            ),
        )
        assert [row[0].text() for row in rows] == [
            "Excel Ref",
            "Excel A",
            "Excel B",
            "Excel C",
        ]
        expected_columns = [
            (widget.geometry().left(), widget.geometry().right(), widget.width())
            for widget in rows[0]
        ]
        for label, edit, reset_button, select_button in rows:
            assert reset_button.text() == "경로 reset"
            assert select_button.text() == "파일 선택"
            assert reset_button.objectName() == "pathResetButton"
            assert reset_button.parentWidget() is select_button.parentWidget()
            assert reset_button.geometry().right() < select_button.geometry().left()
            assert label.width() == window._file_input_label_width
            assert reset_button.width() == window._file_input_reset_width
            assert select_button.width() == window._file_input_select_width
            assert reset_button.width() >= reset_button.sizeHint().width()
            assert select_button.width() >= select_button.sizeHint().width()
            assert edit.width() > 100
            assert [
                (widget.geometry().left(), widget.geometry().right(), widget.width())
                for widget in (label, edit, reset_button, select_button)
            ] == expected_columns
    finally:
        window.close()
        window.deleteLater()


def test_path_reset_clears_only_its_row_and_saved_value(
    qapp, tmp_path, monkeypatch
):
    window = _window(tmp_path, monkeypatch)
    try:
        window.quadra_check.setChecked(True)
        rows = {
            "path_ref": (window.file_ref_edit, window.file_ref_reset_button),
            "path_a": (window.file_a_edit, window.file_a_reset_button),
            "path_b": (window.file_b_edit, window.file_b_reset_button),
            "path_c": (window.file_c_edit, window.file_c_reset_button),
        }

        for reset_key, (_reset_edit, reset_button) in rows.items():
            expected = {}
            for key, (edit, _button) in rows.items():
                value = f"C:/waveforms/{key}.xlsx"
                edit.setText(value)
                window.settings.setValue(key, value)
                expected[key] = value

            reset_button.click()

            for key, (edit, _button) in rows.items():
                if key == reset_key:
                    assert edit.text() == ""
                    assert window.settings.value(key, "", type=str) == ""
                else:
                    assert edit.text() == expected[key]
                    assert window.settings.value(key, "", type=str) == expected[key]
            assert window.settings.value("last_directory", "", type=str) == "D:\\"
    finally:
        window.close()
        window.deleteLater()


def test_file_select_starts_at_d_drive_after_path_reset(
    qapp, tmp_path, monkeypatch
):
    window = _window(tmp_path, monkeypatch)
    dialog_calls = []

    def fake_get_open_file_name(parent, title, start, file_filter):
        dialog_calls.append((parent, title, start, file_filter))
        return "", ""

    monkeypatch.setattr(
        main_window_module.QFileDialog,
        "getOpenFileName",
        fake_get_open_file_name,
    )
    try:
        window.file_ref_edit.setText("C:/previous/location/ref.xlsx")
        window.settings.setValue("last_directory", "C:/previous/location")

        window.file_ref_reset_button.click()
        window.file_ref_button.click()

        assert len(dialog_calls) == 1
        assert dialog_calls[0][0] is window
        assert dialog_calls[0][1] == "Excel 파일 선택"
        assert dialog_calls[0][2] == "D:\\"
        assert dialog_calls[0][3] == "Excel 통합 문서 (*.xlsx *.xlsm)"
    finally:
        window.close()
        window.deleteLater()


def test_main_has_sheet_order_button_next_to_loading_controls(
    qapp, tmp_path, monkeypatch
):
    window = _window(tmp_path, monkeypatch)
    try:
        window.show()
        qapp.processEvents()
        assert window.load_button.text() == "이미지 불러오기"
        assert window.cancel_button.text() == "취소"
        assert window.sheet_mapping_button.text() == "시트순서설정"
        assert window.input_action_row.spacing() == 4
        assert window.sheet_mapping_button.isEnabled()

        window._set_loading(True)
        assert not window.sheet_mapping_button.isEnabled()
        assert not window.file_ref_reset_button.isEnabled()
        assert not window.file_a_reset_button.isEnabled()
        assert not window.file_b_reset_button.isEnabled()
        assert not window.file_c_reset_button.isEnabled()
        window._set_loading(False)
        assert window.sheet_mapping_button.isEnabled()
        assert window.file_ref_reset_button.isEnabled()
        assert window.file_a_reset_button.isEnabled()
        assert window.file_b_reset_button.isEnabled()
        assert window.file_c_reset_button.isEnabled()
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
    assert "위쪽 7개 행" in USAGE_HELP_TEXT
    assert "오른쪽 끝" in USAGE_HELP_TEXT
    assert "PASS/FAIL" in USAGE_HELP_TEXT
    assert "Ctrl+Enter" in USAGE_HELP_TEXT
    assert "시트순서설정" in USAGE_HELP_TEXT
    assert "경로 reset" in USAGE_HELP_TEXT
    assert "자동 매핑 복원" in USAGE_HELP_TEXT
    assert "비교 그룹 추가" in USAGE_HELP_TEXT
    assert "다른파형 함께보기" in USAGE_HELP_TEXT
    assert "모든 파형이 같은 비율 위치로 함께 움직입니다" in USAGE_HELP_TEXT
    assert "Excel 이름 제목을 다른 이미지 칸으로 드래그" in USAGE_HELP_TEXT
    assert "함께보기와 메인 미리보기도 같은 배열" in USAGE_HELP_TEXT
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
        assert captured["kwargs"]["comparison_available"] is True
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


def test_list_preview_order_updates_main_reopens_and_resets_on_new_load(
    qapp, tmp_path, monkeypatch
):
    item = InspectionItem(
        1,
        1,
        _extracted(
            _png(tmp_path / "ref-order.png", (200, 20, 20)),
            sheet_index=1,
            sheet_name="MAIN",
            cell="A1",
        ),
        _extracted(
            _png(tmp_path / "a-order.png", (20, 200, 20)),
            sheet_index=1,
            sheet_name="MAIN",
            cell="B1",
        ),
        _extracted(
            _png(tmp_path / "b-order.png", (20, 20, 200)),
            sheet_index=1,
            sheet_name="MAIN",
            cell="C1",
        ),
        _extracted(
            _png(tmp_path / "c-order.png", (200, 200, 20)),
            sheet_index=1,
            sheet_name="MAIN",
            cell="D1",
        ),
    )
    paths = {role: f"{role}.xlsx" for role in ("ref", "a", "b", "c")}
    window = _window(tmp_path, monkeypatch)

    def main_physical_order():
        role_by_container = {
            id(container): role
            for role, container in window._preview_containers.items()
        }
        return tuple(
            role_by_container[id(window.preview_splitter.widget(index))]
            for index in range(4)
        )

    try:
        window.show()
        window.state.set_items({1: [item]}, mode="quadra", workbook_paths=paths)
        window._populate_sheet_combo()
        window._apply_preview_mode("quadra")
        window._render_current()
        window._open_image_list_window()
        qapp.processEvents()
        first_list = window.image_list_window
        assert first_list is not None

        first_list._swap_preview_roles("ref", "c")
        qapp.processEvents()
        expected = ("c", "a", "b", "ref")
        assert window.state.preview_order == expected
        assert main_physical_order() == expected
        assert first_list.preview_order == expected

        window._close_image_list_window()
        qapp.processEvents()
        window._open_image_list_window()
        qapp.processEvents()
        reopened = window.image_list_window
        assert reopened is not None
        assert reopened.preview_order == expected

        window._close_image_list_window()
        qapp.processEvents()
        window.state.set_items({1: [item]}, mode="quadra", workbook_paths=paths)
        window._apply_preview_mode("quadra")
        assert window.state.preview_order == ("ref", "a", "b", "c")
        assert main_physical_order() == ("ref", "a", "b", "c")
    finally:
        window._close_image_list_window()
        qapp.processEvents()
        window.close()
        window.deleteLater()


def test_main_comparison_uses_current_reordered_preview_layout(
    qapp, tmp_path, monkeypatch
):
    captured = {}

    class FakeViewer:
        def __init__(self, *_args, **_kwargs):
            pass

        def exec(self):
            return main_window_module.SHOW_COMPARISON_RESULT

    class FakeComparison:
        def __init__(self, _item, _mode, _paths, _parent, **kwargs):
            captured.update(kwargs)

        def exec_maximized(self):
            captured["maximized"] = True
            return 0

    monkeypatch.setattr(main_window_module, "ImageViewerDialog", FakeViewer)
    monkeypatch.setattr(main_window_module, "ImageComparisonDialog", FakeComparison)
    ref_path = _png(tmp_path / "ref-compare.png", (200, 20, 20))
    a_path = _png(tmp_path / "a-compare.png", (20, 200, 20))
    item = InspectionItem(
        1,
        1,
        _extracted(ref_path, sheet_index=1, sheet_name="MAIN", cell="A1"),
        _extracted(a_path, sheet_index=1, sheet_name="MAIN", cell="B1"),
    )
    window = _window(tmp_path, monkeypatch)
    try:
        window.state.set_items({1: [item]}, mode="double")
        window.state.set_preview_order(("a", "ref"))
        window._apply_preview_mode("double")
        window._render_current()

        window._enlarge("ref")

        assert captured["role_order"] == ("a", "ref")
        assert captured["maximized"] is True
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
