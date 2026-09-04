"""Double/Triple 육안 검사 메인 화면의 탐색 smoke test."""

from pathlib import Path

from PIL import Image
from PyQt6.QtCore import QSettings

import ui.main_window as main_window_module
from models import ExtractedImage, InspectionItem


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
        window.triple_radio.setChecked(True)
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

        window.double_radio.setChecked(True)
        assert "현재 화면은 이전 Triple 결과" in window.status_label.text()
        assert not window.b_container.isHidden()
    finally:
        window.close()
        window.deleteLater()
