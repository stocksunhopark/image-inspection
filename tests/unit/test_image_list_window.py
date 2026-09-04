"""ImageListWindow table, keyboard navigation, and MainWindow integration."""

from __future__ import annotations

from pathlib import Path
from typing import Dict

import pytest
from PIL import Image
from PyQt6.QtCore import QSettings, QTimer, Qt
from PyQt6.QtTest import QTest
from PyQt6.QtWidgets import QApplication

import ui.dialogs as dialogs_module
import ui.main_window as main_window_module
from models import ExtractedImage, InspectionItem
from ui.dialogs import ImageListWindow


@pytest.fixture(scope="module")
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication(["image-list-window-test", "-platform", "offscreen"])
    return app


def _real_image(
    tmp_path: Path,
    role: str,
    sheet_index: int,
    cell: str,
    color: tuple[int, int, int],
) -> ExtractedImage:
    path = tmp_path / f"{role}-s{sheet_index}-{cell}.png"
    Image.new("RGB", (24, 16), color).save(path, format="PNG")
    row = int("".join(character for character in cell if character.isdigit())) - 1
    column = ord(cell[0].upper()) - ord("A")
    return ExtractedImage(
        sheet_index=sheet_index,
        sheet_name=f"Sheet {sheet_index}",
        cell_address=cell,
        merged_range="",
        anchor_row=row,
        anchor_col=column,
        source_path=str(path),
        preview_path=str(path),
    )


def _null_image(sheet_index: int, cell: str) -> ExtractedImage:
    row = int("".join(character for character in cell if character.isdigit())) - 1
    column = ord(cell[0].upper()) - ord("A")
    return ExtractedImage(
        sheet_index=sheet_index,
        sheet_name=f"Sheet {sheet_index}",
        cell_address=cell,
        merged_range="",
        anchor_row=row,
        anchor_col=column,
        is_null=True,
    )


@pytest.fixture
def sample_items(tmp_path: Path):
    refs = [
        _real_image(tmp_path, "ref-1", 1, "A1", (220, 20, 20)),
        _real_image(tmp_path, "ref-2", 1, "A2", (20, 20, 220)),
        _real_image(tmp_path, "ref-3", 2, "B1", (220, 200, 20)),
    ]
    compares_a = [
        _real_image(tmp_path, "a-1", 1, "A1", (20, 200, 40)),
        _null_image(1, "A2"),
        _real_image(tmp_path, "a-3", 2, "B1", (150, 40, 180)),
    ]
    compares_b = [
        _real_image(tmp_path, "b-1", 1, "A1", (20, 190, 190)),
        _null_image(1, "A2"),
        _real_image(tmp_path, "b-3", 2, "B1", (210, 40, 150)),
    ]
    compares_c = [
        _real_image(tmp_path, "c-1", 1, "A1", (240, 180, 20)),
        _null_image(1, "A2"),
        _real_image(tmp_path, "c-3", 2, "B1", (40, 220, 220)),
    ]

    double_items = [
        InspectionItem(
            sheet_index=ref.sheet_index,
            index=(1, 2, 1)[position],
            image_ref=ref,
            image_a=compares_a[position],
        )
        for position, ref in enumerate(refs)
    ]
    triple_items = [
        InspectionItem(
            sheet_index=item.sheet_index,
            index=item.index,
            image_ref=item.image_ref,
            image_a=item.image_a,
            image_b=compares_b[position],
        )
        for position, item in enumerate(double_items)
    ]
    quadra_items = [
        InspectionItem(
            sheet_index=item.sheet_index,
            index=item.index,
            image_ref=item.image_ref,
            image_a=item.image_a,
            image_b=item.image_b,
            image_c=compares_c[position],
        )
        for position, item in enumerate(triple_items)
    ]
    paths = {
        "ref": str(tmp_path / "reference.xlsx"),
        "a": str(tmp_path / "compare-a.xlsx"),
        "b": str(tmp_path / "compare-b.xlsx"),
        "c": str(tmp_path / "compare-c.xlsx"),
    }
    return {
        "double": double_items,
        "triple": triple_items,
        "quadra": quadra_items,
        "paths": paths,
    }


def _source_rgb(label) -> tuple[int, int, int]:
    assert not label._source.isNull()
    color = label._source.toImage().pixelColor(0, 0)
    return color.red(), color.green(), color.blue()


def _close_window(window: ImageListWindow, qapp: QApplication) -> None:
    window.close()
    qapp.processEvents()


def test_double_table_initial_index_and_keyboard_update_preview_immediately(
    qapp,
    sample_items,
):
    window = ImageListWindow(
        sample_items["double"],
        "double",
        sample_items["paths"],
        initial_index=1,
    )
    try:
        assert window.isModal() is False
        assert window.windowModality() == Qt.WindowModality.NonModal
        assert bool(
            window.windowFlags() & Qt.WindowType.WindowMaximizeButtonHint
        )
        window.show()
        qapp.processEvents()
        assert window.isVisible()

        assert window.table.rowCount() == 3
        assert window.table.columnCount() == 5
        assert [
            window.table.horizontalHeaderItem(column).text()
            for column in range(window.table.columnCount())
        ] == ["전체", "시트", "No.", "Ref 셀", "비교A 셀"]
        assert window.table.item(1, 4).text() == "이미지 없음"
        assert window.table.currentRow() == 1
        assert window.position_label.text() == "전체 2 / 3"
        assert "A2" in window.ref_meta.text()
        assert _source_rgb(window.ref_image) == (20, 20, 220)
        assert "이미지 없음" in window.a_meta.text()
        assert window.a_image.text() == "해당 위치에 이미지가 없습니다."
        assert window.a_image._source.isNull()
        assert window.b_container.isHidden()

        # QTest delivers the key synchronously: row, metadata and pixmap must
        # all represent the new item before keyClick returns.
        QTest.keyClick(window.table, Qt.Key.Key_Up)
        assert window.table.currentRow() == 0
        assert window.position_label.text() == "전체 1 / 3"
        assert "A1" in window.ref_meta.text()
        assert _source_rgb(window.ref_image) == (220, 20, 20)
        assert "A1" in window.a_meta.text()
        assert _source_rgb(window.a_image) == (20, 200, 40)

        # The first-row boundary does not wrap.
        QTest.keyClick(window.table, Qt.Key.Key_Up)
        assert window.table.currentRow() == 0
        assert _source_rgb(window.ref_image) == (220, 20, 20)

        QTest.keyClick(window.table, Qt.Key.Key_Down)
        assert window.table.currentRow() == 1
        assert "이미지 없음" in window.a_meta.text()
        assert window.a_image._source.isNull()
        QTest.keyClick(window.table, Qt.Key.Key_Down)
        assert window.table.currentRow() == 2
        assert window.position_label.text() == "전체 3 / 3"
        assert "B1" in window.ref_meta.text()
        assert _source_rgb(window.ref_image) == (220, 200, 20)
        assert _source_rgb(window.a_image) == (150, 40, 180)

        # The last-row boundary does not wrap.
        QTest.keyClick(window.table, Qt.Key.Key_Down)
        assert window.table.currentRow() == 2
        assert _source_rgb(window.ref_image) == (220, 200, 20)
    finally:
        _close_window(window, qapp)


def test_triple_adds_visible_b_column_and_updates_null_and_real_b_preview(
    qapp,
    sample_items,
):
    window = ImageListWindow(
        sample_items["triple"],
        "triple",
        sample_items["paths"],
        initial_index=0,
    )
    try:
        window.show()
        qapp.processEvents()

        assert window.table.rowCount() == 3
        assert window.table.columnCount() == 6
        assert window.table.horizontalHeaderItem(5).text() == "비교B 셀"
        assert window.table.item(1, 5).text() == "이미지 없음"
        assert not window.b_container.isHidden()
        assert window.b_container.isVisible()
        assert "A1" in window.b_meta.text()
        assert _source_rgb(window.b_image) == (20, 190, 190)

        QTest.keyClick(window.table, Qt.Key.Key_Down)
        assert window.table.currentRow() == 1
        assert "이미지 없음" in window.b_meta.text()
        assert window.b_image.text() == "해당 위치에 이미지가 없습니다."
        assert window.b_image._source.isNull()

        QTest.keyClick(window.table, Qt.Key.Key_Down)
        assert window.table.currentRow() == 2
        assert "B1" in window.b_meta.text()
        assert _source_rgb(window.b_image) == (210, 40, 150)
    finally:
        _close_window(window, qapp)


def test_quadra_adds_visible_c_column_and_preview(
    qapp,
    sample_items,
):
    window = ImageListWindow(
        sample_items["quadra"],
        "quadra",
        sample_items["paths"],
        initial_index=0,
    )
    try:
        window.show()
        qapp.processEvents()

        assert window.table.columnCount() == 7
        assert window.table.horizontalHeaderItem(6).text() == "비교C 셀"
        assert window.table.item(1, 6).text() == "이미지 없음"
        assert not window.b_container.isHidden()
        assert not window.c_container.isHidden()
        assert _source_rgb(window.c_image) == (240, 180, 20)

        QTest.keyClick(window.table, Qt.Key.Key_Down)
        assert window.c_image.text() == "해당 위치에 이미지가 없습니다."
    finally:
        _close_window(window, qapp)


def test_list_column_plan_gives_half_leftover_to_visible_cell_columns():
    double = ImageListWindow.list_column_widths(5, 1000)
    triple = ImageListWindow.list_column_widths(6, 1000)
    quadra = ImageListWindow.list_column_widths(7, 1000)

    assert double[0] == ImageListWindow._INDEX_COLUMN_WIDTH
    assert double[2] == ImageListWindow._NO_COLUMN_WIDTH
    leftover_double = 1000 - (
        ImageListWindow._INDEX_COLUMN_WIDTH
        + ImageListWindow._NO_COLUMN_WIDTH
        + 2 * ImageListWindow._CELL_COLUMN_WIDTH
    )
    assert double[1] == leftover_double // 2
    assert sum(double[3:]) == 2 * ImageListWindow._CELL_COLUMN_WIDTH + (
        leftover_double - leftover_double // 2
    )
    assert max(double[3:]) - min(double[3:]) <= 1

    leftover_triple = 1000 - (
        ImageListWindow._INDEX_COLUMN_WIDTH
        + ImageListWindow._NO_COLUMN_WIDTH
        + 3 * ImageListWindow._CELL_COLUMN_WIDTH
    )
    assert triple[1] == leftover_triple // 2
    assert max(triple[3:]) - min(triple[3:]) <= 1

    leftover_quadra = 1000 - (
        ImageListWindow._INDEX_COLUMN_WIDTH
        + ImageListWindow._NO_COLUMN_WIDTH
        + 4 * ImageListWindow._CELL_COLUMN_WIDTH
    )
    assert quadra[1] == leftover_quadra // 2
    assert max(quadra[3:]) - min(quadra[3:]) <= 1


def test_sheet_half_leftover_is_shared_by_visible_cell_columns_in_all_modes(
    qapp,
    sample_items,
):
    for mode in ("double", "triple", "quadra"):
        window = ImageListWindow(
            sample_items[mode],
            mode,
            sample_items["paths"],
            initial_index=0,
        )
        try:
            window.resize(1550, 880)
            window.show()
            qapp.processEvents()
            plan = ImageListWindow.list_column_widths(
                window.table.columnCount(),
                window.table.viewport().width(),
            )
            actual = [
                window.table.columnWidth(column)
                for column in range(window.table.columnCount())
            ]
            assert actual == plan
            assert max(actual[3:]) - min(actual[3:]) <= 1
            assert actual[3] > ImageListWindow._CELL_COLUMN_WIDTH
        finally:
            _close_window(window, qapp)


def test_sheet_selector_filters_rows_and_keeps_arrow_navigation_in_sheet(
    qapp,
    sample_items,
):
    window = ImageListWindow(
        sample_items["double"],
        "double",
        sample_items["paths"],
        initial_index=0,
    )
    try:
        window.show()
        qapp.processEvents()

        assert window.sheet_combo.count() == 3
        assert window.sheet_combo.itemData(0) is None
        assert "전체 시트 (3개)" == window.sheet_combo.itemText(0)
        assert window.sheet_combo.itemData(1) == 1
        assert "(2개)" in window.sheet_combo.itemText(1)
        assert window.sheet_combo.itemData(2) == 2
        assert "(1개)" in window.sheet_combo.itemText(2)

        window.sheet_combo.setCurrentIndex(2)
        qapp.processEvents()
        assert window.table.rowCount() == 1
        assert window.table.currentRow() == 0
        assert "Sheet 2" in window.table.item(0, 1).text()
        assert window.table.item(0, 0).text() == "3"
        assert "시트 1 / 1" in window.position_label.text()
        assert "전체 3 / 3" in window.position_label.text()
        assert "B1" in window.ref_meta.text()

        # 선택한 시트의 첫/마지막 경계에서도 화살표가 다른 시트로 넘어가지 않는다.
        QTest.keyClick(window.table, Qt.Key.Key_Up)
        QTest.keyClick(window.table, Qt.Key.Key_Down)
        assert window.table.currentRow() == 0

        window.sheet_combo.setCurrentIndex(1)
        qapp.processEvents()
        assert window.table.rowCount() == 2
        assert window.table.currentRow() == 0
        QTest.keyClick(window.table, Qt.Key.Key_Down)
        assert window.table.currentRow() == 1
        assert "A2" in window.ref_meta.text()

        # 메인 창의 전체 index 동기화 요청이 현재 필터 밖이면 전체 보기로 복귀한다.
        window.set_current_index(2)
        assert window.sheet_combo.currentIndex() == 0
        assert window.table.rowCount() == 3
        assert window.table.currentRow() == 2
        assert "B1" in window.ref_meta.text()
    finally:
        _close_window(window, qapp)


def test_main_window_enables_list_button_and_reuses_one_nonmodal_window(
    qapp,
    tmp_path: Path,
    monkeypatch,
    sample_items,
):
    settings_path = tmp_path / "settings.ini"
    monkeypatch.setattr(
        main_window_module,
        "QSettings",
        lambda *_args: QSettings(str(settings_path), QSettings.Format.IniFormat),
    )
    main = main_window_module.MainWindow()
    try:
        main.show()
        qapp.processEvents()
        assert not main.list_button.isEnabled()
        assert main.image_list_window is None

        double_items = sample_items["double"]
        main.state.set_items(
            {1: double_items[:2], 2: double_items[2:]},
            mode="double",
            workbook_paths={
                "ref": sample_items["paths"]["ref"],
                "a": sample_items["paths"]["a"],
            },
        )
        main._populate_sheet_combo()
        main._render_current()
        assert main.list_button.isEnabled()

        QTest.mouseClick(main.list_button, Qt.MouseButton.LeftButton)
        qapp.processEvents()
        first_window = main.image_list_window
        assert first_window is not None
        assert first_window.isVisible()
        assert first_window.isModal() is False
        assert first_window.table.currentRow() == 0

        main.state.current_index = 2
        main._open_image_list_window()
        qapp.processEvents()
        assert main.image_list_window is first_window
        assert first_window.table.currentRow() == 2
        assert first_window.position_label.text() == "전체 3 / 3"
    finally:
        main._close_image_list_window()
        qapp.processEvents()
        main.close()
        main.deleteLater()
        qapp.processEvents()


def test_closing_list_while_image_viewer_is_open_does_not_touch_deleted_table(
    qapp,
    sample_items,
):
    window = ImageListWindow(
        sample_items["double"],
        "double",
        sample_items["paths"],
        initial_index=0,
    )
    window.show()
    qapp.processEvents()

    # _enlarge()의 modal exec 중에도 새 로드 성공/메인 종료 signal은 처리된다.
    # 이때 부모 목록창이 먼저 닫힌 뒤 삭제된 table에 focus를 주면 안 된다.
    QTimer.singleShot(0, window.prepare_for_close)
    window.ref_image.clicked.emit()
    qapp.processEvents()


def _patch_list_settings(tmp_path: Path, monkeypatch):
    settings_path = tmp_path / "list-settings.ini"
    monkeypatch.setattr(
        dialogs_module,
        "QSettings",
        lambda *_args: QSettings(str(settings_path), QSettings.Format.IniFormat),
    )


def test_hyperlink_mode_is_off_by_default_and_only_double_click_jumps(
    qapp,
    tmp_path: Path,
    monkeypatch,
    sample_items,
):
    _patch_list_settings(tmp_path, monkeypatch)
    jumps = []
    monkeypatch.setattr(
        dialogs_module,
        "jump_to_excel_cell",
        lambda path, sheet, cell, error_cb=None: jumps.append((path, sheet, cell)),
    )
    for path in sample_items["paths"].values():
        Path(path).write_bytes(b"xlsx")

    window = ImageListWindow(
        sample_items["double"],
        "double",
        sample_items["paths"],
        initial_index=0,
    )
    try:
        window.show()
        qapp.processEvents()
        assert window.hyperlink_check.isChecked() is False
        assert window.sheet_combo.width() > 0
        assert window.hyperlink_check.text() == "하이퍼링크 모드"

        window._on_cell_double_clicked(0, 3)
        assert jumps == []

        QTest.keyClick(window.table, Qt.Key.Key_Down)
        assert window.table.currentRow() == 1
        assert jumps == []

        window.hyperlink_check.setChecked(True)
        window._on_cell_double_clicked(0, 0)
        window._on_cell_double_clicked(0, 1)
        window._on_cell_double_clicked(0, 2)
        assert jumps == []

        window._on_cell_double_clicked(0, 3)
        assert len(jumps) == 1
        assert jumps[0][1] == "Sheet 1"
        assert jumps[0][2] == "A1"
        assert jumps[0][0].endswith("reference.xlsx")

        window._on_cell_double_clicked(0, 4)
        assert len(jumps) == 2
        assert jumps[1][2] == "A1"
        assert jumps[1][0].endswith("compare-a.xlsx")

        window._on_cell_double_clicked(1, 4)
        assert len(jumps) == 2
    finally:
        _close_window(window, qapp)


def test_quadra_hyperlink_double_click_uses_b_and_c_columns(
    qapp,
    tmp_path: Path,
    monkeypatch,
    sample_items,
):
    _patch_list_settings(tmp_path, monkeypatch)
    jumps = []
    monkeypatch.setattr(
        dialogs_module,
        "jump_to_excel_cell",
        lambda path, sheet, cell, error_cb=None: jumps.append((path, sheet, cell)),
    )
    for path in sample_items["paths"].values():
        Path(path).write_bytes(b"xlsx")

    window = ImageListWindow(
        sample_items["quadra"],
        "quadra",
        sample_items["paths"],
        initial_index=0,
    )
    try:
        window.hyperlink_check.setChecked(True)
        window._on_cell_double_clicked(0, 5)
        window._on_cell_double_clicked(0, 6)
        window._on_cell_double_clicked(1, 6)
        assert [item[2] for item in jumps] == ["A1", "A1"]
        assert jumps[0][0].endswith("compare-b.xlsx")
        assert jumps[1][0].endswith("compare-c.xlsx")
    finally:
        _close_window(window, qapp)


def test_hyperlink_mode_persists_and_respects_sheet_filter_and_missing_file(
    qapp,
    tmp_path: Path,
    monkeypatch,
    sample_items,
):
    _patch_list_settings(tmp_path, monkeypatch)
    jumps = []
    warnings = []
    monkeypatch.setattr(
        dialogs_module,
        "jump_to_excel_cell",
        lambda path, sheet, cell, error_cb=None: jumps.append((path, sheet, cell)),
    )
    monkeypatch.setattr(
        dialogs_module.QMessageBox,
        "warning",
        lambda *args, **kwargs: warnings.append((args, kwargs)),
    )

    first = ImageListWindow(
        sample_items["double"],
        "double",
        sample_items["paths"],
        initial_index=0,
    )
    try:
        first.hyperlink_check.setChecked(True)
        first.sheet_combo.setCurrentIndex(2)
        qapp.processEvents()
        first._on_cell_double_clicked(0, 3)
        assert jumps == []
        assert warnings
        assert "reference.xlsx" in str(warnings[0])
    finally:
        _close_window(first, qapp)

    for path in sample_items["paths"].values():
        Path(path).write_bytes(b"xlsx")

    second = ImageListWindow(
        sample_items["double"],
        "double",
        sample_items["paths"],
        initial_index=0,
    )
    try:
        assert second.hyperlink_check.isChecked() is True
        second.sheet_combo.setCurrentIndex(2)
        qapp.processEvents()
        assert second.table.rowCount() == 1
        second._on_cell_double_clicked(0, 3)
        assert len(jumps) == 1
        assert jumps[0][1] == "Sheet 2"
        assert jumps[0][2] == "B1"
    finally:
        _close_window(second, qapp)
