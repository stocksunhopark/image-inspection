"""동일 크기 letterbox preview와 상단 목록/하단 이미지 배치 테스트."""

from PIL import Image
import pytest
from PyQt6.QtCore import QPoint, QSettings, Qt
from PyQt6.QtTest import QTest

import ui.main_window as main_window_module
from models import ExtractedImage, InspectionItem
from ui.dialogs import ImageListWindow


def _extracted(size, color) -> ExtractedImage:
    return ExtractedImage(
        sheet_index=1,
        sheet_name="MAIN",
        cell_address="A1",
        merged_range="",
        anchor_row=0,
        anchor_col=0,
        image=Image.new("RGB", size, color),
    )


def _triple_item() -> InspectionItem:
    return InspectionItem(
        sheet_index=1,
        index=1,
        image_ref=_extracted((400, 100), (230, 30, 30)),
        image_a=_extracted((180, 180), (30, 220, 30)),
        image_b=_extracted((100, 400), (30, 30, 230)),
    )


def _display_size(label) -> tuple[float, float]:
    pixmap = label.pixmap()
    assert pixmap is not None and not pixmap.isNull()
    dpr = max(1.0, float(pixmap.devicePixelRatio()))
    return pixmap.width() / dpr, pixmap.height() / dpr


def _colored_center_ratio(label) -> float:
    """중앙 가로/세로선의 non-letterbox 길이 비율."""
    pixmap = label.pixmap()
    assert pixmap is not None
    image = pixmap.toImage()
    background = (2, 6, 23)

    def is_content(x: int, y: int) -> bool:
        color = image.pixelColor(x, y)
        rgb = (color.red(), color.green(), color.blue())
        return sum(abs(rgb[index] - background[index]) for index in range(3)) > 30

    center_y = image.height() // 2
    center_x = image.width() // 2
    content_width = sum(is_content(x, center_y) for x in range(image.width()))
    content_height = sum(is_content(center_x, y) for y in range(image.height()))
    assert content_width > 0 and content_height > 0
    return content_width / content_height


def _assert_equal_letterbox_previews(window) -> None:
    labels = (window.ref_image, window.a_image, window.b_image)
    label_sizes = [(label.width(), label.height()) for label in labels]
    assert max(width for width, _height in label_sizes) - min(
        width for width, _height in label_sizes
    ) <= 1
    assert max(height for _width, height in label_sizes) - min(
        height for _width, height in label_sizes
    ) <= 1

    pixmap_sizes = [_display_size(label) for label in labels]
    assert max(width for width, _height in pixmap_sizes) - min(
        width for width, _height in pixmap_sizes
    ) <= 1
    assert max(height for _width, height in pixmap_sizes) - min(
        height for _width, height in pixmap_sizes
    ) <= 1

    ratios = [_colored_center_ratio(label) for label in labels]
    assert ratios[0] == pytest.approx(4.0, rel=0.08)
    assert ratios[1] == pytest.approx(1.0, rel=0.08)
    assert ratios[2] == pytest.approx(0.25, rel=0.08)


def test_main_triple_previews_use_equal_letterbox_canvases_after_resize(
    qapp, tmp_path, monkeypatch
):
    settings_path = tmp_path / "settings.ini"
    monkeypatch.setattr(
        main_window_module,
        "QSettings",
        lambda *_args: QSettings(str(settings_path), QSettings.Format.IniFormat),
    )
    window = main_window_module.MainWindow()
    try:
        window.resize(1500, 900)
        window.show()
        item = _triple_item()
        window.state.set_items(
            {1: [item]},
            mode="triple",
            workbook_paths={"ref": "ref.xlsx", "a": "a.xlsx", "b": "b.xlsx"},
        )
        window._populate_sheet_combo()
        window._render_current()
        qapp.processEvents()
        QTest.qWait(70)
        _assert_equal_letterbox_previews(window)

        window.resize(1100, 700)
        qapp.processEvents()
        QTest.qWait(70)
        _assert_equal_letterbox_previews(window)
    finally:
        window.close()
        window.deleteLater()


def test_list_window_places_table_above_equal_triple_previews(qapp):
    item = _triple_item()
    window = ImageListWindow(
        [item],
        "triple",
        {"ref": "ref.xlsx", "a": "a.xlsx", "b": "b.xlsx"},
    )
    try:
        window.show()
        qapp.processEvents()
        QTest.qWait(70)

        assert window.content_splitter.orientation() == Qt.Orientation.Vertical
        assert window.content_splitter.count() == 2
        assert window.content_splitter.widget(0) is window.table
        assert window.content_splitter.widget(1) is window.preview_widget
        assert all(size > 0 for size in window.content_splitter.sizes())
        table_y = window.table.mapTo(window, QPoint(0, 0)).y()
        preview_y = window.preview_widget.mapTo(window, QPoint(0, 0)).y()
        assert table_y < preview_y

        widths = [
            pane.width()
            for pane in (window.ref_container, window.a_container, window.b_container)
        ]
        assert max(widths) - min(widths) <= 1
        _assert_equal_letterbox_previews(window)
    finally:
        window.close()
