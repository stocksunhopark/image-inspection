"""단일 이미지 확대 창의 창맞춤과 확대/축소 동작."""

import pytest
from PIL import Image
from PyQt6.QtCore import QEvent, QPoint, QPointF, Qt
from PyQt6.QtGui import QKeyEvent, QMouseEvent, QWheelEvent
from PyQt6.QtWidgets import QApplication

from models import MISSING_SHEET, ExtractedImage, InspectionItem
from ui.dialogs import (
    SHOW_COMPARISON_RESULT,
    ImageComparisonDialog,
    ImageViewerDialog,
)


def _send_wheel(viewport, *, delta: int, modifiers) -> None:
    position = QPoint(viewport.width() // 3, viewport.height() // 3)
    event = QWheelEvent(
        QPointF(position),
        QPointF(viewport.mapToGlobal(position)),
        QPoint(0, 0),
        QPoint(0, delta),
        Qt.MouseButton.NoButton,
        modifiers,
        Qt.ScrollPhase.ScrollUpdate,
        False,
    )
    QApplication.sendEvent(viewport, event)


def _send_mouse(
    viewport,
    event_type,
    position: QPoint,
    *,
    button,
    buttons,
    modifiers,
) -> None:
    event = QMouseEvent(
        event_type,
        QPointF(position),
        QPointF(viewport.mapToGlobal(position)),
        button,
        buttons,
        modifiers,
    )
    QApplication.sendEvent(viewport, event)


def _extracted_image(
    color: tuple[int, int, int],
    *,
    role: str,
    cell: str = "A1",
    size: tuple[int, int] = (320, 180),
) -> ExtractedImage:
    return ExtractedImage(
        sheet_index=1,
        sheet_name="MAIN",
        cell_address=cell,
        merged_range="",
        anchor_row=0,
        anchor_col=0,
        image=Image.new("RGB", size, color),
        source_role=role,
    )


def _comparison_item(
    mode: str,
    *,
    size: tuple[int, int] = (320, 180),
) -> InspectionItem:
    return InspectionItem(
        1,
        1,
        _extracted_image(
            (200, 20, 20), role="ref", cell="A1", size=size
        ),
        _extracted_image((20, 200, 20), role="a", cell="B1", size=size),
        _extracted_image((20, 20, 200), role="b", cell="C1", size=size)
        if mode in {"triple", "quadra"}
        else None,
        _extracted_image((200, 200, 20), role="c", cell="D1", size=size)
        if mode == "quadra"
        else None,
    )


def test_image_viewer_fits_very_wide_image_and_allows_zoom(qapp):
    dialog = ImageViewerDialog(
        Image.new("RGB", (12000, 120), (20, 40, 80)),
        "확대 보기",
    )
    try:
        dialog.show()
        qapp.processEvents()
        qapp.processEvents()

        assert dialog._fit_mode is True
        assert 0 < dialog._scale < 0.1
        fitted_scale = dialog._scale

        dialog._zoom_by(1.25)
        assert dialog._fit_mode is False
        assert dialog._scale > fitted_scale

        dialog._show_actual_size()
        assert dialog._scale == 1.0
    finally:
        dialog.close()
        dialog.deleteLater()


def test_image_viewer_comparison_button_returns_requested_result(qapp):
    dialog = ImageViewerDialog(
        Image.new("RGB", (320, 180), (20, 40, 80)),
        "확대",
        comparison_available=True,
    )
    try:
        dialog.show()
        qapp.processEvents()
        assert dialog.comparison_button.text() == "다른파형 함께보기"
        assert dialog.comparison_button.isVisible()

        dialog.comparison_button.click()

        assert dialog.result() == SHOW_COMPARISON_RESULT
        assert not dialog.isVisible()
    finally:
        dialog.close()
        dialog.deleteLater()


@pytest.mark.parametrize(
    ("mode", "expected_roles", "expected_positions"),
    [
        ("double", ("ref", "a"), {"ref": (0, 0), "a": (0, 1)}),
        (
            "triple",
            ("ref", "a", "b"),
            {"ref": (0, 0), "a": (0, 1), "b": (0, 2)},
        ),
        (
            "quadra",
            ("ref", "a", "b", "c"),
            {"ref": (0, 0), "a": (0, 1), "b": (1, 0), "c": (1, 1)},
        ),
    ],
)
def test_comparison_dialog_builds_mode_specific_layout(
    qapp, mode, expected_roles, expected_positions
):
    dialog = ImageComparisonDialog(
        _comparison_item(mode),
        mode,
        {role: f"{role}.xlsx" for role in expected_roles},
    )
    try:
        dialog.showMaximized()
        qapp.processEvents()
        qapp.processEvents()

        assert tuple(dialog.panes) == expected_roles
        assert dialog.pane_positions == expected_positions
        assert all(pane.has_image for pane in dialog.panes.values())
        assert dialog.isMaximized()
        assert dialog.windowFlags() & Qt.WindowType.WindowMaximizeButtonHint
    finally:
        dialog.close()
        dialog.deleteLater()


@pytest.mark.parametrize(
    ("mode", "role_order", "expected_positions"),
    [
        ("double", ("a", "ref"), {"a": (0, 0), "ref": (0, 1)}),
        (
            "triple",
            ("b", "a", "ref"),
            {"b": (0, 0), "a": (0, 1), "ref": (0, 2)},
        ),
        (
            "quadra",
            ("c", "a", "b", "ref"),
            {"c": (0, 0), "a": (0, 1), "b": (1, 0), "ref": (1, 1)},
        ),
    ],
)
def test_comparison_dialog_uses_reordered_list_layout(
    qapp, mode, role_order, expected_positions
):
    dialog = ImageComparisonDialog(
        _comparison_item(mode),
        mode,
        {role: f"{role}.xlsx" for role in role_order},
        role_order=role_order,
    )
    try:
        dialog.show()
        qapp.processEvents()

        assert dialog.role_order == role_order
        assert tuple(dialog.panes) == role_order
        assert dialog.pane_positions == expected_positions
        for role in role_order:
            assert role in dialog.panes[role].metadata.text()
    finally:
        dialog.close()
        dialog.deleteLater()


def test_comparison_dialog_keeps_missing_role_as_placeholder(qapp):
    missing_a = ExtractedImage(
        sheet_index=1,
        sheet_name="MAIN",
        cell_address="A1",
        merged_range="",
        anchor_row=0,
        anchor_col=0,
        is_null=True,
        source_role="a",
        missing_reason=MISSING_SHEET,
    )
    item = InspectionItem(
        1,
        1,
        _extracted_image((200, 20, 20), role="ref"),
        missing_a,
    )
    dialog = ImageComparisonDialog(item, "double", {"ref": "ref.xlsx"})
    try:
        dialog.show()
        qapp.processEvents()

        assert tuple(dialog.panes) == ("ref", "a")
        assert dialog.panes["ref"].has_image
        assert not dialog.panes["a"].has_image
        assert (
            dialog.panes["a"]._image_label.text()
            == "현재 Excel에는 이 시트가 없습니다."
        )
    finally:
        dialog.close()
        dialog.deleteLater()


@pytest.mark.parametrize("mode", ["double", "triple", "quadra"])
def test_comparison_ctrl_wheel_zooms_every_image(qapp, mode):
    dialog = ImageComparisonDialog(
        _comparison_item(mode, size=(1200, 900)),
        mode,
        {},
    )
    try:
        dialog.resize(1500, 900)
        dialog.show()
        qapp.processEvents()
        qapp.processEvents()
        panes = list(dialog.panes.values())
        initial_scales = {pane.role: pane._scale for pane in panes}
        source = panes[-1]

        _send_wheel(
            source._scroll.viewport(),
            delta=120,
            modifiers=Qt.KeyboardModifier.ControlModifier,
        )

        for pane in panes:
            assert pane._scale == pytest.approx(
                initial_scales[pane.role] * 1.25
            )
        x_centers = [pane.normalized_axis_center("x") for pane in panes]
        y_centers = [pane.normalized_axis_center("y") for pane in panes]
        assert max(x_centers) - min(x_centers) < 0.005
        assert max(y_centers) - min(y_centers) < 0.005
    finally:
        dialog.close()
        dialog.deleteLater()


def test_reordered_comparison_keeps_synchronized_zoom(qapp):
    order = ("c", "a", "b", "ref")
    dialog = ImageComparisonDialog(
        _comparison_item("quadra", size=(1200, 900)),
        "quadra",
        {},
        role_order=order,
    )
    try:
        dialog.resize(1500, 900)
        dialog.show()
        qapp.processEvents()
        qapp.processEvents()
        initial = {role: dialog.panes[role]._scale for role in order}

        _send_wheel(
            dialog.panes["c"]._scroll.viewport(),
            delta=120,
            modifiers=Qt.KeyboardModifier.ControlModifier,
        )

        assert tuple(dialog.panes) == order
        for role in order:
            assert dialog.panes[role]._scale == pytest.approx(initial[role] * 1.25)
    finally:
        dialog.close()
        dialog.deleteLater()


def test_comparison_zoom_sync_uses_normalized_position_for_different_sizes(qapp):
    item = InspectionItem(
        1,
        1,
        _extracted_image(
            (200, 20, 20), role="ref", size=(1600, 1200)
        ),
        _extracted_image((20, 200, 20), role="a", size=(800, 600)),
    )
    dialog = ImageComparisonDialog(item, "double", {})
    try:
        dialog.resize(1400, 800)
        dialog.show()
        qapp.processEvents()
        qapp.processEvents()
        source = dialog.panes["a"]

        _send_wheel(
            source._scroll.viewport(),
            delta=120,
            modifiers=Qt.KeyboardModifier.ControlModifier,
        )

        ref_center = dialog.panes["ref"].normalized_view_center()
        a_center = dialog.panes["a"].normalized_view_center()
        assert ref_center[0] == pytest.approx(a_center[0], abs=0.005)
        assert ref_center[1] == pytest.approx(a_center[1], abs=0.005)
    finally:
        dialog.close()
        dialog.deleteLater()


def test_comparison_scrollbars_and_plain_wheel_move_every_image(qapp):
    dialog = ImageComparisonDialog(
        _comparison_item("quadra", size=(1200, 900)),
        "quadra",
        {},
    )
    try:
        dialog.resize(1400, 850)
        dialog.show()
        qapp.processEvents()
        dialog._show_all_actual_size()
        qapp.processEvents()
        source = dialog.panes["b"]
        horizontal = source._scroll.horizontalScrollBar()
        vertical = source._scroll.verticalScrollBar()
        assert horizontal.maximum() > 0
        assert vertical.maximum() > 0

        horizontal.setValue(horizontal.maximum() * 3 // 4)
        vertical.setValue(vertical.maximum() * 2 // 3)
        for pane in dialog.panes.values():
            assert pane.normalized_axis_center("x") == pytest.approx(
                source.normalized_axis_center("x"), abs=0.005
            )
            assert pane.normalized_axis_center("y") == pytest.approx(
                source.normalized_axis_center("y"), abs=0.005
            )

        previous_y = source.normalized_axis_center("y")
        _send_wheel(
            source._scroll.viewport(),
            delta=-120,
            modifiers=Qt.KeyboardModifier.NoModifier,
        )
        assert source.normalized_axis_center("y") > previous_y
        for pane in dialog.panes.values():
            assert pane.normalized_axis_center("y") == pytest.approx(
                source.normalized_axis_center("y"), abs=0.005
            )
    finally:
        dialog.close()
        dialog.deleteLater()


def test_comparison_ctrl_drag_moves_every_image(qapp):
    dialog = ImageComparisonDialog(
        _comparison_item("triple", size=(1200, 900)),
        "triple",
        {},
    )
    try:
        dialog.resize(1400, 800)
        dialog.show()
        qapp.processEvents()
        dialog._show_all_actual_size()
        qapp.processEvents()
        source = dialog.panes["a"]
        horizontal = source._scroll.horizontalScrollBar()
        vertical = source._scroll.verticalScrollBar()
        horizontal.setValue(horizontal.maximum() // 2)
        vertical.setValue(vertical.maximum() // 2)
        previous_center = source.normalized_view_center()
        viewport = source._scroll.viewport()
        start = QPoint(viewport.width() // 2, viewport.height() // 2)
        moved = start + QPoint(70, 55)

        _send_mouse(
            viewport,
            QEvent.Type.MouseButtonPress,
            start,
            button=Qt.MouseButton.LeftButton,
            buttons=Qt.MouseButton.LeftButton,
            modifiers=Qt.KeyboardModifier.ControlModifier,
        )
        _send_mouse(
            viewport,
            QEvent.Type.MouseMove,
            moved,
            button=Qt.MouseButton.NoButton,
            buttons=Qt.MouseButton.LeftButton,
            modifiers=Qt.KeyboardModifier.ControlModifier,
        )
        _send_mouse(
            viewport,
            QEvent.Type.MouseButtonRelease,
            moved,
            button=Qt.MouseButton.LeftButton,
            buttons=Qt.MouseButton.NoButton,
            modifiers=Qt.KeyboardModifier.ControlModifier,
        )

        assert source.normalized_view_center() != previous_center
        for pane in dialog.panes.values():
            assert pane.normalized_axis_center("x") == pytest.approx(
                source.normalized_axis_center("x"), abs=0.005
            )
            assert pane.normalized_axis_center("y") == pytest.approx(
                source.normalized_axis_center("y"), abs=0.005
            )
    finally:
        dialog.close()
        dialog.deleteLater()


def test_comparison_zoom_limit_stops_all_images_together(qapp):
    dialog = ImageComparisonDialog(
        _comparison_item("double", size=(40, 30)),
        "double",
        {},
    )
    try:
        dialog.show()
        qapp.processEvents()
        for pane in dialog.panes.values():
            pane._apply_scale(7.9)

        dialog._zoom_all(1.25)

        assert all(
            pane._scale == pytest.approx(8.0)
            for pane in dialog.panes.values()
        )
    finally:
        dialog.close()
        dialog.deleteLater()


def test_ctrl_mouse_wheel_zooms_in_and_out(qapp):
    dialog = ImageViewerDialog(
        Image.new("RGB", (1600, 1200), (20, 40, 80)),
        "확대",
    )
    try:
        dialog.show()
        qapp.processEvents()
        qapp.processEvents()
        viewport = dialog._scroll.viewport()
        initial_scale = dialog._scale

        _send_wheel(
            viewport,
            delta=120,
            modifiers=Qt.KeyboardModifier.ControlModifier,
        )
        assert dialog._fit_mode is False
        assert dialog._scale == pytest.approx(initial_scale * 1.25)

        _send_wheel(
            viewport,
            delta=-120,
            modifiers=Qt.KeyboardModifier.ControlModifier,
        )
        assert dialog._scale == pytest.approx(initial_scale)
    finally:
        dialog.close()
        dialog.deleteLater()


def test_plain_mouse_wheel_scrolls_without_zooming(qapp):
    dialog = ImageViewerDialog(
        Image.new("RGB", (1600, 2400), (20, 40, 80)),
        "확대",
    )
    try:
        dialog.show()
        qapp.processEvents()
        dialog._show_actual_size()
        qapp.processEvents()
        viewport = dialog._scroll.viewport()
        initial_scale = dialog._scale
        scrollbar = dialog._scroll.verticalScrollBar()
        assert scrollbar.maximum() > 0
        assert scrollbar.value() == 0

        _send_wheel(
            viewport,
            delta=-120,
            modifiers=Qt.KeyboardModifier.NoModifier,
        )
        assert dialog._scale == initial_scale
        assert scrollbar.value() > 0
    finally:
        dialog.close()
        dialog.deleteLater()


def test_ctrl_wheel_keeps_pointed_image_position_stable(qapp):
    dialog = ImageViewerDialog(
        Image.new("RGB", (2400, 1800), (20, 40, 80)),
        "확대",
    )
    try:
        dialog.show()
        qapp.processEvents()
        dialog._show_actual_size()
        qapp.processEvents()
        horizontal = dialog._scroll.horizontalScrollBar()
        vertical = dialog._scroll.verticalScrollBar()
        horizontal.setValue(horizontal.maximum() // 2)
        vertical.setValue(vertical.maximum() // 2)
        viewport = dialog._scroll.viewport()
        position = QPoint(viewport.width() // 3, viewport.height() // 3)
        old_width = dialog._image_label.width()
        old_height = dialog._image_label.height()
        old_image_x = position.x() - dialog._image_label.x()
        old_image_y = position.y() - dialog._image_label.y()

        event = QWheelEvent(
            QPointF(position),
            QPointF(viewport.mapToGlobal(position)),
            QPoint(0, 0),
            QPoint(0, 120),
            Qt.MouseButton.NoButton,
            Qt.KeyboardModifier.ControlModifier,
            Qt.ScrollPhase.ScrollUpdate,
            False,
        )
        QApplication.sendEvent(viewport, event)
        qapp.processEvents()

        new_image_x = position.x() - dialog._image_label.x()
        new_image_y = position.y() - dialog._image_label.y()
        assert new_image_x / dialog._image_label.width() == pytest.approx(
            old_image_x / old_width,
            abs=0.002,
        )
        assert new_image_y / dialog._image_label.height() == pytest.approx(
            old_image_y / old_height,
            abs=0.002,
        )
    finally:
        dialog.close()
        dialog.deleteLater()


def test_ctrl_changes_cursor_and_drag_pans_zoomed_image(qapp):
    dialog = ImageViewerDialog(
        Image.new("RGB", (2400, 1800), (20, 40, 80)),
        "확대",
    )
    try:
        dialog.show()
        qapp.processEvents()
        dialog._show_actual_size()
        qapp.processEvents()
        viewport = dialog._scroll.viewport()
        horizontal = dialog._scroll.horizontalScrollBar()
        vertical = dialog._scroll.verticalScrollBar()
        horizontal.setValue(horizontal.maximum() // 2)
        vertical.setValue(vertical.maximum() // 2)
        start_horizontal = horizontal.value()
        start_vertical = vertical.value()

        QApplication.sendEvent(
            dialog,
            QKeyEvent(
                QEvent.Type.KeyPress,
                Qt.Key.Key_Control,
                Qt.KeyboardModifier.ControlModifier,
            ),
        )
        assert viewport.cursor().shape() == Qt.CursorShape.OpenHandCursor

        start = QPoint(viewport.width() // 2, viewport.height() // 2)
        moved = start + QPoint(60, 45)
        _send_mouse(
            viewport,
            QEvent.Type.MouseButtonPress,
            start,
            button=Qt.MouseButton.LeftButton,
            buttons=Qt.MouseButton.LeftButton,
            modifiers=Qt.KeyboardModifier.ControlModifier,
        )
        assert viewport.cursor().shape() == Qt.CursorShape.ClosedHandCursor
        _send_mouse(
            viewport,
            QEvent.Type.MouseMove,
            moved,
            button=Qt.MouseButton.NoButton,
            buttons=Qt.MouseButton.LeftButton,
            modifiers=Qt.KeyboardModifier.ControlModifier,
        )
        assert horizontal.value() == start_horizontal - 60
        assert vertical.value() == start_vertical - 45

        _send_mouse(
            viewport,
            QEvent.Type.MouseButtonRelease,
            moved,
            button=Qt.MouseButton.LeftButton,
            buttons=Qt.MouseButton.NoButton,
            modifiers=Qt.KeyboardModifier.ControlModifier,
        )
        assert viewport.cursor().shape() == Qt.CursorShape.OpenHandCursor

        QApplication.sendEvent(
            dialog,
            QKeyEvent(
                QEvent.Type.KeyRelease,
                Qt.Key.Key_Control,
                Qt.KeyboardModifier.NoModifier,
            ),
        )
        assert viewport.cursor().shape() == Qt.CursorShape.ArrowCursor
    finally:
        dialog.close()
        dialog.deleteLater()


def test_drag_without_ctrl_does_not_pan_image(qapp):
    dialog = ImageViewerDialog(
        Image.new("RGB", (2400, 1800), (20, 40, 80)),
        "확대",
    )
    try:
        dialog.show()
        qapp.processEvents()
        dialog._show_actual_size()
        qapp.processEvents()
        viewport = dialog._scroll.viewport()
        horizontal = dialog._scroll.horizontalScrollBar()
        vertical = dialog._scroll.verticalScrollBar()
        horizontal.setValue(horizontal.maximum() // 2)
        vertical.setValue(vertical.maximum() // 2)
        start_values = (horizontal.value(), vertical.value())
        start = QPoint(viewport.width() // 2, viewport.height() // 2)
        moved = start + QPoint(80, 50)

        _send_mouse(
            viewport,
            QEvent.Type.MouseButtonPress,
            start,
            button=Qt.MouseButton.LeftButton,
            buttons=Qt.MouseButton.LeftButton,
            modifiers=Qt.KeyboardModifier.NoModifier,
        )
        _send_mouse(
            viewport,
            QEvent.Type.MouseMove,
            moved,
            button=Qt.MouseButton.NoButton,
            buttons=Qt.MouseButton.LeftButton,
            modifiers=Qt.KeyboardModifier.NoModifier,
        )
        _send_mouse(
            viewport,
            QEvent.Type.MouseButtonRelease,
            moved,
            button=Qt.MouseButton.LeftButton,
            buttons=Qt.MouseButton.NoButton,
            modifiers=Qt.KeyboardModifier.NoModifier,
        )

        assert (horizontal.value(), vertical.value()) == start_values
        assert viewport.cursor().shape() == Qt.CursorShape.ArrowCursor
    finally:
        dialog.close()
        dialog.deleteLater()


def test_jump_button_stays_disabled_without_callback(qapp):
    dialog = ImageViewerDialog(Image.new("RGB", (40, 30), (10, 20, 30)), "확대")
    try:
        dialog.show()
        qapp.processEvents()
        assert dialog.jump_button.text() == "엑셀 파형 바로가기"
        assert dialog.jump_button.isEnabled() is False
        dialog.jump_button.click()
    finally:
        dialog.close()
        dialog.deleteLater()


def test_jump_button_calls_callback_when_enabled(qapp):
    called = []
    dialog = ImageViewerDialog(
        Image.new("RGB", (40, 30), (10, 20, 30)),
        "확대",
        jump_enabled=True,
        jump_callback=lambda: called.append("jump"),
    )
    try:
        dialog.show()
        qapp.processEvents()
        assert dialog.jump_button.isEnabled() is True
        dialog.jump_button.click()
        assert called == ["jump"]
    finally:
        dialog.close()
        dialog.deleteLater()
