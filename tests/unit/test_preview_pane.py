"""미리보기 헤더는 글자를 가운데에, 바로가기 버튼을 오른쪽 끝에 둔다."""

from PyQt6.QtCore import Qt

from ui.dialogs import create_preview_pane


def _meta_row(container):
    return container.layout().itemAt(1).layout()


def _assert_jump_button_at_right_edge(container, metadata, button, qapp) -> None:
    qapp.processEvents()
    meta_row = _meta_row(container)
    assert meta_row.count() == 2
    assert meta_row.itemAt(0).widget() is metadata
    assert meta_row.itemAt(1).widget() is button
    assert meta_row.stretch(0) == 1
    assert meta_row.stretch(1) == 0
    assert int(metadata.alignment()) & int(Qt.AlignmentFlag.AlignHCenter)

    margins = container.layout().contentsMargins()
    button_right = button.mapTo(container, button.rect().topRight()).x()
    assert button_right >= container.width() - margins.right() - 2

    meta_left = metadata.mapTo(container, metadata.rect().topLeft()).x()
    assert meta_left <= margins.left() + 4
    assert button.mapTo(container, button.rect().topLeft()).x() > meta_left
    assert metadata.width() >= button.width()


def test_create_preview_pane_puts_jump_button_on_the_right(qapp):
    container, _title, metadata, _image, button = create_preview_pane("Excel Ref")
    metadata.setText("SC CLKx 64 case · D137 · 시트 내 137/156")
    container.resize(420, 280)
    container.show()
    qapp.processEvents()
    try:
        _assert_jump_button_at_right_edge(container, metadata, button, qapp)
    finally:
        container.close()
        container.deleteLater()
