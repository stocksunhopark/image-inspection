"""단일 이미지 확대 창의 창맞춤과 확대/축소 동작."""

from PIL import Image

from ui.dialogs import ImageViewerDialog


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
