"""Qt 이미지 표시용 작은 변환 함수."""

from PIL import Image
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QImage, QPixmap


def pil_to_pixmap(
    image: Image.Image,
    max_w: int = 2560,
    max_h: int = 1440,
    *,
    upscale: bool = False,
) -> QPixmap:
    source = image.convert("RGB")
    width, height = source.size
    if width > max_w or height > max_h:
        source = source.copy()
        source.thumbnail((max_w, max_h), Image.Resampling.LANCZOS)
        width, height = source.size
    raw = source.tobytes("raw", "RGB")
    qimage = QImage(
        raw,
        width,
        height,
        width * 3,
        QImage.Format.Format_RGB888,
    ).copy()
    pixmap = QPixmap.fromImage(qimage)
    if not upscale or width >= max_w or height >= max_h:
        return pixmap
    return pixmap.scaled(
        max_w,
        max_h,
        Qt.AspectRatioMode.KeepAspectRatio,
        Qt.TransformationMode.SmoothTransformation,
    )
