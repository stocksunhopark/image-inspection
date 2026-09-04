from typing import Optional

from PyQt6.QtCore import QSize, Qt, QTimer, pyqtSignal
from PyQt6.QtGui import QColor, QPainter, QPixmap
from PyQt6.QtWidgets import QLabel, QLineEdit, QSizePolicy


class DropLineEdit(QLineEdit):
    """엑셀 파일을 끌어다 놓으면 경로가 입력되는 입력창."""
    @staticmethod
    def _extract_xlsx(event) -> Optional[str]:
        md = event.mimeData()
        if not md.hasUrls():
            return None
        for url in md.urls():
            path = url.toLocalFile()
            if path.lower().endswith((".xlsx", ".xlsm")):
                return path
        return None
    def dragEnterEvent(self, event):
        if self._extract_xlsx(event):
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)
    def dragMoveEvent(self, event):
        if self._extract_xlsx(event):
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)
    def dropEvent(self, event):
        path = self._extract_xlsx(event)
        if path:
            self.setText(path)
            event.acceptProposedAction()
        else:
            super().dropEvent(event)


class ClickableLabel(QLabel):
    """클릭하면 신호를 보내는 미리보기 라벨.

    원본 픽스맵 크기가 창 최소 크기를 키우지 않도록, 칸 크기에 맞춰 축소한다.
    """
    clicked = pyqtSignal()

    def __init__(self, text: str = ""):
        super().__init__(text)
        self._source = QPixmap()
        self._fit_timer = QTimer(self)
        self._fit_timer.setSingleShot(True)
        self._fit_timer.setInterval(40)
        self._fit_timer.timeout.connect(self._refit_smooth)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setScaledContents(False)
        self.setMinimumSize(80, 80)
        self.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Ignored)

    def minimumSizeHint(self) -> QSize:
        return QSize(80, 80)

    def sizeHint(self) -> QSize:
        return QSize(240, 160)

    def setPixmap(self, pixmap: QPixmap) -> None:
        self._source = QPixmap(pixmap) if pixmap is not None and not pixmap.isNull() else QPixmap()
        if self._source.isNull():
            self._fit_timer.stop()
            super().setPixmap(QPixmap())
            return
        self._refit(fast=True)
        self._fit_timer.start()

    def setText(self, text: str) -> None:
        if text:
            self._source = QPixmap()
            self._fit_timer.stop()
        super().setText(text)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        if not self._source.isNull():
            self._refit(fast=True)
            self._fit_timer.start()

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)

    def _refit_smooth(self) -> None:
        self._refit(fast=False)

    def _refit(self, fast: bool) -> None:
        if self._source.isNull():
            return
        dpr = max(1.0, float(self.devicePixelRatioF()))
        w = max(1, int((self.width() - 4) * dpr))
        h = max(1, int((self.height() - 4) * dpr))
        mode = (
            Qt.TransformationMode.FastTransformation
            if fast
            else Qt.TransformationMode.SmoothTransformation
        )
        scaled = self._source.scaled(w, h, Qt.AspectRatioMode.KeepAspectRatio, mode)
        # 서로 다른 해상도/종횡비라도 모든 preview는 같은 크기의 캔버스를
        # 사용한다. 원본 비율은 유지하고 남는 영역만 letterbox로 채운다.
        canvas = QPixmap(w, h)
        canvas.fill(QColor("#020617"))
        painter = QPainter(canvas)
        painter.drawPixmap(
            max(0, (w - scaled.width()) // 2),
            max(0, (h - scaled.height()) // 2),
            scaled,
        )
        painter.end()
        canvas.setDevicePixelRatio(dpr)
        super().setPixmap(canvas)
