"""육안 검사 화면에서 사용하는 확대 창과 이미지 리스트 창."""

import os
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Callable, Dict, List, Optional, Sequence, Tuple

from PIL import Image
from PyQt6.QtCore import (
    QByteArray,
    QEvent,
    QMimeData,
    QPoint,
    QPointF,
    QSettings,
    QSignalBlocker,
    Qt,
    QTimer,
    pyqtSignal,
)
from PyQt6.QtGui import QDrag, QKeyEvent, QKeySequence, QPixmap, QShortcut
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QGridLayout,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from excel_jump import jump_target_for_list_cell, jump_target_for_side, jump_to_excel_cell
from excel_manager import load_extracted_pil
from models import (
    MODE_ROLES,
    ROLE_ORDER,
    SUPPORTED_MODES,
    ExtractedImage,
    InspectionItem,
    mode_label,
    normalize_preview_order,
    swapped_preview_order,
)
from review_annotations import (
    ROLE_DISPLAY_NAMES,
    ReviewAnnotationStore,
    export_annotations_xlsx,
)
from ui.helpers import pil_to_pixmap
from ui.widgets import ClickableLabel


SHOW_COMPARISON_RESULT = 2
_PREVIEW_ROLE_MIME = "application/x-image-inspection-preview-role"
_WAVEFORM_VERSION_PATTERNS = (
    re.compile(
        r"(?<![A-Z0-9])M\d{1,2}E\d{1,2}(?![A-Z0-9])",
        re.IGNORECASE,
    ),
    re.compile(
        r"(?<![A-Z0-9])MVT\d{1,2}-\d{1,2}(?![A-Z0-9])",
        re.IGNORECASE,
    ),
)
_WAVEFORM_TEMPERATURE_PATTERN = re.compile(
    r"(?<![A-Z0-9])(ROOM|HOT|LOW)(?![A-Z0-9])",
    re.IGNORECASE,
)
class _CompactAttributeLineEdit(QLineEdit):
    """대표적인 긴 속성 문자열이 온전히 보이는 자유 입력칸."""

    _REFERENCE_TEXT = "MVT99-99"

    def _compact_size_hint(self):
        hint = super().sizeHint()
        text_width = self.fontMetrics().horizontalAdvance(self._REFERENCE_TEXT)
        hint.setWidth(max(78, text_width + 18))
        return hint

    def sizeHint(self):
        return self._compact_size_hint()

    def minimumSizeHint(self):
        return self._compact_size_hint()


def parse_waveform_attributes_from_path(path: str) -> Tuple[str, str]:
    """파일명에서 알려진 버전 표기와 온도를 찾아 기본값과 함께 반환한다."""
    filename = os.path.basename(str(path or ""))
    version = "M0E0"
    for pattern in _WAVEFORM_VERSION_PATTERNS:
        version_match = pattern.search(filename)
        if version_match is not None:
            version = version_match.group(0).upper()
            break
    temperature_match = _WAVEFORM_TEMPERATURE_PATTERN.search(filename)
    temperature = (
        temperature_match.group(1).upper()
        if temperature_match is not None
        else "ROOM"
    )
    return version, temperature


def waveform_attribute_text(version: str, temperature: str) -> str:
    """확대 화면에 표시할 읽기 전용 버전·온도 문구를 만든다."""
    return f"버전: {str(version).strip() or '-'} · 온도: {str(temperature).strip() or '-'}"


@dataclass(frozen=True)
class PreviewAttributeControls:
    """한 Excel 미리보기 패널에 속한 버전·온도 입력 위젯."""

    version_edit: QLineEdit
    temperature_edit: QLineEdit

    @property
    def version_text(self) -> str:
        return self.version_edit.text()

    @property
    def temperature(self) -> str:
        return self.temperature_edit.text()

    def set_values(self, version: str, temperature: str) -> None:
        self.version_edit.setText(str(version))
        self.temperature_edit.setText(str(temperature))


class _DraggablePreviewTitle(QLabel):
    """목록 미리보기의 의미 역할을 유지한 채 제목에서 드래그를 시작한다."""

    def __init__(self, text: str = "", parent=None) -> None:
        super().__init__(text, parent)
        self._drag_role = ""
        self._drag_start: Optional[QPoint] = None

    @property
    def drag_role(self) -> str:
        return self._drag_role

    def enable_role_drag(self, role: str) -> None:
        self._drag_role = str(role)
        self.setCursor(Qt.CursorShape.OpenHandCursor)

    def mousePressEvent(self, event) -> None:
        if (
            self._drag_role
            and event.button() == Qt.MouseButton.LeftButton
        ):
            self._drag_start = event.position().toPoint()
        else:
            self._drag_start = None
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event) -> None:
        if (
            not self._drag_role
            or self._drag_start is None
            or not (event.buttons() & Qt.MouseButton.LeftButton)
        ):
            super().mouseMoveEvent(event)
            return
        distance = (event.position().toPoint() - self._drag_start).manhattanLength()
        if distance < QApplication.startDragDistance():
            super().mouseMoveEvent(event)
            return

        mime = QMimeData()
        mime.setData(_PREVIEW_ROLE_MIME, QByteArray(self._drag_role.encode("ascii")))
        drag = QDrag(self)
        drag.setMimeData(mime)
        drag.setPixmap(self.grab())
        drag.setHotSpot(event.position().toPoint())
        self._drag_start = None
        self.setCursor(Qt.CursorShape.ClosedHandCursor)
        try:
            drag.exec(Qt.DropAction.MoveAction, Qt.DropAction.MoveAction)
        finally:
            self.setCursor(Qt.CursorShape.OpenHandCursor)
        event.accept()

    def mouseReleaseEvent(self, event) -> None:
        self._drag_start = None
        super().mouseReleaseEvent(event)


class _PreviewDropContainer(QWidget):
    """같은 목록 창의 다른 역할 제목을 놓으면 두 표시 위치의 교환을 요청한다."""

    preview_swap_requested = pyqtSignal(str, str)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        # QWidget을 상속한 사용자 정의 컨테이너도 QSS의 배경·테두리를 직접 그린다.
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self._drop_role = ""
        self.setProperty("dragTarget", False)

    @property
    def drop_role(self) -> str:
        return self._drop_role

    def enable_role_drop(self, role: str) -> None:
        self._drop_role = str(role)
        self.setAcceptDrops(True)

    def _source_role(self, event) -> Optional[str]:
        if not self._drop_role or not event.mimeData().hasFormat(_PREVIEW_ROLE_MIME):
            return None
        source = event.source()
        if (
            not isinstance(source, _DraggablePreviewTitle)
            or source.window() is not self.window()
        ):
            return None
        try:
            role = bytes(event.mimeData().data(_PREVIEW_ROLE_MIME)).decode("ascii")
        except (UnicodeDecodeError, ValueError):
            return None
        if (
            role != source.drag_role
            or role == self._drop_role
            or role not in ROLE_ORDER
        ):
            return None
        return role

    def _set_drag_target(self, active: bool) -> None:
        active = bool(active)
        if self.property("dragTarget") == active:
            return
        self.setProperty("dragTarget", active)
        self.style().unpolish(self)
        self.style().polish(self)
        self.update()

    def dragEnterEvent(self, event) -> None:
        if self._source_role(event) is not None:
            self._set_drag_target(True)
            event.setDropAction(Qt.DropAction.MoveAction)
            event.accept()
            return
        event.ignore()

    def dragMoveEvent(self, event) -> None:
        if self._source_role(event) is not None:
            event.setDropAction(Qt.DropAction.MoveAction)
            event.accept()
            return
        self._set_drag_target(False)
        event.ignore()

    def dragLeaveEvent(self, event) -> None:
        self._set_drag_target(False)
        event.accept()

    def dropEvent(self, event) -> None:
        source_role = self._source_role(event)
        self._set_drag_target(False)
        if source_role is None:
            event.ignore()
            return
        self.preview_swap_requested.emit(source_role, self._drop_role)
        event.setDropAction(Qt.DropAction.MoveAction)
        event.accept()


def swap_splitter_widgets(
    first_splitter: QSplitter,
    first_index: int,
    first_widget: QWidget,
    second_splitter: QSplitter,
    second_index: int,
    second_widget: QWidget,
) -> None:
    """같은/다른 splitter의 두 슬롯을 임시 자리로 안전하게 교환한다."""
    first_hidden = first_widget.isHidden()
    second_hidden = second_widget.isHidden()
    placeholder = QWidget()
    placeholder.setVisible(False)
    first_splitter.replaceWidget(first_index, placeholder)
    second_splitter.replaceWidget(second_index, first_widget)
    first_splitter.replaceWidget(first_index, second_widget)
    first_widget.setVisible(not first_hidden)
    second_widget.setVisible(not second_hidden)
    placeholder.deleteLater()


class _ZoomScrollArea(QScrollArea):
    """Ctrl+휠 확대와 Ctrl+좌클릭 드래그 이동을 제공한다."""

    def __init__(
        self,
        zoom_callback: Callable[[float, QPointF], None],
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._zoom_callback = zoom_callback
        self._control_pressed = False
        self._panning = False
        self._pan_start = QPoint()
        self._pan_scroll_start = QPoint()
        self.viewport().setMouseTracking(True)
        self.viewport().installEventFilter(self)
        application = QApplication.instance()
        if application is not None:
            application.installEventFilter(self)

    def _set_control_pressed(self, pressed: bool) -> None:
        self._control_pressed = bool(pressed)
        if not self._control_pressed and self._panning:
            self._panning = False
        self._update_pan_cursor()

    def _update_pan_cursor(self) -> None:
        if self._panning:
            self.viewport().setCursor(Qt.CursorShape.ClosedHandCursor)
        elif self._control_pressed:
            self.viewport().setCursor(Qt.CursorShape.OpenHandCursor)
        else:
            self.viewport().unsetCursor()

    def _begin_pan(self, position: QPoint) -> None:
        self._panning = True
        self._pan_start = QPoint(position)
        self._pan_scroll_start = QPoint(
            self.horizontalScrollBar().value(),
            self.verticalScrollBar().value(),
        )
        self.setFocus(Qt.FocusReason.MouseFocusReason)
        self._update_pan_cursor()

    def _move_pan(self, position: QPoint) -> None:
        delta = position - self._pan_start
        self.horizontalScrollBar().setValue(
            self._pan_scroll_start.x() - delta.x()
        )
        self.verticalScrollBar().setValue(
            self._pan_scroll_start.y() - delta.y()
        )

    def _end_pan(self) -> None:
        self._panning = False
        self._update_pan_cursor()

    def eventFilter(self, watched, event) -> bool:
        event_type = event.type()
        if (
            event_type == QEvent.Type.KeyPress
            and event.key() == Qt.Key.Key_Control
            and not event.isAutoRepeat()
        ):
            self._set_control_pressed(True)
        elif (
            event_type == QEvent.Type.KeyRelease
            and event.key() == Qt.Key.Key_Control
            and not event.isAutoRepeat()
        ):
            self._set_control_pressed(False)
        elif event_type == QEvent.Type.ApplicationDeactivate:
            self._set_control_pressed(False)

        if watched is self.viewport():
            if event_type == QEvent.Type.Enter:
                self._set_control_pressed(
                    bool(
                        QApplication.keyboardModifiers()
                        & Qt.KeyboardModifier.ControlModifier
                    )
                )
            elif event_type == QEvent.Type.MouseButtonPress:
                control_pressed = bool(
                    event.modifiers() & Qt.KeyboardModifier.ControlModifier
                )
                self._set_control_pressed(control_pressed)
                if (
                    event.button() == Qt.MouseButton.LeftButton
                    and control_pressed
                ):
                    self._begin_pan(event.position().toPoint())
                    event.accept()
                    return True
            elif event_type == QEvent.Type.MouseMove:
                control_pressed = bool(
                    event.modifiers() & Qt.KeyboardModifier.ControlModifier
                )
                self._set_control_pressed(control_pressed)
                if self._panning and (
                    event.buttons() & Qt.MouseButton.LeftButton
                ):
                    self._move_pan(event.position().toPoint())
                    event.accept()
                    return True
                if self._panning:
                    self._end_pan()
            elif event_type == QEvent.Type.MouseButtonRelease:
                if (
                    event.button() == Qt.MouseButton.LeftButton
                    and self._panning
                ):
                    self._end_pan()
                    self._set_control_pressed(
                        bool(
                            event.modifiers()
                            & Qt.KeyboardModifier.ControlModifier
                        )
                    )
                    event.accept()
                    return True
            elif event_type in {
                QEvent.Type.FocusOut,
                QEvent.Type.Hide,
            }:
                self._set_control_pressed(False)

        return super().eventFilter(watched, event)

    def wheelEvent(self, event) -> None:
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            delta = event.angleDelta().y()
            if delta == 0:
                delta = event.pixelDelta().y()
            if delta != 0:
                # 일반 마우스 휠 1칸(120)당 버튼과 같은 1.25배를 적용한다.
                factor = 1.25 ** (delta / 120.0)
                self._zoom_callback(factor, event.position())
                event.accept()
                return
        super().wheelEvent(event)


class ImageViewerDialog(QDialog):
    """원본 비율을 유지하며 창맞춤·확대·축소를 제공한다."""

    def __init__(
        self,
        pil_image: Image.Image,
        title: str,
        parent=None,
        *,
        jump_enabled: bool = False,
        jump_callback: Optional[Callable[[], None]] = None,
        comparison_available: bool = False,
        role: Optional[str] = None,
        workbook_path: str = "",
        location_text: str = "",
        version: str = "",
        temperature: str = "",
    ):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(1100, 800)
        self._jump_callback = jump_callback
        self._base_pixmap = pil_to_pixmap(
            pil_image,
            max_w=max(1, pil_image.width),
            max_h=max(1, pil_image.height),
            upscale=False,
        )
        self._scale = 1.0
        self._fit_mode = True

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        toolbar = QHBoxLayout()
        zoom_in = QPushButton("확대 +")
        zoom_out = QPushButton("축소 −")
        fit = QPushButton("창에 맞춤")
        actual = QPushButton("실제 크기")
        close_button = QPushButton("닫기")
        zoom_in.clicked.connect(lambda: self._zoom_by(1.25))
        zoom_out.clicked.connect(lambda: self._zoom_by(1.0 / 1.25))
        fit.clicked.connect(self._fit_to_window)
        actual.clicked.connect(self._show_actual_size)
        close_button.clicked.connect(self.accept)
        for button in (zoom_in, zoom_out, fit, actual, close_button):
            button.setObjectName("inputActionBtn")
            toolbar.addWidget(button)
        self.comparison_button = QPushButton("다른파형 함께보기")
        self.comparison_button.setObjectName("inputActionBtn")
        self.comparison_button.setToolTip(
            "현재 위치의 Ref/비교 파형을 한 화면에서 함께 봅니다."
        )
        self.comparison_button.setVisible(comparison_available)
        self.comparison_button.clicked.connect(
            lambda: self.done(SHOW_COMPARISON_RESULT)
        )
        toolbar.addWidget(self.comparison_button)
        toolbar.addStretch(1)
        self.jump_button = QPushButton("엑셀 파형 바로가기")
        self.jump_button.setObjectName("inputActionBtn")
        self.jump_button.setToolTip(
            "하이퍼링크 모드가 켜져 있으면 이 Excel의 해당 시트·셀로 이동합니다."
        )
        self.jump_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.jump_button.setEnabled(bool(jump_enabled) and jump_callback is not None)
        self.jump_button.setAutoDefault(False)
        self.jump_button.setDefault(False)
        self.jump_button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.jump_button.clicked.connect(self._on_jump_clicked)
        toolbar.addWidget(self.jump_button)
        layout.addLayout(toolbar)

        context_available = role is not None
        self.role_label = QLabel(
            ROLE_DISPLAY_NAMES.get(role, str(role or ""))
        )
        self.role_label.setObjectName("previewTitle")
        self.role_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.role_label.setVisible(context_available)
        layout.addWidget(self.role_label)

        filename = os.path.basename(workbook_path) if workbook_path else "-"
        self.metadata = QLabel(f"{filename} · {location_text or '-'}")
        self.metadata.setObjectName("imageMeta")
        self.metadata.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.metadata.setWordWrap(True)
        self.metadata.setVisible(context_available)
        layout.addWidget(self.metadata)

        self.attribute_metadata = QLabel(
            waveform_attribute_text(version, temperature)
        )
        self.attribute_metadata.setObjectName("waveformAttributeMeta")
        self.attribute_metadata.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.attribute_metadata.setVisible(context_available)
        layout.addWidget(self.attribute_metadata)

        self._scroll = _ZoomScrollArea(self._zoom_at)
        self._scroll.setWidgetResizable(False)
        self._scroll.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._scroll.setToolTip(
            "Ctrl + 마우스 휠: 확대·축소 / Ctrl + 좌클릭 드래그: 이미지 이동"
        )
        self._image_label = QLabel()
        self._image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._scroll.setWidget(self._image_label)
        layout.addWidget(self._scroll, stretch=1)

        QShortcut(QKeySequence("+"), self, activated=lambda: self._zoom_by(1.25))
        QShortcut(QKeySequence("-"), self, activated=lambda: self._zoom_by(0.8))
        QShortcut(QKeySequence("0"), self, activated=self._fit_to_window)
        self._apply_scale(1.0)
        QTimer.singleShot(0, self._fit_to_window)

    def _on_jump_clicked(self) -> None:
        if not self.jump_button.isEnabled():
            return
        if self._jump_callback is not None:
            self._jump_callback()

    def _apply_scale(self, scale: float) -> None:
        self._scale = max(0.01, min(8.0, float(scale)))
        width = max(1, int(self._base_pixmap.width() * self._scale))
        height = max(1, int(self._base_pixmap.height() * self._scale))
        scaled = self._base_pixmap.scaled(
            width,
            height,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        self._image_label.setPixmap(scaled)
        self._image_label.resize(scaled.size())

    def _zoom_by(self, factor: float) -> None:
        self._fit_mode = False
        self._apply_scale(self._scale * factor)

    def _zoom_at(self, factor: float, viewport_position: QPointF) -> None:
        """마우스 포인터 아래의 이미지 위치를 유지하며 확대·축소한다."""
        old_width = self._image_label.width()
        old_height = self._image_label.height()
        if old_width <= 0 or old_height <= 0:
            self._zoom_by(factor)
            return

        image_x = viewport_position.x() - self._image_label.x()
        image_y = viewport_position.y() - self._image_label.y()
        pointer_over_image = (
            0 <= image_x <= old_width and 0 <= image_y <= old_height
        )
        self._zoom_by(factor)
        if not pointer_over_image:
            return

        new_width = self._image_label.width()
        new_height = self._image_label.height()
        viewport = self._scroll.viewport()
        if new_width > viewport.width():
            target_x = image_x * new_width / old_width - viewport_position.x()
            self._scroll.horizontalScrollBar().setValue(round(target_x))
        if new_height > viewport.height():
            target_y = image_y * new_height / old_height - viewport_position.y()
            self._scroll.verticalScrollBar().setValue(round(target_y))

    def _show_actual_size(self) -> None:
        self._fit_mode = False
        self._apply_scale(1.0)

    def _fit_to_window(self) -> None:
        if self._base_pixmap.isNull():
            return
        width = self._scroll.viewport().width() - 4
        height = self._scroll.viewport().height() - 4
        if width <= 0 or height <= 0:
            return
        self._fit_mode = True
        scale = min(
            width / self._base_pixmap.width(),
            height / self._base_pixmap.height(),
        )
        self._apply_scale(scale)

    def showEvent(self, event) -> None:
        super().showEvent(event)
        QTimer.singleShot(0, self._fit_to_window)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        if self._fit_mode:
            QTimer.singleShot(0, self._fit_to_window)


class _ComparisonImagePane(QWidget):
    """다중 비교창에서 한 역할의 파형과 정규화된 뷰 위치를 관리한다."""

    def __init__(
        self,
        role: str,
        extracted: Optional[ExtractedImage],
        workbook_path: str,
        *,
        version: str = "",
        temperature: str = "",
        jump_callback: Optional[Callable[[], None]] = None,
        zoom_request_callback: Optional[
            Callable[[str, float, QPointF], None]
        ] = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.role = role
        self._zoom_request_callback = zoom_request_callback
        self._base_pixmap = QPixmap()
        self._scale = 1.0
        self._fit_mode = True

        layout = QVBoxLayout(self)
        layout.setContentsMargins(3, 3, 3, 3)
        layout.setSpacing(4)
        header = QHBoxLayout()
        self.role_label = QLabel(ROLE_DISPLAY_NAMES.get(role, role.upper()))
        self.role_label.setObjectName("previewTitle")
        self.role_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header.addWidget(self.role_label, stretch=1)
        self.jump_button = QPushButton("엑셀 파형 바로가기")
        self.jump_button.setObjectName("inputActionBtn")
        self.jump_button.setEnabled(jump_callback is not None)
        self.jump_button.setAutoDefault(False)
        self.jump_button.setDefault(False)
        self.jump_button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        if jump_callback is not None:
            self.jump_button.clicked.connect(jump_callback)
        header.addWidget(self.jump_button)
        layout.addLayout(header)

        filename = os.path.basename(workbook_path) if workbook_path else "-"
        location = extracted.location_text if extracted is not None else "이미지 없음"
        self.metadata = QLabel(f"{filename} · {location}")
        self.metadata.setObjectName("imageMeta")
        self.metadata.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.metadata.setWordWrap(True)
        layout.addWidget(self.metadata)

        self.attribute_metadata = QLabel(
            waveform_attribute_text(version, temperature)
        )
        self.attribute_metadata.setObjectName("waveformAttributeMeta")
        self.attribute_metadata.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.attribute_metadata)

        self._scroll = _ZoomScrollArea(self._zoom_at)
        self._scroll.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._scroll.setToolTip(
            "Ctrl + 마우스 휠: 모든 파형 확대·축소 / "
            "Ctrl + 좌클릭 드래그: 모든 파형 이동"
        )
        self._image_label = QLabel()
        self._image_label.setObjectName("previewPane")
        self._image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._scroll.setWidget(self._image_label)
        layout.addWidget(self._scroll, stretch=1)

        image = load_extracted_pil(extracted)
        if image is None:
            self._scroll.setWidgetResizable(True)
            self._image_label.setText(self._placeholder_text(extracted))
        else:
            self._scroll.setWidgetResizable(False)
            self._base_pixmap = pil_to_pixmap(
                image,
                max_w=max(1, image.width),
                max_h=max(1, image.height),
                upscale=False,
            )
            self._apply_scale(1.0)
            QTimer.singleShot(0, self.fit_to_window)

    @staticmethod
    def _placeholder_text(extracted: Optional[ExtractedImage]) -> str:
        if extracted is None:
            return "이미지 없음"
        if extracted.placeholder_text == "시트 없음":
            return "현재 Excel에는 이 시트가 없습니다."
        if extracted.cell_address == "-":
            return "이 시트에는 이미지가 없습니다."
        if extracted.is_null:
            return "해당 위치에 이미지가 없습니다."
        return "이미지를 불러올 수 없습니다."

    @property
    def has_image(self) -> bool:
        return not self._base_pixmap.isNull()

    def _apply_scale(self, scale: float) -> None:
        if not self.has_image:
            return
        self._scale = max(0.01, min(8.0, float(scale)))
        width = max(1, int(self._base_pixmap.width() * self._scale))
        height = max(1, int(self._base_pixmap.height() * self._scale))
        scaled = self._base_pixmap.scaled(
            width,
            height,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        self._image_label.setPixmap(scaled)
        self._image_label.resize(scaled.size())

    def zoom_by(self, factor: float) -> None:
        if not self.has_image:
            return
        self._fit_mode = False
        self._apply_scale(self._scale * factor)

    def _zoom_at(self, factor: float, viewport_position: QPointF) -> None:
        if not self.has_image:
            return
        if self._zoom_request_callback is not None:
            self._zoom_request_callback(self.role, factor, viewport_position)
            return

        image_anchor, viewport_anchor = self.zoom_anchors(viewport_position)
        self.zoom_at_normalized(factor, image_anchor, viewport_anchor)

    def zoom_anchors(
        self, viewport_position: QPointF
    ) -> Tuple[Tuple[float, float], Tuple[float, float]]:
        """포인터가 가리키는 이미지/viewport 위치를 0~1 비율로 반환한다."""
        if not self.has_image:
            return (0.5, 0.5), (0.5, 0.5)
        width = self._image_label.width()
        height = self._image_label.height()
        viewport = self._scroll.viewport()
        if width <= 0 or height <= 0:
            return self.normalized_view_center(), (0.5, 0.5)

        image_x = viewport_position.x() - self._image_label.x()
        image_y = viewport_position.y() - self._image_label.y()
        if not (0 <= image_x <= width and 0 <= image_y <= height):
            return self.normalized_view_center(), (0.5, 0.5)
        viewport_width = max(1, viewport.width())
        viewport_height = max(1, viewport.height())
        return (
            image_x / width,
            image_y / height,
        ), (
            max(0.0, min(1.0, viewport_position.x() / viewport_width)),
            max(0.0, min(1.0, viewport_position.y() / viewport_height)),
        )

    def zoom_at_normalized(
        self,
        factor: float,
        image_anchor: Tuple[float, float],
        viewport_anchor: Tuple[float, float],
    ) -> None:
        """정규화된 이미지 지점을 같은 viewport 지점에 유지하며 확대한다."""
        if not self.has_image:
            return
        self.zoom_by(factor)
        self.set_normalized_view_anchor(image_anchor, viewport_anchor)

    def normalized_axis_center(self, axis: str) -> float:
        if not self.has_image:
            return 0.5
        viewport = self._scroll.viewport()
        if axis == "x":
            image_size = self._image_label.width()
            viewport_size = viewport.width()
            image_start = self._image_label.x()
        else:
            image_size = self._image_label.height()
            viewport_size = viewport.height()
            image_start = self._image_label.y()
        if image_size <= 0:
            return 0.5
        image_center = viewport_size / 2.0 - image_start
        return max(0.0, min(1.0, image_center / image_size))

    def normalized_view_center(self) -> Tuple[float, float]:
        return (
            self.normalized_axis_center("x"),
            self.normalized_axis_center("y"),
        )

    def set_normalized_axis_center(self, axis: str, value: float) -> None:
        if not self.has_image:
            return
        normalized = max(0.0, min(1.0, float(value)))
        viewport = self._scroll.viewport()
        if axis == "x":
            image_size = self._image_label.width()
            viewport_size = viewport.width()
            scrollbar = self._scroll.horizontalScrollBar()
        else:
            image_size = self._image_label.height()
            viewport_size = viewport.height()
            scrollbar = self._scroll.verticalScrollBar()
        scrollbar.setValue(round(normalized * image_size - viewport_size / 2.0))

    def set_normalized_view_anchor(
        self,
        image_anchor: Tuple[float, float],
        viewport_anchor: Tuple[float, float],
    ) -> None:
        if not self.has_image:
            return
        viewport = self._scroll.viewport()
        horizontal = self._scroll.horizontalScrollBar()
        vertical = self._scroll.verticalScrollBar()
        horizontal.setValue(
            round(
                max(0.0, min(1.0, image_anchor[0]))
                * self._image_label.width()
                - max(0.0, min(1.0, viewport_anchor[0]))
                * viewport.width()
            )
        )
        vertical.setValue(
            round(
                max(0.0, min(1.0, image_anchor[1]))
                * self._image_label.height()
                - max(0.0, min(1.0, viewport_anchor[1]))
                * viewport.height()
            )
        )

    def show_actual_size(self) -> None:
        if not self.has_image:
            return
        self._fit_mode = False
        self._apply_scale(1.0)

    def fit_to_window(self) -> None:
        if not self.has_image:
            return
        width = self._scroll.viewport().width() - 4
        height = self._scroll.viewport().height() - 4
        if width <= 0 or height <= 0:
            return
        self._fit_mode = True
        self._apply_scale(
            min(
                width / self._base_pixmap.width(),
                height / self._base_pixmap.height(),
            )
        )

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        if self._fit_mode and self.has_image:
            QTimer.singleShot(0, self.fit_to_window)


class ImageComparisonDialog(QDialog):
    """현재 위치의 모든 파형을 동기화된 확대·이동 상태로 표시한다."""

    def __init__(
        self,
        item: InspectionItem,
        mode: str,
        workbook_paths: Dict[str, str],
        parent=None,
        *,
        jump_callback: Optional[Callable[[str], None]] = None,
        jump_enabled_sides: Sequence[str] = (),
        role_order: Optional[Sequence[str]] = None,
        waveform_attributes: Optional[Dict[str, Tuple[str, str]]] = None,
    ) -> None:
        super().__init__(parent)
        if mode not in MODE_ROLES:
            raise ValueError(f"지원하지 않는 검사 모드: {mode}")
        self._mode = mode
        self._item = item
        self.role_order = normalize_preview_order(mode, role_order)
        self._syncing_view = False
        self.panes: Dict[str, _ComparisonImagePane] = {}
        self.pane_positions: Dict[str, Tuple[int, int]] = {}
        self.setWindowTitle(
            f"다른파형 함께보기 · {mode_label(mode)} · "
            f"{item.sheet_name} · {item.cell_address}"
        )
        self.setWindowFlags(
            Qt.WindowType.Window
            | Qt.WindowType.WindowTitleHint
            | Qt.WindowType.WindowSystemMenuHint
            | Qt.WindowType.WindowMinimizeButtonHint
            | Qt.WindowType.WindowMaximizeButtonHint
            | Qt.WindowType.WindowCloseButtonHint
        )
        self.resize(1500, 900)
        self.setMinimumSize(900, 600)

        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)
        toolbar = QHBoxLayout()
        heading = QLabel(
            f"{mode_label(mode)} · {item.sheet_name} · {item.cell_address}"
        )
        heading.setObjectName("positionLabel")
        toolbar.addWidget(heading)
        hint = QLabel(
            "어느 파형에서든 Ctrl+휠: 전체 확대/축소 · "
            "스크롤/Ctrl+드래그: 전체 이동"
        )
        hint.setObjectName("keyboardHint")
        toolbar.addWidget(hint)
        toolbar.addStretch(1)
        zoom_in = QPushButton("전체 확대 +")
        zoom_out = QPushButton("전체 축소 −")
        fit = QPushButton("전체 창에 맞춤")
        actual = QPushButton("전체 실제 크기")
        close_button = QPushButton("닫기")
        zoom_in.clicked.connect(lambda: self._zoom_all(1.25))
        zoom_out.clicked.connect(lambda: self._zoom_all(1.0 / 1.25))
        fit.clicked.connect(self._fit_all)
        actual.clicked.connect(self._show_all_actual_size)
        close_button.clicked.connect(self.accept)
        for button in (zoom_in, zoom_out, fit, actual, close_button):
            button.setObjectName("inputActionBtn")
            toolbar.addWidget(button)
        root.addLayout(toolbar)

        self._grid = QGridLayout()
        self._grid.setContentsMargins(0, 0, 0, 0)
        self._grid.setHorizontalSpacing(6)
        self._grid.setVerticalSpacing(6)
        root.addLayout(self._grid, stretch=1)

        extracted_by_role = dict(item.side_images())
        attributes_by_role = dict(waveform_attributes or {})
        enabled_sides = set(jump_enabled_sides)
        roles = self.role_order
        positions = (
            ((0, 0), (0, 1), (1, 0), (1, 1))
            if mode == "quadra"
            else tuple((0, index) for index in range(len(roles)))
        )
        for role, (row, column) in zip(roles, positions):
            pane_jump = None
            if jump_callback is not None and role in enabled_sides:
                pane_jump = lambda owned_role=role: jump_callback(owned_role)
            version, temperature = attributes_by_role.get(
                role,
                parse_waveform_attributes_from_path(workbook_paths.get(role, "")),
            )
            pane = _ComparisonImagePane(
                role,
                extracted_by_role.get(role),
                workbook_paths.get(role, ""),
                version=version,
                temperature=temperature,
                jump_callback=pane_jump,
                zoom_request_callback=self._on_pane_zoom_requested,
                parent=self,
            )
            self.panes[role] = pane
            self.pane_positions[role] = (row, column)
            self._grid.addWidget(pane, row, column)
            pane._scroll.horizontalScrollBar().valueChanged.connect(
                lambda _value, owned_role=role: self._sync_scroll_from(
                    owned_role, "x"
                )
            )
            pane._scroll.verticalScrollBar().valueChanged.connect(
                lambda _value, owned_role=role: self._sync_scroll_from(
                    owned_role, "y"
                )
            )

        row_count = 2 if mode == "quadra" else 1
        column_count = 2 if mode == "quadra" else len(roles)
        for row in range(row_count):
            self._grid.setRowStretch(row, 1)
        for column in range(column_count):
            self._grid.setColumnStretch(column, 1)

        QShortcut(QKeySequence("+"), self, activated=lambda: self._zoom_all(1.25))
        QShortcut(QKeySequence("-"), self, activated=lambda: self._zoom_all(0.8))
        QShortcut(QKeySequence("0"), self, activated=self._fit_all)

    def _image_panes(self) -> List[_ComparisonImagePane]:
        return [pane for pane in self.panes.values() if pane.has_image]

    def _shared_view_center(self) -> Tuple[float, float]:
        panes = self._image_panes()
        return panes[0].normalized_view_center() if panes else (0.5, 0.5)

    def _effective_zoom_factor(self, requested_factor: float) -> float:
        """모든 파형이 같은 배율만큼 움직이도록 공통 한계를 적용한다."""
        panes = self._image_panes()
        factor = float(requested_factor)
        if not panes or factor == 1.0:
            return 1.0
        if factor > 1.0:
            return max(
                1.0,
                min(factor, *(8.0 / pane._scale for pane in panes)),
            )
        return min(
            1.0,
            max(factor, *(0.01 / pane._scale for pane in panes)),
        )

    def _on_pane_zoom_requested(
        self,
        role: str,
        factor: float,
        viewport_position: QPointF,
    ) -> None:
        pane = self.panes.get(role)
        if pane is None or not pane.has_image:
            return
        image_anchor, viewport_anchor = pane.zoom_anchors(viewport_position)
        self._zoom_all(
            factor,
            image_anchor=image_anchor,
            viewport_anchor=viewport_anchor,
        )

    def _zoom_all(
        self,
        factor: float,
        *,
        image_anchor: Optional[Tuple[float, float]] = None,
        viewport_anchor: Tuple[float, float] = (0.5, 0.5),
    ) -> None:
        panes = self._image_panes()
        if not panes:
            return
        if image_anchor is None:
            image_anchor = self._shared_view_center()
        effective_factor = self._effective_zoom_factor(factor)
        self._syncing_view = True
        try:
            for pane in panes:
                pane.zoom_at_normalized(
                    effective_factor,
                    image_anchor,
                    viewport_anchor,
                )
        finally:
            self._syncing_view = False

    def _sync_scroll_from(self, source_role: str, axis: str) -> None:
        if self._syncing_view:
            return
        source = self.panes.get(source_role)
        if source is None or not source.has_image:
            return
        normalized_center = source.normalized_axis_center(axis)
        self._syncing_view = True
        try:
            for role, pane in self.panes.items():
                if role != source_role and pane.has_image:
                    pane.set_normalized_axis_center(axis, normalized_center)
        finally:
            self._syncing_view = False

    def _fit_all(self) -> None:
        self._syncing_view = True
        try:
            for pane in self.panes.values():
                pane.fit_to_window()
        finally:
            self._syncing_view = False

    def _show_all_actual_size(self) -> None:
        view_center = self._shared_view_center()
        self._syncing_view = True
        try:
            for pane in self.panes.values():
                pane.show_actual_size()
                pane.set_normalized_view_anchor(view_center, (0.5, 0.5))
        finally:
            self._syncing_view = False

    def exec_maximized(self) -> int:
        self.showMaximized()
        return self.exec()


def create_preview_pane(
    default_title: str,
) -> Tuple[QWidget, QLabel, QLabel, ClickableLabel, QPushButton]:
    """셀 주소 옆에 엑셀 바로가기 버튼이 있는 미리보기 칸을 만든다."""
    container = _PreviewDropContainer()
    container.setMinimumWidth(0)
    container.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Ignored)
    layout = QVBoxLayout(container)
    layout.setContentsMargins(2, 2, 2, 2)
    layout.setSpacing(4)
    title = _DraggablePreviewTitle(default_title)
    title.setObjectName("previewTitle")
    title.setAlignment(Qt.AlignmentFlag.AlignCenter)
    title.setWordWrap(True)
    title.setFixedHeight(40)
    metadata = QLabel("-")
    metadata.setObjectName("imageMeta")
    metadata.setAlignment(
        Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter
    )
    metadata.setWordWrap(True)
    metadata.setSizePolicy(
        QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
    )
    jump_button = QPushButton("엑셀 파형 바로가기")
    jump_button.setObjectName("inputActionBtn")
    jump_button.setToolTip(
        "하이퍼링크 모드가 켜져 있으면 이 Excel의 해당 시트·셀로 이동합니다."
    )
    jump_button.setEnabled(False)
    jump_button.setCursor(Qt.CursorShape.PointingHandCursor)
    jump_button.setAutoDefault(False)
    jump_button.setDefault(False)
    jump_button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
    jump_button.setSizePolicy(
        QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed
    )
    meta_row = QHBoxLayout()
    meta_row.setContentsMargins(0, 0, 0, 0)
    meta_row.setSpacing(8)
    meta_row.addWidget(metadata, stretch=1)
    meta_row.addWidget(
        jump_button,
        stretch=0,
        alignment=Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
    )
    image = ClickableLabel("이미지 없음")
    image.setObjectName("previewPane")
    image.setCursor(Qt.CursorShape.PointingHandCursor)
    layout.addWidget(title)
    layout.addLayout(meta_row)
    layout.addWidget(image, stretch=1)
    return container, title, metadata, image, jump_button


def create_review_preview_pane(
    default_title: str,
) -> Tuple[
    QWidget,
    QLabel,
    QLabel,
    ClickableLabel,
    QCheckBox,
    QPushButton,
    QPushButton,
    PreviewAttributeControls,
]:
    """리스트 창용 불량·메모 컨트롤이 포함된 미리보기 칸."""
    container, title, metadata, image, jump_button = create_preview_pane(
        default_title
    )
    container.setObjectName("reviewPreviewContainer")

    container_layout = container.layout()
    title_item = container_layout.takeAt(0)
    if title_item is None or title_item.widget() is not title:
        raise RuntimeError("미리보기 제목 레이아웃을 구성할 수 없습니다.")
    title_header = QGridLayout()
    title_header.setContentsMargins(0, 0, 0, 0)
    title_header.setHorizontalSpacing(6)
    title_header.setVerticalSpacing(1)
    title_header.addWidget(title, 0, 0, 2, 1)
    title_header.setColumnStretch(0, 1)

    version_label = QLabel("버전")
    version_label.setObjectName("previewAttributeLabel")
    version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    version_edit = _CompactAttributeLineEdit()
    version_edit.setObjectName("previewVersionEdit")
    version_edit.setAlignment(Qt.AlignmentFlag.AlignCenter)
    version_edit.setPlaceholderText("M0E0")
    version_edit.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
    version_edit.setToolTip(
        "버전을 자유롭게 입력합니다. 문자·숫자·기호를 모두 사용할 수 있으며, "
        "긴 값은 입력칸 안에서 좌우로 이동합니다."
    )
    title_header.addWidget(version_label, 0, 1)
    title_header.addWidget(version_edit, 1, 1)

    temperature_label = QLabel("온도")
    temperature_label.setObjectName("previewAttributeLabel")
    temperature_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    temperature_edit = _CompactAttributeLineEdit()
    temperature_edit.setObjectName("previewTemperatureEdit")
    temperature_edit.setAlignment(Qt.AlignmentFlag.AlignCenter)
    temperature_edit.setPlaceholderText("ROOM")
    temperature_edit.setSizePolicy(
        QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed
    )
    temperature_edit.setToolTip(
        "온도 문구를 자유롭게 입력합니다. 문자·숫자·기호를 모두 사용할 수 있습니다."
    )
    title_header.addWidget(temperature_label, 0, 2)
    title_header.addWidget(temperature_edit, 1, 2)
    container_layout.insertLayout(0, title_header)
    attribute_controls = PreviewAttributeControls(
        version_edit=version_edit,
        temperature_edit=temperature_edit,
    )

    defect_check = QCheckBox("불량")
    defect_check.setObjectName("defectCheck")
    defect_check.setToolTip("이 이미지가 불량일 때만 체크합니다.")
    memo_button = QPushButton("메모")
    memo_button.setObjectName("memoButton")
    memo_button.setToolTip("불량 체크와 관계없이 메모를 작성할 수 있습니다.")
    memo_button.setAutoDefault(False)
    memo_button.setDefault(False)
    memo_button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
    metadata_row = container_layout.itemAt(1).layout()
    metadata_row.insertWidget(
        1, defect_check, 0, Qt.AlignmentFlag.AlignVCenter
    )
    metadata_row.insertWidget(
        2, memo_button, 0, Qt.AlignmentFlag.AlignVCenter
    )
    return (
        container,
        title,
        metadata,
        image,
        defect_check,
        memo_button,
        jump_button,
        attribute_controls,
    )


class ReviewMemoDialog(QDialog):
    """이미지를 보면서 독립적으로 작성하는 10줄 메모 플로팅 창."""

    def __init__(
        self,
        title: str,
        note: str,
        save_callback: Callable[[str], None],
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._save_callback = save_callback
        self._last_saved = note
        self.setWindowTitle("검토 메모")
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        self.setWindowModality(Qt.WindowModality.NonModal)
        self.setWindowFlags(
            Qt.WindowType.Window
            | Qt.WindowType.WindowTitleHint
            | Qt.WindowType.WindowSystemMenuHint
            | Qt.WindowType.WindowMinMaxButtonsHint
            | Qt.WindowType.WindowCloseButtonHint
        )
        self.resize(560, 360)
        self.setMinimumSize(460, 300)

        layout = QVBoxLayout(self)
        heading = QLabel(title)
        heading.setObjectName("locationLabel")
        heading.setWordWrap(True)
        layout.addWidget(heading)
        hint = QLabel(
            "불량 체크와 별개로 저장됩니다. 10줄을 넘으면 스크롤되며 입력은 자동 저장됩니다."
        )
        hint.setObjectName("keyboardHint")
        hint.setWordWrap(True)
        layout.addWidget(hint)
        self.text_edit = QPlainTextEdit()
        self.text_edit.setObjectName("reviewMemoEdit")
        self.text_edit.setPlaceholderText("관찰 내용, 이상 위치, 재확인 사항 등을 입력하세요.")
        line_height = self.text_edit.fontMetrics().lineSpacing()
        self.text_edit.setMinimumHeight(line_height * 10 + 24)
        self.text_edit.setPlainText(note)
        layout.addWidget(self.text_edit, stretch=1)

        buttons = QHBoxLayout()
        self.save_status = QLabel("자동 저장")
        self.save_status.setObjectName("keyboardHint")
        buttons.addWidget(self.save_status)
        buttons.addStretch(1)
        delete_button = QPushButton("메모 삭제")
        delete_button.setObjectName("memoDeleteButton")
        delete_button.clicked.connect(self._delete_note)
        close_button = QPushButton("닫기")
        close_button.clicked.connect(self.close)
        buttons.addWidget(delete_button)
        buttons.addWidget(close_button)
        layout.addLayout(buttons)

        self._save_timer = QTimer(self)
        self._save_timer.setSingleShot(True)
        self._save_timer.setInterval(350)
        self._save_timer.timeout.connect(self.flush)
        self.text_edit.textChanged.connect(self._schedule_save)

    def _schedule_save(self) -> None:
        self.save_status.setText("저장 중…")
        self._save_timer.start()

    def _delete_note(self) -> None:
        self.text_edit.clear()
        self.flush()

    def flush(self) -> None:
        self._save_timer.stop()
        note = self.text_edit.toPlainText()
        if note != self._last_saved:
            self._save_callback(note)
            self._last_saved = note
        self.save_status.setText("자동 저장됨")

    def closeEvent(self, event) -> None:
        self.flush()
        super().closeEvent(event)


def start_excel_jump(
    parent,
    item: InspectionItem,
    side: str,
    workbook_paths: Dict[str, str],
    *,
    error_cb=None,
) -> None:
    """해당 미리보기 칸의 Excel 시트·셀로 이동한다."""
    target = jump_target_for_side(item, side, workbook_paths)
    if target is None:
        return
    path, sheet_name, cell_address = target
    if not os.path.isfile(path):
        QMessageBox.warning(
            parent,
            "Excel 열기",
            f"Excel 파일을 찾을 수 없습니다:\n{path}",
        )
        return
    jump_to_excel_cell(
        path,
        sheet_name,
        cell_address,
        error_cb=error_cb,
    )


USAGE_HELP_TEXT = """OSC 파형 수동 비교기 사용 방법

이 프로그램은 Excel에 들어 있는 파형 이미지를 같은 위치끼리 나란히 보여 줍니다.
자동으로 점수를 내거나 PASS/FAIL을 판정하지 않습니다. 화면을 보며 직접 비교하는 도구입니다.


1. 검사 모드 선택
• Double (Excel 2개): Excel Ref + 비교A
• Triple (Excel 3개): Excel Ref + 비교A + 비교B
• Quadra (Excel 4개): Excel Ref + 비교A + 비교B + 비교C
모드는 하나만 선택할 수 있습니다. Triple이면 비교B 칸이, Quadra이면 비교B·비교C 칸이 나타납니다.
선택한 모드는 다음에 프로그램을 열 때도 유지됩니다.


2. Excel 파일 지정
지원 파일은 .xlsx, .xlsm 입니다.
각 칸에 경로를 입력하거나, 파일을 끌어다 놓거나, [파일 선택]으로 지정합니다.
[경로 reset]은 해당 칸의 현재·저장 경로를 지우고, 다음 [파일 선택]을 D:\에서 시작합니다.
• Excel Ref: 기준 파일
• Excel A / Excel B / Excel C: 나란히 볼 비교 파일
시트 수나 탭 순서가 달라도 사용할 수 있습니다. 프로그램은 시트 이름으로 연결합니다.
메뉴 [파일] → [Ref Excel 선택...] (Ctrl+O)으로 Ref 파일만 고를 수도 있습니다.


3. 시트 순서 설정 (선택)
필요한 파일을 지정하면 [시트순서설정]에서 현재 자동 매핑 전체를 행으로 확인할 수 있습니다.
• 같은 이름의 시트와 ONLY 시트가 모두 표시됩니다.
• 역할별 목록에서 실제 시트명을 선택하면 이름이 다른 시트를 같은 비교 그룹으로 연결합니다.
• [비교 그룹 추가]: 빈 그룹을 만들고 REF/비교A/비교B/비교C 시트를 선택합니다.
• [위로] / [아래로]: Review 실행 순서를 변경합니다.
• [자동 매핑 복원]: 프로그램의 기본 매핑과 순서로 돌아갑니다.
설정을 바꾸지 않으면 현재 프로그램의 자동 매핑을 그대로 사용합니다.


4. 이미지 불러오기
필요한 파일을 모두 지정한 뒤 [이미지 불러오기] 또는 Ctrl+Enter를 누릅니다.
불러오는 동안 [취소]로 중단할 수 있습니다.
모든 Excel의 시트 이름 합집합을 Ref 순서 우선으로 구성합니다.
정확한 시트 이름을 먼저 연결하고, 일대일로 명확할 때만 앞뒤 공백과 영문 대소문자를 보정합니다.
한 파일에만 있는 시트와 이미지가 0개인 시트도 Review 목록에서 제외하지 않습니다.
이미지는 픽셀을 비교하지 않고, 아래 순서로 같은 위치로 맞춰 묶습니다.
  1) 정확히 같은 셀
  2) 같은 병합영역
  3) 행·열이 각각 1칸 이내인 인접 셀
  4) 한쪽에만 있으면 다른 칸은 '이미지 없음'으로 표시
해당 Excel에 시트 자체가 없으면 '시트 없음'으로 표시해 '이미지 없음'과 구분합니다.
같은 셀에 이미지가 여러 장 있으면 모두 보존해 'D693 (1/2)', 'D693 (2/2)'처럼 표시하고 해당 시트에 확인 경고를 남깁니다.
Queue 생성 뒤 Sheet 역할 관계와 Source Image ID를 검사하며, 원본 이미지의 Queue 누락·중복 배치·역할 오배치가 있으면 Review를 시작하지 않습니다.


5. 메인 화면에서 넘기기
불러온 뒤 같은 위치의 이미지가 한 화면에 나란히 보입니다.
• [처음] [이전] [다음] [마지막] 버튼
• 왼쪽 방향키, PageUp: 이전
• 오른쪽 방향키, PageDown, Space: 다음
• Home: 처음 / End: 마지막
• 상단 [시트] 목록: 시트별 [REF + A], [REF ONLY], [C ONLY] 상태를 보고 이동
• 슬라이더: 전체 위치 중 원하는 곳으로 바로 이동
미리보기의 시트명·셀주소는 가운데에, [엑셀 파형 바로가기]는 오른쪽 끝에 있습니다.
한쪽에만 이미지가 없어도 위치는 건너뛰지 않습니다.


6. 리스트 창
[리스트실행]을 누르면 목록과 미리보기가 함께 있는 창이 열립니다.
• 위: 전체 / 시트 / No. / Ref 셀 / 비교A 셀 (Triple이면 비교B 셀, Quadra이면 비교C 셀까지)
• 아래: 선택한 행의 이미지. Double·Triple은 가로로, Quadra는 Ref|A / B|C 2×2
• 미리보기: 시트명·셀주소는 가운데, [엑셀 파형 바로가기]는 오른쪽 끝
• 각 Excel 제목 오른쪽에서 자유 형식 버전과 온도 문구를 확인·수정합니다.
  파일명에 M숫자E숫자 또는 MVT숫자-숫자와 온도가 있으면 처음 열 때 자동으로 입력됩니다.
• Excel 이름 제목을 다른 이미지 칸으로 드래그하면 두 칸의 표시 위치가 서로 바뀝니다.
• 바꾼 순서는 행·시트를 이동하거나 리스트를 다시 열어도 유지되며, 새 Excel 묶음을 불러오면 기본 순서로 돌아갑니다.
• 한 번 클릭 또는 ↑/↓: 미리보기만 바뀝니다. Excel은 열리지 않습니다.
• 상단 [시트 선택]: 전체 시트 또는 특정 시트만 목록에 표시
제목 표시줄을 더블클릭하면 최대화할 수 있습니다.


7. 하이퍼링크 모드
메인 화면의 [시트] 옆과 리스트 창의 [시트 선택] 옆에 같은 [하이퍼링크 모드] 체크박스가 있습니다. 기본값은 꺼짐입니다.
한쪽에서 켜거나 끄면 다른 쪽에도 같이 적용되고, 다음에 프로그램을 열 때도 유지됩니다.
• 꺼짐: 더블클릭이나 바로가기 버튼을 눌러도 Excel을 열지 않습니다. 바로가기 버튼은 비활성입니다.
• 켜짐: 아래 방법으로 해당 Excel을 열고 그 시트·셀로 이동합니다.
  - 리스트의 Ref 셀 / 비교A 셀 / 비교B 셀 / 비교C 셀 더블클릭
  - 메인·리스트 미리보기 오른쪽 끝의 [엑셀 파형 바로가기]
  - 확대 팝업 오른쪽 위의 [엑셀 파형 바로가기]
  - 전체, 시트, No. 칸을 더블클릭해도 Excel은 열리지 않습니다.
  - '시트 없음'과 '이미지 없음'은 열지 않습니다.
  - 이동한 셀은 선택 상태로 유지하며 위쪽 7개 행을 함께 표시합니다.
  - 이미 Excel이 열려 있으면 새 창을 또 열지 않고, 그 창에서 해당 파일·칸으로 이동합니다.
  - 두 번째 이후에도 Excel 창이 맨 앞으로 올라옵니다.
목록을 넘기는 속도는 하이퍼링크를 켜도 거의 같습니다. Excel은 더블클릭하거나 바로가기 버튼을 누른 순간에만 엽니다.


8. 이미지 확대
메인 화면이나 리스트 창에서 미리보기 이미지를 클릭하면 확대 창이 열립니다.
• [확대 +] / [축소 −] 또는 + / − 키
• Ctrl + 마우스 휠로 포인터 위치를 중심으로 확대/축소
• Ctrl + 좌클릭 드래그로 확대된 이미지 이동
• [창에 맞춤] 또는 0 키
• [실제 크기]
• 상단 정보: 해당 파형의 Ref/비교 역할, Excel 파일명·위치, 버전·온도를 읽기 전용으로 표시
• [다른파형 함께보기]: 현재 위치의 Ref/비교A/비교B/비교C를 모드에 맞춰 최대화 창으로 표시
  - 리스트에서 순서를 바꿨다면 함께보기와 메인 미리보기도 같은 배열을 사용합니다.
  - 각 파형의 버전·온도는 리스트 입력값을 읽기 전용으로 표시합니다.
• 오른쪽 위 [엑셀 파형 바로가기]: 하이퍼링크 모드가 켜져 있으면 그 이미지의 Excel 칸으로 이동합니다.
창 크기를 바꾸면 맞춤 모드일 때 이미지도 다시 맞춰집니다.
함께보기 창은 제목 표시줄을 더블클릭해 최대화하거나 원래 크기로 복원할 수 있습니다.
함께보기 창에서는 어느 파형에서 확대·축소하거나 스크롤바·일반 휠·Ctrl 드래그로 이동해도 모든 파형이 같은 비율 위치로 함께 움직입니다.


9. 참고
• 이 프로그램은 유사도 점수, PASS/FAIL, 리포트를 만들지 않습니다.
• 정상 로딩 결과는 원본 이미지의 Queue 누락 0, 중복 배치 0을 내부 검증한 결과입니다.
• 선택한 Excel 경로와 창 위치는 다음에 열 때도 기억합니다.
• 메뉴 [도움말] → [사용 방법]에서 이 안내를 다시 볼 수 있습니다.
"""


class UsageHelpDialog(QDialog):
    """도움말 메뉴에서 여는 사용 방법 창."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("사용 방법")
        self.resize(720, 680)
        self.setMinimumSize(520, 420)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        self.text = QPlainTextEdit()
        self.text.setReadOnly(True)
        self.text.setPlainText(USAGE_HELP_TEXT)
        self.text.setObjectName("usageHelpText")
        layout.addWidget(self.text, stretch=1)
        buttons = QHBoxLayout()
        buttons.addStretch(1)
        close_button = QPushButton("닫기")
        close_button.setObjectName("inputActionBtn")
        close_button.clicked.connect(self.accept)
        buttons.addWidget(close_button)
        layout.addLayout(buttons)


class LoadSummaryDialog(QDialog):
    """불러오기 결과를 제한된 크기의 스크롤 영역에 표시한다."""

    def __init__(self, message: str, *, has_warnings: bool, parent=None):
        super().__init__(parent)
        self.setWindowTitle(
            "불러오기 완료 · 확인 필요"
            if has_warnings
            else "시트 구성 및 무결성 안내"
        )
        self.resize(720, 480)
        self.setMinimumSize(520, 340)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        header = QLabel(
            "이미지 불러오기는 완료되었습니다. 아래 확인 항목을 검토해 주세요."
            if has_warnings
            else "이미지 불러오기 및 무결성 검사가 완료되었습니다."
        )
        header.setObjectName("positionLabel")
        header.setWordWrap(True)
        layout.addWidget(header)

        self.text = QPlainTextEdit()
        self.text.setReadOnly(True)
        self.text.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
        self.text.setPlainText(message)
        self.text.setObjectName("loadSummaryText")
        layout.addWidget(self.text, stretch=1)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        close_button = QPushButton("확인")
        close_button.setObjectName("inputActionBtn")
        close_button.setDefault(True)
        close_button.clicked.connect(self.accept)
        buttons.addWidget(close_button)
        layout.addLayout(buttons)


class ImageListWindow(QDialog):
    """Excel 위치 목록과 해당 Double/Triple 이미지를 동시에 표시한다."""

    excel_jump_failed = pyqtSignal(str)
    hyperlink_mode_changed = pyqtSignal(bool)
    preview_order_changed = pyqtSignal(object)
    _INDEX_COLUMN_WIDTH = 58
    _NO_COLUMN_WIDTH = 72
    _CELL_COLUMN_WIDTH = 92

    def __init__(
        self,
        items: Sequence[InspectionItem],
        mode: str,
        workbook_paths: Dict[str, str],
        *,
        initial_index: int = 0,
        annotation_store: Optional[ReviewAnnotationStore] = None,
        preview_order: Optional[Sequence[str]] = None,
        parent=None,
    ):
        super().__init__(parent)
        if mode not in SUPPORTED_MODES:
            raise ValueError(f"지원하지 않는 검사 모드: {mode}")
        self._all_items = list(items)
        self._items = list(self._all_items)
        self._global_index_by_id = {
            id(item): index for index, item in enumerate(self._all_items)
        }
        self._mode = mode
        self._preview_order = normalize_preview_order(mode, preview_order)
        self._layout_order = list(ROLE_ORDER)
        self._workbook_paths = dict(workbook_paths)
        self.annotation_store = annotation_store or ReviewAnnotationStore(
            self._workbook_paths,
            mode,
            parent=self,
        )
        self._memo_dialog: Optional[ReviewMemoDialog] = None
        self._current_index = -1
        self._closing = False
        self._quadra_list_sizes_applied = False
        label = mode_label(mode)
        filenames = [
            os.path.basename(path)
            for path in self._workbook_paths.values()
            if path
        ]
        files_text = f" · {', '.join(filenames)}" if filenames else ""
        self.setWindowTitle(
            f"Excel 이미지 리스트 · {label} · {len(self._all_items)}개{files_text}"
        )
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        self.setWindowFlags(
            Qt.WindowType.Window
            | Qt.WindowType.WindowTitleHint
            | Qt.WindowType.WindowSystemMenuHint
            | Qt.WindowType.WindowMinimizeButtonHint
            | Qt.WindowType.WindowMaximizeButtonHint
            | Qt.WindowType.WindowCloseButtonHint
        )
        self.resize(1550, 880)
        self.setMinimumSize(980, 620)

        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(7)

        header = QHBoxLayout()
        title = QLabel(f"Excel 이미지 리스트 · {label}")
        title.setObjectName("positionLabel")
        header.addWidget(title)
        header.addSpacing(14)
        header.addWidget(QLabel("시트 선택"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.setMinimumWidth(360)
        self.sheet_combo.setToolTip("목록에 표시할 Excel 시트를 선택합니다.")
        header.addWidget(self.sheet_combo)
        self.hyperlink_check = QCheckBox("하이퍼링크 모드")
        self.hyperlink_check.setToolTip(
            "켜면 Ref/비교 셀을 더블클릭해 Excel에서 해당 시트와 칸으로 이동합니다."
        )
        self.hyperlink_check.setChecked(
            QSettings("excel-image-inspector", "gui").value(
                "list_hyperlink_mode", False, type=bool
            )
        )
        self.hyperlink_check.toggled.connect(self._on_hyperlink_mode_toggled)
        header.addWidget(self.hyperlink_check)
        header.addStretch(1)
        self.position_label = QLabel("0 / 0")
        self.position_label.setObjectName("positionLabel")
        header.addWidget(self.position_label)
        close_button = QPushButton("닫기")
        close_button.clicked.connect(self.close)
        header.addWidget(close_button)
        root.addLayout(header)

        review_bar = QHBoxLayout()
        review_bar.addWidget(QLabel("검토 표시"))
        self.review_count_label = QLabel("기록 0건 · 불량 0건 · 메모만 0건")
        self.review_count_label.setObjectName("reviewCountLabel")
        review_bar.addWidget(self.review_count_label)
        review_bar.addSpacing(12)
        review_bar.addWidget(QLabel("목록 필터"))
        self.record_filter_combo = QComboBox()
        self.record_filter_combo.addItem("전체", "all")
        self.record_filter_combo.addItem("기록 있는 항목", "recorded")
        self.record_filter_combo.addItem("불량 항목", "defect")
        self.record_filter_combo.addItem("메모 있는 항목", "memo")
        self.record_filter_combo.setToolTip(
            "현재 시트 필터와 함께 적용됩니다. 한 이미지라도 조건에 맞으면 표시합니다."
        )
        review_bar.addWidget(self.record_filter_combo)
        review_bar.addStretch(1)
        self.export_button = QPushButton("검토 결과 Excel 추출")
        self.export_button.setObjectName("reviewExportButton")
        self.export_button.setToolTip(
            "불량 체크 또는 메모가 있는 이미지만 새 Excel 파일로 추출합니다."
        )
        self.export_button.clicked.connect(self._export_review_results)
        review_bar.addWidget(self.export_button)
        root.addLayout(review_bar)

        self.location_label = QLabel("목록에서 이미지를 선택해 주세요.")
        self.location_label.setObjectName("locationLabel")
        root.addWidget(self.location_label)

        self.content_splitter = QSplitter(Qt.Orientation.Vertical)
        self.content_splitter.setChildrenCollapsible(False)
        self.table = QTableWidget()
        self.table.setObjectName("imageListTable")
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(False)
        self.table.setWordWrap(False)
        self.table.verticalHeader().setVisible(False)
        self.table.currentCellChanged.connect(self._on_current_cell_changed)
        self.table.cellDoubleClicked.connect(self._on_cell_double_clicked)
        self.excel_jump_failed.connect(self._on_excel_jump_failed)
        self.content_splitter.addWidget(self.table)

        self.preview_widget = QWidget()
        preview_layout = QVBoxLayout(self.preview_widget)
        preview_layout.setContentsMargins(0, 4, 0, 0)
        preview_layout.setSpacing(5)
        preview_hint = QLabel(
            "목록에서 ↑/↓ 키 또는 마우스로 이동하면 이미지가 함께 바뀝니다. · "
            "Excel 이름을 다른 이미지 칸으로 드래그하면 두 칸의 위치가 바뀝니다."
        )
        preview_hint.setObjectName("keyboardHint")
        preview_layout.addWidget(preview_hint)

        (
            self.ref_container,
            self.ref_title,
            self.ref_meta,
            self.ref_image,
            self.ref_defect_check,
            self.ref_memo_button,
            self.ref_jump_button,
            self.ref_attribute_controls,
        ) = self._create_preview_pane("Excel Ref")
        (
            self.a_container,
            self.a_title,
            self.a_meta,
            self.a_image,
            self.a_defect_check,
            self.a_memo_button,
            self.a_jump_button,
            self.a_attribute_controls,
        ) = self._create_preview_pane("Excel 비교A")
        (
            self.b_container,
            self.b_title,
            self.b_meta,
            self.b_image,
            self.b_defect_check,
            self.b_memo_button,
            self.b_jump_button,
            self.b_attribute_controls,
        ) = self._create_preview_pane("Excel 비교B")
        (
            self.c_container,
            self.c_title,
            self.c_meta,
            self.c_image,
            self.c_defect_check,
            self.c_memo_button,
            self.c_jump_button,
            self.c_attribute_controls,
        ) = self._create_preview_pane("Excel 비교C")
        self._preview_containers = {
            "ref": self.ref_container,
            "a": self.a_container,
            "b": self.b_container,
            "c": self.c_container,
        }
        self._preview_titles = {
            "ref": self.ref_title,
            "a": self.a_title,
            "b": self.b_title,
            "c": self.c_title,
        }
        self._preview_attribute_controls = {
            "ref": self.ref_attribute_controls,
            "a": self.a_attribute_controls,
            "b": self.b_attribute_controls,
            "c": self.c_attribute_controls,
        }
        for role, controls in self._preview_attribute_controls.items():
            controls.set_values(
                *parse_waveform_attributes_from_path(
                    self._workbook_paths.get(role, "")
                )
            )
        for role in MODE_ROLES[mode]:
            title_widget = self._preview_titles[role]
            container_widget = self._preview_containers[role]
            if isinstance(title_widget, _DraggablePreviewTitle):
                title_widget.enable_role_drag(role)
            if isinstance(container_widget, _PreviewDropContainer):
                container_widget.enable_role_drop(role)
                container_widget.preview_swap_requested.connect(
                    self._swap_preview_roles
                )
        self.ref_image.clicked.connect(lambda: self._enlarge("ref"))
        self.a_image.clicked.connect(lambda: self._enlarge("a"))
        self.b_image.clicked.connect(lambda: self._enlarge("b"))
        self.c_image.clicked.connect(lambda: self._enlarge("c"))
        self.ref_jump_button.clicked.connect(lambda: self._jump_from_preview("ref"))
        self.a_jump_button.clicked.connect(lambda: self._jump_from_preview("a"))
        self.b_jump_button.clicked.connect(lambda: self._jump_from_preview("b"))
        self.c_jump_button.clicked.connect(lambda: self._jump_from_preview("c"))
        for side, checkbox, memo_button in (
            ("ref", self.ref_defect_check, self.ref_memo_button),
            ("a", self.a_defect_check, self.a_memo_button),
            ("b", self.b_defect_check, self.b_memo_button),
            ("c", self.c_defect_check, self.c_memo_button),
        ):
            checkbox.toggled.connect(
                lambda checked, owned_side=side: self._toggle_defect(
                    owned_side, checked
                )
            )
            memo_button.clicked.connect(
                lambda _checked=False, owned_side=side: self._open_memo_editor(
                    owned_side
                )
            )
        self.b_container.setVisible(mode in {"triple", "quadra"})
        self.c_container.setVisible(mode == "quadra")
        if mode == "quadra":
            self.preview_splitter = None
            self.quadra_top = QSplitter(Qt.Orientation.Horizontal)
            self.quadra_bottom = QSplitter(Qt.Orientation.Horizontal)
            self.quadra_preview = QSplitter(Qt.Orientation.Vertical)
            self.quadra_top.addWidget(self.ref_container)
            self.quadra_top.addWidget(self.a_container)
            self.quadra_bottom.addWidget(self.b_container)
            self.quadra_bottom.addWidget(self.c_container)
            for row_splitter in (self.quadra_top, self.quadra_bottom):
                row_splitter.setChildrenCollapsible(False)
                row_splitter.setHandleWidth(1)
                row_splitter.setStretchFactor(0, 1)
                row_splitter.setStretchFactor(1, 1)
                row_splitter.handle(1).setEnabled(False)
            self.quadra_preview.addWidget(self.quadra_top)
            self.quadra_preview.addWidget(self.quadra_bottom)
            self.quadra_preview.setChildrenCollapsible(False)
            self.quadra_preview.setHandleWidth(1)
            self.quadra_preview.setStretchFactor(0, 1)
            self.quadra_preview.setStretchFactor(1, 1)
            preview_layout.addWidget(self.quadra_preview, stretch=1)
        else:
            self.preview_splitter = QSplitter(Qt.Orientation.Horizontal)
            for container in (
                self.ref_container,
                self.a_container,
                self.b_container,
                self.c_container,
            ):
                self.preview_splitter.addWidget(container)
            self.preview_splitter.setChildrenCollapsible(False)
            for index in range(4):
                self.preview_splitter.setStretchFactor(index, 1)
            self.preview_splitter.setHandleWidth(1)
            for handle_index in (1, 2, 3):
                self.preview_splitter.handle(handle_index).setEnabled(False)
            preview_layout.addWidget(self.preview_splitter, stretch=1)
        self._apply_preview_order(self._preview_order)
        self.content_splitter.addWidget(self.preview_widget)
        self.content_splitter.setStretchFactor(0, 0)
        self.content_splitter.setStretchFactor(1, 1)
        if mode == "quadra":
            self.content_splitter.setSizes([self._list_table_default_height(), 900])
        else:
            self.content_splitter.setSizes([240, 600])
        root.addWidget(self.content_splitter, stretch=1)

        self._populate_sheet_combo()
        self.sheet_combo.currentIndexChanged.connect(
            self._on_sheet_filter_changed
        )
        self.record_filter_combo.currentIndexChanged.connect(
            self._on_record_filter_changed
        )
        self.annotation_store.changed.connect(self._on_annotations_changed)
        self._populate_table()
        self._update_review_count()
        self.set_current_index(initial_index)

    def _create_preview_pane(
        self, default_title: str
    ) -> Tuple[
        QWidget,
        QLabel,
        QLabel,
        ClickableLabel,
        QCheckBox,
        QPushButton,
        QPushButton,
        PreviewAttributeControls,
    ]:
        return create_review_preview_pane(default_title)

    def keyPressEvent(self, event: QKeyEvent) -> None:
        """Enter가 대화상자의 자동 기본 버튼을 실행하지 않게 한다."""
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            event.accept()
            return
        super().keyPressEvent(event)

    @property
    def preview_order(self) -> Tuple[str, ...]:
        return self._preview_order

    @property
    def waveform_attributes(self) -> Dict[str, Tuple[str, str]]:
        """역할별로 리스트에서 입력한 버전·온도 문구를 반환한다."""
        return {
            role: (controls.version_text, controls.temperature)
            for role, controls in self._preview_attribute_controls.items()
        }

    def _preview_slot(self, index: int) -> Tuple[QSplitter, int]:
        if self._mode == "quadra":
            if index < 2:
                return self.quadra_top, index
            return self.quadra_bottom, index - 2
        return self.preview_splitter, index

    def _apply_preview_order(self, role_order: Sequence[str]) -> None:
        normalized = normalize_preview_order(self._mode, role_order)
        desired_layout = list(normalized) + [
            role for role in ROLE_ORDER if role not in normalized
        ]
        for target_index, target_role in enumerate(desired_layout):
            current_index = self._layout_order.index(target_role)
            if current_index == target_index:
                continue
            current_role = self._layout_order[target_index]
            target_splitter, target_slot = self._preview_slot(target_index)
            current_splitter, current_slot = self._preview_slot(current_index)
            swap_splitter_widgets(
                target_splitter,
                target_slot,
                self._preview_containers[current_role],
                current_splitter,
                current_slot,
                self._preview_containers[target_role],
            )
            self._layout_order[target_index], self._layout_order[current_index] = (
                self._layout_order[current_index],
                self._layout_order[target_index],
            )
        self._preview_order = normalized
        if self._mode == "quadra":
            self.quadra_top.handle(1).setEnabled(False)
            self.quadra_bottom.handle(1).setEnabled(False)
        else:
            for handle_index in range(1, self.preview_splitter.count()):
                self.preview_splitter.handle(handle_index).setEnabled(False)
        self._equalize_preview_panes()

    def set_preview_order(
        self,
        role_order: Sequence[str],
        *,
        emit: bool = False,
    ) -> bool:
        normalized = normalize_preview_order(self._mode, role_order)
        if normalized == self._preview_order:
            return False
        self._apply_preview_order(normalized)
        if emit:
            self.preview_order_changed.emit(self._preview_order)
        return True

    def _swap_preview_roles(self, source_role: str, target_role: str) -> None:
        swapped = swapped_preview_order(
            self._mode,
            self._preview_order,
            source_role,
            target_role,
        )
        if swapped == self._preview_order:
            return
        self._apply_preview_order(swapped)
        self.preview_order_changed.emit(self._preview_order)

    @staticmethod
    def _cell_text(extracted: Optional[ExtractedImage]) -> str:
        if extracted is None:
            return "이미지 없음"
        if extracted.is_null:
            return extracted.placeholder_text
        warning = " ⚠" if extracted.anchor_count > 1 else ""
        return f"{extracted.display_cell_address}{warning}"

    def _populate_sheet_combo(self) -> None:
        counts: Dict[int, int] = {}
        names: Dict[int, str] = {}
        statuses: Dict[int, str] = {}
        warning_counts: Dict[int, int] = {}
        empty_sheets = set()
        for item in self._all_items:
            counts[item.sheet_index] = counts.get(item.sheet_index, 0) + 1
            names.setdefault(item.sheet_index, item.sheet_name)
            statuses.setdefault(item.sheet_index, item.sheet_status_text)
            if item.sheet_info is not None:
                warning_counts.setdefault(
                    item.sheet_index, len(item.sheet_info.image_warnings)
                )
            if item.is_empty_sheet:
                empty_sheets.add(item.sheet_index)
        blocker = QSignalBlocker(self.sheet_combo)
        self.sheet_combo.clear()
        self.sheet_combo.addItem(
            f"전체 시트 ({len(counts)}개) · Review {len(self._all_items)}개", None
        )
        for sheet_index in sorted(counts):
            review_count = 0 if sheet_index in empty_sheets else counts[sheet_index]
            warning = (
                f" · ⚠ 동일 셀 {warning_counts[sheet_index]}곳"
                if warning_counts.get(sheet_index)
                else ""
            )
            self.sheet_combo.addItem(
                f"{sheet_index}. {names[sheet_index]} "
                f"[{statuses[sheet_index]}] ({review_count}개){warning}",
                sheet_index,
            )
        del blocker

    def _on_sheet_filter_changed(self, combo_index: int) -> None:
        if self._closing or combo_index < 0:
            return
        self._apply_filters()
        self.focus_list()

    def _on_record_filter_changed(self, combo_index: int) -> None:
        if self._closing or combo_index < 0:
            return
        self._apply_filters()
        self.focus_list()

    def _apply_filters(self, preferred_item: Optional[InspectionItem] = None) -> None:
        current_item = (
            self._items[self._current_index]
            if 0 <= self._current_index < len(self._items)
            else None
        )
        current_item = preferred_item or current_item
        sheet_index = self.sheet_combo.currentData()
        record_filter = self.record_filter_combo.currentData() or "all"
        self._items = [
            item
            for item in self._all_items
            if (sheet_index is None or item.sheet_index == int(sheet_index))
            and self._item_matches_record_filter(item, str(record_filter))
        ]
        self._current_index = -1
        self._populate_table()
        target = next(
            (
                index
                for index, item in enumerate(self._items)
                if item is current_item
            ),
            0,
        )
        self._set_filtered_index(target)

    def _item_matches_record_filter(
        self, item: InspectionItem, record_filter: str
    ) -> bool:
        if record_filter == "all":
            return True
        annotations = [
            annotation
            for role, image in item.side_images()
            if (annotation := self.annotation_store.annotation_for(role, image))
            is not None
        ]
        if record_filter == "recorded":
            return any(annotation.should_keep for annotation in annotations)
        if record_filter == "defect":
            return any(annotation.is_defect for annotation in annotations)
        if record_filter == "memo":
            return any(annotation.note.strip() for annotation in annotations)
        return True

    def _populate_table(self) -> None:
        headers = ["전체", "시트", "No.", "Ref 셀", "비교A 셀"]
        if self._mode in {"triple", "quadra"}:
            headers.append("비교B 셀")
        if self._mode == "quadra":
            headers.append("비교C 셀")
        self.table.setUpdatesEnabled(False)
        blocker = QSignalBlocker(self.table)
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setRowCount(len(self._items))
        sheet_totals: Dict[int, int] = {}
        for item in self._items:
            sheet_totals[item.sheet_index] = sheet_totals.get(item.sheet_index, 0) + 1
        sheet_positions: Dict[int, int] = {}

        for row, item in enumerate(self._items):
            sheet_positions[item.sheet_index] = sheet_positions.get(item.sheet_index, 0) + 1
            values: List[str] = [
                str(self._global_index_by_id[id(item)] + 1),
                f"{item.sheet_index}. {item.sheet_name} [{item.sheet_status_text}]",
                f"{sheet_positions[item.sheet_index]}/{sheet_totals[item.sheet_index]}",
                self._cell_text(item.image_ref),
                self._cell_text(item.image_a),
            ]
            if self._mode in {"triple", "quadra"}:
                values.append(self._cell_text(item.image_b))
            if self._mode == "quadra":
                values.append(self._cell_text(item.image_c))
            for column, value in enumerate(values):
                table_item = QTableWidgetItem(value)
                table_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                table_item.setData(Qt.ItemDataRole.UserRole, row)
                self.table.setItem(row, column, table_item)

        header = self.table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self._apply_list_column_widths()
        del blocker
        self.table.setUpdatesEnabled(True)

    def set_current_index(self, index: int) -> None:
        """전체 목록 기준 index로 현재 행을 맞춘다."""
        if self._closing:
            return
        if not self._all_items:
            self._current_index = -1
            self._render_current()
            return
        global_index = max(0, min(len(self._all_items) - 1, int(index)))
        target_item = self._all_items[global_index]
        if not any(item is target_item for item in self._items):
            blocker = QSignalBlocker(self.sheet_combo)
            self.sheet_combo.setCurrentIndex(0)
            del blocker
            with QSignalBlocker(self.record_filter_combo):
                self.record_filter_combo.setCurrentIndex(0)
            self._apply_filters(preferred_item=target_item)
        filtered_index = next(
            index
            for index, item in enumerate(self._items)
            if item is target_item
        )
        self._set_filtered_index(filtered_index)

    def _set_filtered_index(self, index: int) -> None:
        if not self._items:
            self._current_index = -1
            self._render_current()
            return
        target = max(0, min(len(self._items) - 1, int(index)))
        self.table.setCurrentCell(target, 0)
        self.table.selectRow(target)
        self.table.scrollToItem(
            self.table.item(target, 0),
            QAbstractItemView.ScrollHint.PositionAtCenter,
        )
        if self._current_index != target:
            self._current_index = target
            self._render_current()

    def _on_current_cell_changed(
        self,
        current_row: int,
        _current_column: int,
        _previous_row: int,
        _previous_column: int,
    ) -> None:
        if self._closing:
            return
        if 0 <= current_row < len(self._items):
            if self._current_index != current_row:
                self._close_memo_editor()
            self._current_index = current_row
            self._render_current()

    def apply_hyperlink_mode(self, checked: bool) -> None:
        with QSignalBlocker(self.hyperlink_check):
            self.hyperlink_check.setChecked(bool(checked))
        self._sync_preview_jump_buttons()

    def _on_hyperlink_mode_toggled(self, checked: bool) -> None:
        QSettings("excel-image-inspector", "gui").setValue(
            "list_hyperlink_mode", bool(checked)
        )
        self._sync_preview_jump_buttons()
        self.hyperlink_mode_changed.emit(bool(checked))

    def _jump_from_preview(self, side: str) -> None:
        if self._closing or not self.hyperlink_check.isChecked():
            return
        if not (0 <= self._current_index < len(self._items)):
            return
        start_excel_jump(
            QApplication.activeWindow() or self,
            self._items[self._current_index],
            side,
            self._workbook_paths,
            error_cb=self.excel_jump_failed.emit,
        )

    def _sync_preview_jump_buttons(self) -> None:
        buttons = {
            "ref": self.ref_jump_button,
            "a": self.a_jump_button,
            "b": self.b_jump_button,
            "c": self.c_jump_button,
        }
        enabled_mode = (
            not self._closing and self.hyperlink_check.isChecked()
        )
        for side, button in buttons.items():
            if side == "b" and self._mode not in {"triple", "quadra"}:
                button.setEnabled(False)
                continue
            if side == "c" and self._mode != "quadra":
                button.setEnabled(False)
                continue
            button.setEnabled(
                enabled_mode and self._preview_jump_available(side)
            )

    @staticmethod
    def _refresh_dynamic_style(widget: QWidget) -> None:
        widget.style().unpolish(widget)
        widget.style().polish(widget)
        widget.update()

    def _current_image(self, side: str) -> Optional[ExtractedImage]:
        if not (0 <= self._current_index < len(self._items)):
            return None
        item = self._items[self._current_index]
        return {
            "ref": item.image_ref,
            "a": item.image_a,
            "b": item.image_b,
            "c": item.image_c,
        }.get(side)

    def _sync_review_controls(self) -> None:
        controls = {
            "ref": (
                self.ref_container,
                self.ref_defect_check,
                self.ref_memo_button,
            ),
            "a": (
                self.a_container,
                self.a_defect_check,
                self.a_memo_button,
            ),
            "b": (
                self.b_container,
                self.b_defect_check,
                self.b_memo_button,
            ),
            "c": (
                self.c_container,
                self.c_defect_check,
                self.c_memo_button,
            ),
        }
        for side, (container, checkbox, memo_button) in controls.items():
            image = self._current_image(side)
            annotation = self.annotation_store.annotation_for(side, image)
            available = annotation is not None
            checkbox.setEnabled(available)
            memo_button.setEnabled(available)
            with QSignalBlocker(checkbox):
                checkbox.setChecked(
                    bool(annotation is not None and annotation.is_defect)
                )
            has_note = bool(annotation is not None and annotation.note.strip())
            memo_button.setText("메모 있음" if has_note else "메모")
            memo_button.setProperty("hasNote", has_note)
            container.setProperty(
                "defectMarked",
                bool(annotation is not None and annotation.is_defect),
            )
            self._refresh_dynamic_style(memo_button)
            self._refresh_dynamic_style(container)

    def _toggle_defect(self, side: str, checked: bool) -> None:
        if self._closing:
            return
        image = self._current_image(side)
        if image is None or image.is_null:
            return
        self.annotation_store.update(side, image, is_defect=checked)

    def _open_memo_editor(self, side: str) -> None:
        if self._closing:
            return
        image = self._current_image(side)
        annotation = self.annotation_store.annotation_for(side, image)
        if image is None or annotation is None:
            return
        self._close_memo_editor()
        role = ROLE_DISPLAY_NAMES.get(side, side)
        title = f"{role} · {image.location_text}"
        dialog = ReviewMemoDialog(
            title,
            annotation.note,
            lambda note, owned_side=side, owned_image=image: self.annotation_store.update(
                owned_side, owned_image, note=note
            ),
            self,
        )
        self._memo_dialog = dialog
        dialog.destroyed.connect(
            lambda _object=None, owned=dialog: self._on_memo_dialog_destroyed(
                owned
            )
        )
        dialog.show()
        dialog.raise_()
        dialog.activateWindow()
        dialog.text_edit.setFocus(Qt.FocusReason.OtherFocusReason)

    def _on_memo_dialog_destroyed(self, dialog: ReviewMemoDialog) -> None:
        if self._memo_dialog is dialog:
            self._memo_dialog = None

    def _close_memo_editor(self) -> None:
        if self._memo_dialog is None:
            return
        dialog = self._memo_dialog
        self._memo_dialog = None
        try:
            dialog.close()
        except RuntimeError:
            pass

    def _update_review_count(self) -> None:
        records = self.annotation_store.records()
        self.review_count_label.setText(
            f"기록 {len(records)}건 · 불량 {self.annotation_store.defect_count}건 · "
            f"메모만 {self.annotation_store.memo_only_count}건"
        )
        self.export_button.setEnabled(bool(records))

    def _on_annotations_changed(self) -> None:
        if self._closing:
            return
        self._update_review_count()
        self._sync_review_controls()
        if (self.record_filter_combo.currentData() or "all") != "all":
            current_item = (
                self._items[self._current_index]
                if 0 <= self._current_index < len(self._items)
                else None
            )
            self._apply_filters(preferred_item=current_item)

    def _export_review_results(self) -> None:
        records = self.annotation_store.records()
        if not records:
            QMessageBox.information(
                self,
                "검토 결과 추출",
                "불량 체크 또는 메모가 있는 기록이 없습니다.",
            )
            return
        first_path = next(
            (path for path in self._workbook_paths.values() if path), ""
        )
        base_dir = os.path.dirname(first_path) if first_path else os.getcwd()
        filename = f"이미지_검토기록_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        output_path, _selected_filter = QFileDialog.getSaveFileName(
            self,
            "검토 결과 Excel 저장",
            os.path.join(base_dir, filename),
            "Excel 통합 문서 (*.xlsx)",
        )
        if not output_path:
            return
        if not output_path.lower().endswith(".xlsx"):
            output_path += ".xlsx"
        try:
            count = export_annotations_xlsx(output_path, records)
        except Exception as exc:
            QMessageBox.critical(
                self,
                "검토 결과 추출 실패",
                f"Excel 파일을 저장하지 못했습니다.\n{exc}",
            )
            return
        QMessageBox.information(
            self,
            "검토 결과 추출 완료",
            f"불량 또는 메모 기록 {count}건을 저장했습니다.\n{output_path}",
        )

    def _preview_jump_available(self, side: str) -> bool:
        if not (0 <= self._current_index < len(self._items)):
            return False
        return (
            jump_target_for_side(
                self._items[self._current_index],
                side,
                self._workbook_paths,
            )
            is not None
        )

    def _on_cell_double_clicked(self, row: int, column: int) -> None:
        if self._closing or not self.hyperlink_check.isChecked():
            return
        if not (0 <= row < len(self._items)):
            return
        target = jump_target_for_list_cell(
            self._items[row], column, self._workbook_paths
        )
        if target is None:
            return
        path, sheet_name, cell_address = target
        if not os.path.isfile(path):
            QMessageBox.warning(
                self,
                "Excel 열기",
                f"Excel 파일을 찾을 수 없습니다:\n{path}",
            )
            return
        jump_to_excel_cell(
            path,
            sheet_name,
            cell_address,
            error_cb=self.excel_jump_failed.emit,
        )

    def _on_excel_jump_failed(self, message: str) -> None:
        if self._closing:
            return
        QMessageBox.warning(self, "Excel 열기", message)

    def _render_current(self) -> None:
        if not (0 <= self._current_index < len(self._items)):
            self.position_label.setText("0 / 0")
            self.location_label.setText("표시할 이미지가 없습니다.")
            self._clear_preview_pane(
                self.ref_title, self.ref_meta, self.ref_image, "Excel Ref"
            )
            self._clear_preview_pane(
                self.a_title, self.a_meta, self.a_image, "Excel 비교A"
            )
            self._clear_preview_pane(
                self.b_title, self.b_meta, self.b_image, "Excel 비교B"
            )
            self._clear_preview_pane(
                self.c_title, self.c_meta, self.c_image, "Excel 비교C"
            )
            self._sync_preview_jump_buttons()
            self._sync_review_controls()
            return

        item = self._items[self._current_index]
        global_index = self._global_index_by_id[id(item)]
        if (
            self.sheet_combo.currentData() is None
            and (self.record_filter_combo.currentData() or "all") == "all"
        ):
            position_text = f"전체 {global_index + 1} / {len(self._all_items)}"
        elif (self.record_filter_combo.currentData() or "all") == "all":
            position_text = (
                f"시트 {self._current_index + 1} / {len(self._items)}"
                f" · 전체 {global_index + 1} / {len(self._all_items)}"
            )
        else:
            position_text = (
                f"목록 {self._current_index + 1} / {len(self._items)}"
                f" · 전체 {global_index + 1} / {len(self._all_items)}"
            )
        self.position_label.setText(position_text)
        duplicate_warning = (
            f" · {item.duplicate_warning_text}"
            if item.duplicate_warning_text
            else ""
        )
        self.location_label.setText(
            f"시트 {item.sheet_index}. {item.sheet_name} [{item.sheet_status_text}] · "
            f"시트 내 {item.index} · 기준 위치 {item.cell_address}"
            f"{duplicate_warning}"
        )
        self._set_preview_pane(
            self.ref_title,
            self.ref_meta,
            self.ref_image,
            "Excel Ref",
            self._workbook_paths.get("ref", ""),
            item.image_ref,
        )
        self._set_preview_pane(
            self.a_title,
            self.a_meta,
            self.a_image,
            "Excel 비교A",
            self._workbook_paths.get("a", ""),
            item.image_a,
        )
        if self._mode in {"triple", "quadra"}:
            if item.image_b is not None:
                self._set_preview_pane(
                    self.b_title,
                    self.b_meta,
                    self.b_image,
                    "Excel 비교B",
                    self._workbook_paths.get("b", ""),
                    item.image_b,
                )
            else:
                self._clear_preview_pane(
                    self.b_title, self.b_meta, self.b_image, "Excel 비교B"
                )
        if self._mode == "quadra":
            if item.image_c is not None:
                self._set_preview_pane(
                    self.c_title,
                    self.c_meta,
                    self.c_image,
                    "Excel 비교C",
                    self._workbook_paths.get("c", ""),
                    item.image_c,
                )
            else:
                self._clear_preview_pane(
                    self.c_title, self.c_meta, self.c_image, "Excel 비교C"
                )
        self._equalize_preview_panes()
        self._sync_preview_jump_buttons()
        self._sync_review_controls()

    def _list_table_default_height(self) -> int:
        visible_rows = 3 if self._mode == "quadra" else 7
        header_height = max(self.table.horizontalHeader().sizeHint().height(), 24)
        row_height = self.table.verticalHeader().defaultSectionSize()
        if self.table.rowCount() > 0:
            row_height = max(row_height, self.table.rowHeight(0))
        return header_height + self.table.frameWidth() * 2 + row_height * visible_rows + 6

    def _apply_quadra_list_sizes(self) -> None:
        table_height = self._list_table_default_height()
        leftover = self.content_splitter.height() - table_height
        preview_height = leftover if leftover > 200 else 700
        self.content_splitter.setSizes([table_height, preview_height])
        self.quadra_top.setSizes([1000, 1000])
        self.quadra_bottom.setSizes([1000, 1000])
        self.quadra_preview.setSizes([1000, 1000])

    def _equalize_preview_panes(self) -> None:
        if self._mode == "quadra":
            self.quadra_top.setSizes([1000, 1000])
            self.quadra_bottom.setSizes([1000, 1000])
            self.quadra_preview.setSizes([1000, 1000])
            return
        sizes = {
            "double": [1000, 1000, 0, 0],
            "triple": [1000, 1000, 1000, 0],
        }.get(self._mode, [1000, 1000, 0, 0])
        self.preview_splitter.setSizes(sizes)

    @staticmethod
    def _clear_preview_pane(
        title_label: QLabel,
        meta_label: QLabel,
        image_label: ClickableLabel,
        title: str,
    ) -> None:
        title_label.setText(title)
        if isinstance(title_label, _DraggablePreviewTitle) and title_label.drag_role:
            title_label.setToolTip(
                f"{title}\n다른 이미지 칸으로 드래그하면 두 칸의 위치가 바뀝니다."
            )
        meta_label.setText("-")
        image_label.setPixmap(None)
        image_label.setText("이미지 없음")

    @staticmethod
    def _set_preview_pane(
        title_label: QLabel,
        meta_label: QLabel,
        image_label: ClickableLabel,
        role: str,
        path: str,
        extracted: ExtractedImage,
    ) -> None:
        filename = os.path.basename(path) if path else "-"
        title_label.setText(f"{role} · {filename}")
        title_tooltip = f"{role} · {filename}"
        if isinstance(title_label, _DraggablePreviewTitle) and title_label.drag_role:
            title_tooltip += "\n다른 이미지 칸으로 드래그하면 두 칸의 위치가 바뀝니다."
        title_label.setToolTip(title_tooltip)
        meta_label.setText(extracted.location_text)
        meta_label.setToolTip(extracted.location_text)
        if extracted.is_null:
            image_label.setPixmap(None)
            if extracted.placeholder_text == "시트 없음":
                image_label.setText("현재 Excel에는 이 시트가 없습니다.")
            elif extracted.cell_address == "-":
                image_label.setText("이 시트에는 이미지가 없습니다.")
            else:
                image_label.setText("해당 위치에 이미지가 없습니다.")
            return
        image = load_extracted_pil(extracted)
        if image is None:
            image_label.setPixmap(None)
            image_label.setText("이미지를 불러올 수 없습니다.")
            return
        image_label.setText("")
        image_label.setPixmap(pil_to_pixmap(image))

    def _enlarge(self, side: str) -> None:
        if not (0 <= self._current_index < len(self._items)):
            return
        item = self._items[self._current_index]
        extracted = {
            "ref": item.image_ref,
            "a": item.image_a,
            "b": item.image_b,
            "c": item.image_c,
        }.get(side)
        if extracted is None or extracted.is_null:
            return
        image = load_extracted_pil(extracted)
        if image is None:
            return
        role = {"ref": "Ref", "a": "비교A", "b": "비교B", "c": "비교C"}[side]
        version, temperature = self.waveform_attributes[side]
        viewer = ImageViewerDialog(
            image,
            f"{role} · {extracted.location_text}",
            self,
            jump_enabled=self.hyperlink_check.isChecked()
            and self._preview_jump_available(side),
            jump_callback=lambda: self._jump_from_preview(side),
            comparison_available=True,
            role=side,
            workbook_path=self._workbook_paths.get(side, ""),
            location_text=extracted.location_text,
            version=version,
            temperature=temperature,
        )
        result = viewer.exec()
        if result == SHOW_COMPARISON_RESULT and not self._closing:
            enabled_sides = tuple(
                owned_side
                for owned_side in MODE_ROLES[self._mode]
                if self.hyperlink_check.isChecked()
                and self._preview_jump_available(owned_side)
            )
            comparison = ImageComparisonDialog(
                item,
                self._mode,
                self._workbook_paths,
                self,
                jump_callback=self._jump_from_preview,
                jump_enabled_sides=enabled_sides,
                role_order=self._preview_order,
                waveform_attributes=self.waveform_attributes,
            )
            comparison.exec_maximized()
        try:
            if not self._closing:
                self.table.setFocus(Qt.FocusReason.OtherFocusReason)
        except RuntimeError:
            # 확대창의 중첩 이벤트 루프에서 새 로드/메인 종료가 완료되면
            # 부모 목록창과 table이 이미 삭제되었을 수 있다.
            pass

    def focus_list(self) -> None:
        if not self._closing:
            self.table.setFocus(Qt.FocusReason.OtherFocusReason)

    def _stop_rendering(self) -> None:
        if self._closing:
            return
        self._closing = True
        self._close_memo_editor()
        self.table.setEnabled(False)
        try:
            self.table.currentCellChanged.disconnect(
                self._on_current_cell_changed
            )
        except (TypeError, RuntimeError):
            pass

    def prepare_for_close(self) -> None:
        self._stop_rendering()
        self.close()

    @classmethod
    def list_column_widths(cls, column_count: int, viewport_width: int) -> List[int]:
        """시트는 남는 폭의 절반, 나머지는 보이는 Ref/A/B/C 열에 균등 분배한다."""
        cell_count = int(column_count) - 3
        if cell_count <= 0:
            return []
        other_fixed = (
            cls._INDEX_COLUMN_WIDTH
            + cls._NO_COLUMN_WIDTH
            + cell_count * cls._CELL_COLUMN_WIDTH
        )
        leftover = max(0, int(viewport_width) - other_fixed)
        sheet_width = leftover // 2
        extra = leftover - sheet_width
        extra_each, extra_rem = divmod(extra, cell_count)
        widths = [
            cls._INDEX_COLUMN_WIDTH,
            sheet_width,
            cls._NO_COLUMN_WIDTH,
        ]
        for index in range(cell_count):
            widths.append(
                cls._CELL_COLUMN_WIDTH
                + extra_each
                + (1 if index < extra_rem else 0)
            )
        return widths

    def _apply_list_column_widths(self) -> None:
        if self._closing:
            return
        plan = self.list_column_widths(
            self.table.columnCount(),
            self.table.viewport().width(),
        )
        if not plan:
            return
        header = self.table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        for column, width in enumerate(plan):
            if self.table.columnWidth(column) != width:
                self.table.setColumnWidth(column, width)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._apply_list_column_widths()

    def showEvent(self, event) -> None:
        super().showEvent(event)
        QTimer.singleShot(0, self._apply_list_column_widths)
        if self._mode == "quadra" and not self._quadra_list_sizes_applied:
            QTimer.singleShot(0, self._finish_quadra_list_layout)
        self.focus_list()

    def _finish_quadra_list_layout(self) -> None:
        if self._closing or self._quadra_list_sizes_applied:
            return
        self._apply_quadra_list_sizes()
        self._quadra_list_sizes_applied = True

    def closeEvent(self, event) -> None:
        self._stop_rendering()
        super().closeEvent(event)
