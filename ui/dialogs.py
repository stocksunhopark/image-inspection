"""육안 검사 화면에서 사용하는 확대 창과 이미지 리스트 창."""

import os
from typing import Dict, List, Optional, Sequence, Tuple

from PIL import Image
from PyQt6.QtCore import QSignalBlocker, Qt, QTimer
from PyQt6.QtGui import QKeySequence, QShortcut
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from excel_manager import load_extracted_pil
from models import ExtractedImage, InspectionItem
from ui.helpers import pil_to_pixmap
from ui.widgets import ClickableLabel


class ImageViewerDialog(QDialog):
    """원본 비율을 유지하며 창맞춤·확대·축소를 제공한다."""

    def __init__(self, pil_image: Image.Image, title: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(1100, 800)
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
        toolbar.addStretch(1)
        layout.addLayout(toolbar)

        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(False)
        self._scroll.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._image_label = QLabel()
        self._image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._scroll.setWidget(self._image_label)
        layout.addWidget(self._scroll, stretch=1)

        QShortcut(QKeySequence("+"), self, activated=lambda: self._zoom_by(1.25))
        QShortcut(QKeySequence("-"), self, activated=lambda: self._zoom_by(0.8))
        QShortcut(QKeySequence("0"), self, activated=self._fit_to_window)
        self._apply_scale(1.0)
        QTimer.singleShot(0, self._fit_to_window)

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


class ImageListWindow(QDialog):
    """Excel 위치 목록과 해당 Double/Triple 이미지를 동시에 표시한다."""

    def __init__(
        self,
        items: Sequence[InspectionItem],
        mode: str,
        workbook_paths: Dict[str, str],
        *,
        initial_index: int = 0,
        parent=None,
    ):
        super().__init__(parent)
        if mode not in {"double", "triple"}:
            raise ValueError(f"지원하지 않는 검사 모드: {mode}")
        self._all_items = list(items)
        self._items = list(self._all_items)
        self._global_index_by_id = {
            id(item): index for index, item in enumerate(self._all_items)
        }
        self._mode = mode
        self._workbook_paths = dict(workbook_paths)
        self._current_index = -1
        self._closing = False
        mode_label = "Triple" if mode == "triple" else "Double"
        filenames = [
            os.path.basename(path)
            for path in self._workbook_paths.values()
            if path
        ]
        files_text = f" · {', '.join(filenames)}" if filenames else ""
        self.setWindowTitle(
            f"Excel 이미지 리스트 · {mode_label} · {len(self._all_items)}개{files_text}"
        )
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        self.resize(1550, 880)
        self.setMinimumSize(980, 620)

        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(7)

        header = QHBoxLayout()
        title = QLabel(f"Excel 이미지 리스트 · {mode_label}")
        title.setObjectName("positionLabel")
        header.addWidget(title)
        header.addSpacing(14)
        header.addWidget(QLabel("시트 선택"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.setMinimumWidth(230)
        self.sheet_combo.setToolTip("목록에 표시할 Excel 시트를 선택합니다.")
        header.addWidget(self.sheet_combo)
        header.addStretch(1)
        self.position_label = QLabel("0 / 0")
        self.position_label.setObjectName("positionLabel")
        header.addWidget(self.position_label)
        close_button = QPushButton("닫기")
        close_button.clicked.connect(self.close)
        header.addWidget(close_button)
        root.addLayout(header)

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
        self.content_splitter.addWidget(self.table)

        self.preview_widget = QWidget()
        preview_layout = QVBoxLayout(self.preview_widget)
        preview_layout.setContentsMargins(0, 4, 0, 0)
        preview_layout.setSpacing(5)
        preview_hint = QLabel(
            "목록에서 ↑/↓ 키 또는 마우스로 이동하면 이미지가 함께 바뀝니다."
        )
        preview_hint.setObjectName("keyboardHint")
        preview_layout.addWidget(preview_hint)

        self.preview_splitter = QSplitter(Qt.Orientation.Horizontal)
        (
            self.ref_container,
            self.ref_title,
            self.ref_meta,
            self.ref_image,
        ) = self._create_preview_pane("Excel Ref")
        (
            self.a_container,
            self.a_title,
            self.a_meta,
            self.a_image,
        ) = self._create_preview_pane("Excel 비교A")
        (
            self.b_container,
            self.b_title,
            self.b_meta,
            self.b_image,
        ) = self._create_preview_pane("Excel 비교B")
        for container in (self.ref_container, self.a_container, self.b_container):
            self.preview_splitter.addWidget(container)
        self.preview_splitter.setChildrenCollapsible(False)
        for index in range(3):
            self.preview_splitter.setStretchFactor(index, 1)
        self.preview_splitter.setHandleWidth(1)
        for handle_index in (1, 2):
            self.preview_splitter.handle(handle_index).setEnabled(False)
        self.b_container.setVisible(mode == "triple")
        self.ref_image.clicked.connect(lambda: self._enlarge("ref"))
        self.a_image.clicked.connect(lambda: self._enlarge("a"))
        self.b_image.clicked.connect(lambda: self._enlarge("b"))
        preview_layout.addWidget(self.preview_splitter, stretch=1)
        self.content_splitter.addWidget(self.preview_widget)
        self.content_splitter.setStretchFactor(0, 0)
        self.content_splitter.setStretchFactor(1, 1)
        self.content_splitter.setSizes([240, 600])
        root.addWidget(self.content_splitter, stretch=1)

        self._populate_sheet_combo()
        self.sheet_combo.currentIndexChanged.connect(
            self._on_sheet_filter_changed
        )
        self._populate_table()
        self.set_current_index(initial_index)

    def _create_preview_pane(
        self, default_title: str
    ) -> Tuple[QWidget, QLabel, QLabel, ClickableLabel]:
        container = QWidget()
        container.setMinimumWidth(0)
        container.setSizePolicy(
            QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Ignored
        )
        layout = QVBoxLayout(container)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(4)
        title = QLabel(default_title)
        title.setObjectName("previewTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setWordWrap(True)
        title.setFixedHeight(40)
        metadata = QLabel("-")
        metadata.setObjectName("imageMeta")
        metadata.setAlignment(Qt.AlignmentFlag.AlignCenter)
        metadata.setWordWrap(True)
        metadata.setFixedHeight(42)
        image = ClickableLabel("이미지 없음")
        image.setObjectName("previewPane")
        image.setCursor(Qt.CursorShape.PointingHandCursor)
        layout.addWidget(title)
        layout.addWidget(metadata)
        layout.addWidget(image, stretch=1)
        return container, title, metadata, image

    @staticmethod
    def _cell_text(extracted: Optional[ExtractedImage]) -> str:
        if extracted is None or extracted.is_null:
            return "이미지 없음"
        return extracted.cell_address

    def _populate_sheet_combo(self) -> None:
        counts: Dict[int, int] = {}
        names: Dict[int, str] = {}
        for item in self._all_items:
            counts[item.sheet_index] = counts.get(item.sheet_index, 0) + 1
            names.setdefault(item.sheet_index, item.sheet_name)
        blocker = QSignalBlocker(self.sheet_combo)
        self.sheet_combo.clear()
        self.sheet_combo.addItem(
            f"전체 시트 ({len(self._all_items)}개)", None
        )
        for sheet_index in sorted(counts):
            self.sheet_combo.addItem(
                f"{sheet_index}. {names[sheet_index]} ({counts[sheet_index]}개)",
                sheet_index,
            )
        del blocker

    def _on_sheet_filter_changed(self, combo_index: int) -> None:
        if self._closing or combo_index < 0:
            return
        current_item = (
            self._items[self._current_index]
            if 0 <= self._current_index < len(self._items)
            else None
        )
        sheet_index = self.sheet_combo.itemData(combo_index)
        self._items = (
            list(self._all_items)
            if sheet_index is None
            else [
                item
                for item in self._all_items
                if item.sheet_index == int(sheet_index)
            ]
        )
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
        self.focus_list()

    def _populate_table(self) -> None:
        headers = ["전체", "시트", "시트 내", "Ref 셀", "비교A 셀"]
        if self._mode == "triple":
            headers.append("비교B 셀")
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
                f"{item.sheet_index}. {item.sheet_name}",
                f"{sheet_positions[item.sheet_index]}/{sheet_totals[item.sheet_index]}",
                self._cell_text(item.image_ref),
                self._cell_text(item.image_a),
            ]
            if self._mode == "triple":
                values.append(self._cell_text(item.image_b))
            for column, value in enumerate(values):
                table_item = QTableWidgetItem(value)
                table_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                table_item.setData(Qt.ItemDataRole.UserRole, row)
                self.table.setItem(row, column, table_item)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(0, 58)
        self.table.setColumnWidth(1, 150)
        self.table.setColumnWidth(2, 72)
        for column in range(3, len(headers)):
            self.table.setColumnWidth(column, 92)
        header.setStretchLastSection(True)
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
            self._items = list(self._all_items)
            self._current_index = -1
            self._populate_table()
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
            self._current_index = current_row
            self._render_current()

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
            return

        item = self._items[self._current_index]
        global_index = self._global_index_by_id[id(item)]
        if self.sheet_combo.currentData() is None:
            position_text = f"전체 {global_index + 1} / {len(self._all_items)}"
        else:
            position_text = (
                f"시트 {self._current_index + 1} / {len(self._items)}"
                f" · 전체 {global_index + 1} / {len(self._all_items)}"
            )
        self.position_label.setText(position_text)
        self.location_label.setText(
            f"시트 {item.sheet_index}. {item.sheet_name} · 시트 내 {item.index} · 기준 위치 {item.cell_address}"
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
        if self._mode == "triple" and item.image_b is not None:
            self._set_preview_pane(
                self.b_title,
                self.b_meta,
                self.b_image,
                "Excel 비교B",
                self._workbook_paths.get("b", ""),
                item.image_b,
            )
        elif self._mode == "triple":
            self._clear_preview_pane(
                self.b_title, self.b_meta, self.b_image, "Excel 비교B"
            )
        self._equalize_preview_panes()

    def _equalize_preview_panes(self) -> None:
        sizes = (
            [1000, 1000, 1000]
            if self._mode == "triple"
            else [1000, 1000, 0]
        )
        self.preview_splitter.setSizes(sizes)

    @staticmethod
    def _clear_preview_pane(
        title_label: QLabel,
        meta_label: QLabel,
        image_label: ClickableLabel,
        title: str,
    ) -> None:
        title_label.setText(title)
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
        title_label.setToolTip(f"{role} · {filename}")
        meta_label.setText(extracted.location_text)
        meta_label.setToolTip(extracted.location_text)
        if extracted.is_null:
            image_label.setPixmap(None)
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
        }.get(side)
        if extracted is None or extracted.is_null:
            return
        image = load_extracted_pil(extracted)
        if image is None:
            return
        role = {"ref": "Ref", "a": "비교A", "b": "비교B"}[side]
        ImageViewerDialog(
            image,
            f"{role} · {extracted.location_text}",
            self,
        ).exec()
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

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self.focus_list()

    def closeEvent(self, event) -> None:
        self._stop_rendering()
        super().closeEvent(event)
