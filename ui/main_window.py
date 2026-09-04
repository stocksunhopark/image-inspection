"""자동 판정 없이 Double/Triple 이미지를 넘겨 보는 메인 창."""

import os
import shutil
from typing import Dict, Optional, Tuple

from PyQt6.QtCore import QSettings, QSignalBlocker, QThread, QTimer, Qt
from PyQt6.QtGui import QAction, QKeySequence
from PyQt6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QSizePolicy,
    QSlider,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from app_state import AppState
from excel_manager import load_extracted_pil
from models import ExtractedImage
from ui.dialogs import ImageListWindow, ImageViewerDialog
from ui.helpers import pil_to_pixmap
from ui.theme import apply_theme
from ui.widgets import ClickableLabel, DropLineEdit
from workers import ExcelLoadWorker


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel 이미지 육안 검사기")
        self.setMinimumSize(900, 650)
        self.resize(1500, 920)
        self.state = AppState()
        self.settings = QSettings("excel-image-inspector", "gui")
        self.load_thread: Optional[QThread] = None
        self.load_worker: Optional[ExcelLoadWorker] = None
        self.image_list_window: Optional[ImageListWindow] = None
        self._close_pending = False
        self._owned_temp_dirs = set()
        self._build_ui()
        self._build_menu()
        self._build_navigation_shortcuts()
        self._restore_settings()
        self._render_current()

    def _build_ui(self) -> None:
        apply_theme(self)
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(8)

        title = QLabel("Excel 이미지 육안 검사기")
        title.setObjectName("titleLabel")
        subtitle = QLabel(
            "자동 비교나 PASS/FAIL 판정 없이, 같은 위치의 이미지를 나란히 넘겨 봅니다."
        )
        subtitle.setObjectName("subTitleLabel")
        layout.addWidget(title)
        layout.addWidget(subtitle)

        input_group = QGroupBox("Excel 입력")
        input_layout = QGridLayout(input_group)
        input_layout.setContentsMargins(10, 14, 10, 10)
        input_layout.setHorizontalSpacing(8)
        input_layout.setVerticalSpacing(7)

        self.double_radio = QRadioButton("Double (Excel 2개)")
        self.triple_radio = QRadioButton("Triple (Excel 3개)")
        self.double_radio.setChecked(True)
        mode_group = QButtonGroup(self)
        mode_group.addButton(self.double_radio)
        mode_group.addButton(self.triple_radio)
        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel("검사 모드"))
        mode_row.addWidget(self.double_radio)
        mode_row.addWidget(self.triple_radio)
        mode_row.addStretch(1)
        input_layout.addLayout(mode_row, 0, 0, 1, 3)

        self.file_ref_edit, self.file_ref_button = self._add_file_row(
            input_layout, 1, "Excel Ref", "기준 Excel 파일"
        )
        self.file_a_edit, self.file_a_button = self._add_file_row(
            input_layout, 2, "Excel 비교A", "첫 번째 비교 Excel 파일"
        )
        self.b_row = QWidget()
        b_layout = QGridLayout(self.b_row)
        b_layout.setContentsMargins(0, 0, 0, 0)
        self.file_b_edit = DropLineEdit()
        self.file_b_edit.setPlaceholderText(".xlsx 또는 .xlsm 경로 입력/끌어놓기")
        self.file_b_button = QPushButton("파일 선택")
        self.file_b_button.clicked.connect(
            lambda: self._pick_file(self.file_b_edit)
        )
        b_layout.addWidget(QLabel("Excel 비교B"), 0, 0)
        b_layout.addWidget(self.file_b_edit, 0, 1)
        b_layout.addWidget(self.file_b_button, 0, 2)
        b_layout.setColumnStretch(1, 1)
        input_layout.addWidget(self.b_row, 3, 0, 1, 3)

        action_row = QHBoxLayout()
        self.load_button = QPushButton("이미지 불러오기")
        self.cancel_button = QPushButton("취소")
        self.cancel_button.setEnabled(False)
        self.load_button.clicked.connect(self._start_loading)
        self.cancel_button.clicked.connect(self._cancel_loading)
        action_row.addWidget(self.load_button)
        action_row.addWidget(self.cancel_button)
        action_row.addStretch(1)
        self.status_label = QLabel("Excel 파일을 선택해 주세요.")
        self.status_label.setObjectName("statusIdle")
        action_row.addWidget(self.status_label)
        input_layout.addLayout(action_row, 4, 0, 1, 3)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        input_layout.addWidget(self.progress_bar, 5, 0, 1, 3)
        input_layout.setColumnStretch(1, 1)
        layout.addWidget(input_group)

        navigation_group = QGroupBox("검사 이미지 이동")
        navigation = QVBoxLayout(navigation_group)
        top_navigation = QHBoxLayout()
        top_navigation.addWidget(QLabel("시트"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.setMinimumWidth(210)
        self.sheet_combo.currentIndexChanged.connect(self._select_sheet)
        top_navigation.addWidget(self.sheet_combo)
        self.first_button = QPushButton("|◀ 처음")
        self.previous_button = QPushButton("◀ 이전")
        self.next_button = QPushButton("다음 ▶")
        self.last_button = QPushButton("마지막 ▶|")
        self.list_button = QPushButton("리스트실행")
        self.list_button.setToolTip("전체 이미지 위치 목록과 이미지를 함께 봅니다.")
        self.first_button.clicked.connect(self._go_first)
        self.previous_button.clicked.connect(lambda: self._move(-1))
        self.next_button.clicked.connect(lambda: self._move(1))
        self.last_button.clicked.connect(self._go_last)
        self.list_button.clicked.connect(self._open_image_list_window)
        for button in (
            self.first_button,
            self.previous_button,
            self.next_button,
            self.last_button,
        ):
            top_navigation.addWidget(button)
        top_navigation.addWidget(self.list_button)
        top_navigation.addStretch(1)
        self.position_label = QLabel("0 / 0")
        self.position_label.setObjectName("positionLabel")
        top_navigation.addWidget(self.position_label)
        self.current_mode_label = QLabel("현재 결과: 없음")
        self.current_mode_label.setObjectName("locationLabel")
        top_navigation.addWidget(self.current_mode_label)
        navigation.addLayout(top_navigation)

        self.position_slider = QSlider(Qt.Orientation.Horizontal)
        self.position_slider.setRange(0, 0)
        self.position_slider.valueChanged.connect(self._jump_to_position)
        navigation.addWidget(self.position_slider)
        self.location_label = QLabel("불러온 이미지가 없습니다.")
        self.location_label.setObjectName("locationLabel")
        navigation.addWidget(self.location_label)
        layout.addWidget(navigation_group)

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
        self.preview_splitter.setStretchFactor(0, 1)
        self.preview_splitter.setStretchFactor(1, 1)
        self.preview_splitter.setStretchFactor(2, 1)
        self.preview_splitter.setHandleWidth(1)
        for handle_index in (1, 2):
            self.preview_splitter.handle(handle_index).setEnabled(False)
        self.ref_image.clicked.connect(lambda: self._enlarge("ref"))
        self.a_image.clicked.connect(lambda: self._enlarge("a"))
        self.b_image.clicked.connect(lambda: self._enlarge("b"))
        layout.addWidget(self.preview_splitter, stretch=1)

        keyboard_hint = QLabel(
            "←/PageUp 이전 · →/PageDown 다음 · Home/End 처음/마지막 · 이미지 클릭 시 확대"
        )
        keyboard_hint.setObjectName("keyboardHint")
        keyboard_hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(keyboard_hint)

        self.double_radio.toggled.connect(self._on_input_mode_changed)
        self.triple_radio.toggled.connect(self._on_input_mode_changed)
        self._on_input_mode_changed()

    def _add_file_row(
        self,
        layout: QGridLayout,
        row: int,
        label: str,
        tooltip: str,
    ) -> Tuple[DropLineEdit, QPushButton]:
        edit = DropLineEdit()
        edit.setPlaceholderText(".xlsx 또는 .xlsm 경로 입력/끌어놓기")
        edit.setToolTip(tooltip)
        button = QPushButton("파일 선택")
        button.clicked.connect(lambda: self._pick_file(edit))
        layout.addWidget(QLabel(label), row, 0)
        layout.addWidget(edit, row, 1)
        layout.addWidget(button, row, 2)
        return edit, button

    def _create_preview_pane(
        self, default_title: str
    ) -> Tuple[QWidget, QLabel, QLabel, ClickableLabel]:
        container = QWidget()
        container.setMinimumWidth(0)
        container.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Ignored)
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

    def _build_menu(self) -> None:
        file_menu = self.menuBar().addMenu("파일(&F)")
        self.select_ref_action = QAction("Ref Excel 선택...", self)
        self.select_ref_action.setShortcut(QKeySequence.StandardKey.Open)
        self.select_ref_action.triggered.connect(
            lambda: self._pick_file(self.file_ref_edit)
        )
        self.load_action = QAction("이미지 불러오기", self)
        self.load_action.setShortcut("Ctrl+Return")
        self.load_action.triggered.connect(self._start_loading)
        exit_action = QAction("종료", self)
        exit_action.setShortcut(QKeySequence.StandardKey.Quit)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(self.select_ref_action)
        file_menu.addAction(self.load_action)
        file_menu.addSeparator()
        file_menu.addAction(exit_action)

        help_menu = self.menuBar().addMenu("도움말(&H)")
        about_action = QAction("사용 방법", self)
        about_action.triggered.connect(self._show_help)
        help_menu.addAction(about_action)

    def _build_navigation_shortcuts(self) -> None:
        shortcuts = (
            (["Left", "PageUp"], lambda: self._move(-1)),
            (["Right", "PageDown", "Space"], lambda: self._move(1)),
            (["Home"], self._go_first),
            (["End"], self._go_last),
        )
        self._navigation_actions = []
        for keys, callback in shortcuts:
            action = QAction(self)
            action.setShortcuts([QKeySequence(key) for key in keys])
            action.setShortcutContext(Qt.ShortcutContext.WindowShortcut)
            action.triggered.connect(callback)
            self.addAction(action)
            self._navigation_actions.append(action)

    def _selected_mode(self) -> str:
        return "triple" if self.triple_radio.isChecked() else "double"

    def _on_input_mode_changed(self) -> None:
        is_triple = self._selected_mode() == "triple"
        self.b_row.setVisible(is_triple)
        if not self.state.flat_items:
            self.b_container.setVisible(is_triple)
            self._equalize_preview_panes(is_triple)
        elif self._selected_mode() != self.state.mode:
            selected = "Triple" if is_triple else "Double"
            loaded = "Triple" if self.state.mode == "triple" else "Double"
            self._set_status(
                f"입력 모드는 {selected}로 변경됨 · 현재 화면은 이전 {loaded} 결과입니다.",
                busy=False,
            )

    def _pick_file(self, target: DropLineEdit) -> None:
        if self.load_thread is not None:
            return
        start = target.text().strip()
        if not start or not os.path.exists(start):
            start = self.settings.value("last_directory", "", type=str)
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Excel 파일 선택",
            start,
            "Excel 통합 문서 (*.xlsx *.xlsm)",
        )
        if path:
            target.setText(path)
            self.settings.setValue("last_directory", os.path.dirname(path))

    def _start_loading(self) -> None:
        if self.load_thread is not None:
            return
        mode = self._selected_mode()
        path_ref = self.file_ref_edit.text().strip()
        path_a = self.file_a_edit.text().strip()
        path_b = self.file_b_edit.text().strip() if mode == "triple" else None
        missing = []
        if not path_ref:
            missing.append("Excel Ref")
        if not path_a:
            missing.append("Excel 비교A")
        if mode == "triple" and not path_b:
            missing.append("Excel 비교B")
        if missing:
            QMessageBox.warning(
                self,
                "입력 확인",
                "다음 파일을 선택해 주세요: " + ", ".join(missing),
            )
            return

        self._set_loading(True)
        self.progress_bar.setValue(0)
        self._set_status("이미지 로딩 준비 중...", busy=True)
        worker = ExcelLoadWorker(path_ref, path_a, path_b)
        thread = QThread(self)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.progress.connect(self._on_progress)
        worker.status.connect(lambda message: self._set_status(message, busy=True))
        worker.finished.connect(self._on_load_finished)
        worker.failed.connect(self._on_load_failed)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        worker.failed.connect(worker.deleteLater)
        thread.finished.connect(self._on_load_thread_finished)
        thread.finished.connect(thread.deleteLater)
        self.load_worker = worker
        self.load_thread = thread
        thread.start()

    def _cancel_loading(self) -> None:
        if self.load_worker is None:
            return
        self.load_worker.request_stop()
        self.cancel_button.setEnabled(False)
        self._set_status("취소 요청됨...", busy=True)

    def _on_progress(self, current: int, total: int, cell: str) -> None:
        percent = int(current * 100 / total) if total else 0
        self.progress_bar.setValue(percent)
        self._set_status(
            f"이미지 추출 중... {current}/{total} · 셀 {cell}", busy=True
        )

    def _on_load_finished(self, payload) -> None:
        items_by_sheet, preview_dir, mode, paths = payload
        self._close_image_list_window()
        self._owned_temp_dirs.add(os.path.realpath(preview_dir))
        old_temp_dir = self.state.preview_temp_dir
        self.state.set_items(
            items_by_sheet,
            mode=mode,
            workbook_paths=paths,
            preview_temp_dir=preview_dir,
        )
        self._cleanup_temp_dir(old_temp_dir)
        self._populate_sheet_combo()
        self.b_container.setVisible(mode == "triple")
        self.progress_bar.setValue(100)
        total = len(self.state.flat_items)
        if total:
            self._set_status(f"완료: {total}개 위치를 불러왔습니다.", busy=False)
        else:
            self._set_status("삽입 이미지가 없습니다.", busy=False)
        self._render_current()
        self._save_settings()

    def _on_load_failed(self, message: str) -> None:
        previous = self._previous_result_text()
        if message == "사용자 취소":
            self._set_status(
                "이미지 불러오기가 취소되었습니다." + previous, busy=False
            )
        else:
            self._set_status("이미지 불러오기 실패" + previous, busy=False)
            QMessageBox.critical(self, "불러오기 오류", message)

    def _previous_result_text(self) -> str:
        if not self.state.flat_items:
            return ""
        mode = "Triple" if self.state.mode == "triple" else "Double"
        names = [
            os.path.basename(path)
            for path in self.state.workbook_paths.values()
            if path
        ]
        suffix = f" ({', '.join(names)})" if names else ""
        return f" · 아래 화면은 이전 {mode} 결과입니다{suffix}."

    def _on_load_thread_finished(self) -> None:
        self.load_worker = None
        self.load_thread = None
        self._set_loading(False)
        if self._close_pending:
            QTimer.singleShot(0, self.close)

    def _set_loading(self, loading: bool) -> None:
        for widget in (
            self.double_radio,
            self.triple_radio,
            self.file_ref_edit,
            self.file_a_edit,
            self.file_b_edit,
            self.file_ref_button,
            self.file_a_button,
            self.file_b_button,
            self.load_button,
        ):
            widget.setEnabled(not loading)
        self.cancel_button.setEnabled(loading)
        self.select_ref_action.setEnabled(not loading)
        self.load_action.setEnabled(not loading)

    def _set_status(self, text: str, *, busy: bool) -> None:
        self.status_label.setObjectName("statusBusy" if busy else "statusIdle")
        self.status_label.setText(text)
        self.status_label.style().unpolish(self.status_label)
        self.status_label.style().polish(self.status_label)

    def _populate_sheet_combo(self) -> None:
        with QSignalBlocker(self.sheet_combo):
            self.sheet_combo.clear()
            for sheet_index in sorted(self.state.items_by_sheet):
                rows = self.state.items_by_sheet[sheet_index]
                if not rows:
                    continue
                self.sheet_combo.addItem(
                    f"{sheet_index}. {rows[0].sheet_name} ({len(rows)}개)",
                    sheet_index,
                )

    def _select_sheet(self, combo_index: int) -> None:
        if combo_index < 0:
            return
        sheet_index = self.sheet_combo.itemData(combo_index)
        if sheet_index is not None:
            self.state.select_sheet(int(sheet_index))
            self._render_current()

    def _move(self, offset: int) -> None:
        if self.state.move(offset):
            self._render_current()

    def _go_first(self) -> None:
        if self.state.go_first():
            self._render_current()

    def _go_last(self) -> None:
        if self.state.go_last():
            self._render_current()

    def _jump_to_position(self, value: int) -> None:
        if not self.state.flat_items:
            return
        target = max(0, min(len(self.state.flat_items) - 1, int(value)))
        if target != self.state.current_index:
            self.state.current_index = target
            self._render_current()

    def _render_current(self) -> None:
        item = self.state.current_item
        total = len(self.state.flat_items)
        has_items = item is not None
        if not has_items:
            self.position_label.setText("0 / 0")
            self.current_mode_label.setText("현재 결과: 없음")
            self.location_label.setText("불러온 이미지가 없습니다.")
            with QSignalBlocker(self.position_slider):
                self.position_slider.setRange(0, 0)
                self.position_slider.setValue(0)
            self._clear_preview_pane(
                self.ref_title, self.ref_meta, self.ref_image, "Excel Ref"
            )
            self._clear_preview_pane(
                self.a_title, self.a_meta, self.a_image, "Excel 비교A"
            )
            self._clear_preview_pane(
                self.b_title, self.b_meta, self.b_image, "Excel 비교B"
            )
            self._update_navigation_buttons()
            return

        current = self.state.current_index
        loaded_mode = "Triple" if self.state.mode == "triple" else "Double"
        self.current_mode_label.setText(f"현재 결과: {loaded_mode}")
        sheet_position, sheet_total = self.state.current_sheet_position()
        self.position_label.setText(f"전체 {current + 1} / {total}")
        self.location_label.setText(
            f"시트 {item.sheet_index} · 시트 내 {sheet_position}/{sheet_total} · 기준 위치 {item.cell_address}"
        )
        with QSignalBlocker(self.position_slider):
            self.position_slider.setRange(0, max(0, total - 1))
            self.position_slider.setValue(current)
        combo_index = self.sheet_combo.findData(item.sheet_index)
        if combo_index >= 0:
            with QSignalBlocker(self.sheet_combo):
                self.sheet_combo.setCurrentIndex(combo_index)

        self._set_preview_pane(
            self.ref_title,
            self.ref_meta,
            self.ref_image,
            "Excel Ref",
            self.state.workbook_paths.get("ref", ""),
            item.image_ref,
        )
        self._set_preview_pane(
            self.a_title,
            self.a_meta,
            self.a_image,
            "Excel 비교A",
            self.state.workbook_paths.get("a", ""),
            item.image_a,
        )
        if self.state.mode == "triple" and item.image_b is not None:
            self.b_container.setVisible(True)
            self._set_preview_pane(
                self.b_title,
                self.b_meta,
                self.b_image,
                "Excel 비교B",
                self.state.workbook_paths.get("b", ""),
                item.image_b,
            )
        else:
            self.b_container.setVisible(False)
        self._equalize_preview_panes(self.state.mode == "triple")
        self._update_navigation_buttons()

    def _equalize_preview_panes(self, is_triple: bool) -> None:
        sizes = [1000, 1000, 1000] if is_triple else [1000, 1000, 0]
        self.preview_splitter.setSizes(sizes)

    def _clear_preview_pane(
        self,
        title_label: QLabel,
        meta_label: QLabel,
        image_label: ClickableLabel,
        title: str,
    ) -> None:
        title_label.setText(title)
        meta_label.setText("-")
        image_label.setPixmap(None)
        image_label.setText("이미지 없음")

    def _set_preview_pane(
        self,
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

    def _update_navigation_buttons(self) -> None:
        total = len(self.state.flat_items)
        current = self.state.current_index
        has_previous = total > 0 and current > 0
        has_next = total > 0 and current < total - 1
        self.first_button.setEnabled(has_previous)
        self.previous_button.setEnabled(has_previous)
        self.next_button.setEnabled(has_next)
        self.last_button.setEnabled(has_next)
        self.list_button.setEnabled(total > 0)
        self.sheet_combo.setEnabled(total > 0)
        self.position_slider.setEnabled(total > 0)

    def _open_image_list_window(self) -> None:
        if not self.state.flat_items:
            return
        if self.image_list_window is not None:
            try:
                self.image_list_window.set_current_index(self.state.current_index)
                if self.image_list_window.isMinimized():
                    self.image_list_window.showNormal()
                else:
                    self.image_list_window.show()
                self.image_list_window.raise_()
                self.image_list_window.activateWindow()
                self.image_list_window.focus_list()
                return
            except RuntimeError:
                self.image_list_window = None

        window = ImageListWindow(
            self.state.flat_items,
            self.state.mode,
            self.state.workbook_paths,
            initial_index=self.state.current_index,
            parent=self,
        )
        self.image_list_window = window
        window.destroyed.connect(
            lambda _object=None, owned=window: self._on_image_list_destroyed(owned)
        )
        window.finished.connect(
            lambda _result=0, owned=window: self._on_image_list_destroyed(owned)
        )
        window.show()
        window.raise_()
        window.activateWindow()

    def _on_image_list_destroyed(self, window: ImageListWindow) -> None:
        if self.image_list_window is window:
            self.image_list_window = None

    def _close_image_list_window(self) -> None:
        if self.image_list_window is None:
            return
        window = self.image_list_window
        self.image_list_window = None
        try:
            window.prepare_for_close()
        except RuntimeError:
            pass

    def _enlarge(self, side: str) -> None:
        item = self.state.current_item
        if item is None:
            return
        extracted = {
            "ref": item.image_ref,
            "a": item.image_a,
            "b": item.image_b,
        }.get(side)
        if extracted is None or extracted.is_null:
            return
        image = load_extracted_pil(extracted)
        if image is None:
            QMessageBox.warning(self, "이미지 오류", "이미지를 불러올 수 없습니다.")
            return
        role = {"ref": "Ref", "a": "비교A", "b": "비교B"}[side]
        ImageViewerDialog(
            image,
            f"{role} · {extracted.location_text}",
            self,
        ).exec()

    def _show_help(self) -> None:
        QMessageBox.information(
            self,
            "사용 방법",
            "1. Double 또는 Triple 모드를 선택합니다.\n"
            "2. Excel 2개 또는 3개를 지정하고 이미지 불러오기를 누릅니다.\n"
            "3. 이전/다음 버튼이나 방향키로 같은 위치의 이미지를 넘겨 봅니다.\n"
            "4. 리스트 실행을 누르면 목록과 이미지를 함께 보고, 상단에서 시트를 선택할 수 있습니다.\n"
            "5. 이미지를 클릭하면 확대 창에서 원본을 확인할 수 있습니다.\n\n"
            "자동 점수 계산과 PASS/FAIL 판정은 수행하지 않습니다.",
        )

    def _restore_settings(self) -> None:
        mode = self.settings.value("mode", "double", type=str)
        self.triple_radio.setChecked(mode == "triple")
        self.double_radio.setChecked(mode != "triple")
        self.file_ref_edit.setText(self.settings.value("path_ref", "", type=str))
        self.file_a_edit.setText(self.settings.value("path_a", "", type=str))
        self.file_b_edit.setText(self.settings.value("path_b", "", type=str))
        geometry = self.settings.value("geometry")
        if geometry is not None:
            self.restoreGeometry(geometry)
        self._on_input_mode_changed()

    def _save_settings(self) -> None:
        self.settings.setValue("mode", self._selected_mode())
        self.settings.setValue("path_ref", self.file_ref_edit.text().strip())
        self.settings.setValue("path_a", self.file_a_edit.text().strip())
        self.settings.setValue("path_b", self.file_b_edit.text().strip())
        self.settings.setValue("geometry", self.saveGeometry())

    def _cleanup_temp_dir(self, path: Optional[str]) -> None:
        if not path or not os.path.isdir(path):
            return
        resolved = os.path.realpath(path)
        if resolved not in self._owned_temp_dirs:
            return
        shutil.rmtree(resolved, ignore_errors=True)
        self._owned_temp_dirs.discard(resolved)

    def closeEvent(self, event) -> None:
        if self.load_thread is not None and self.load_thread.isRunning():
            self._close_pending = True
            self._cancel_loading()
            event.ignore()
            return
        self._close_image_list_window()
        self._save_settings()
        self._cleanup_temp_dir(self.state.preview_temp_dir)
        self.state.preview_temp_dir = None
        event.accept()
