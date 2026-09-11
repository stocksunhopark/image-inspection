"""육안 검사 화면에서 사용하는 확대 창과 이미지 리스트 창."""

import os
from typing import Callable, Dict, List, Optional, Sequence, Tuple

from PIL import Image
from PyQt6.QtCore import QSettings, QSignalBlocker, Qt, QTimer, pyqtSignal
from PyQt6.QtGui import QKeySequence, QShortcut
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QHeaderView,
    QHBoxLayout,
    QLabel,
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
from models import SUPPORTED_MODES, ExtractedImage, InspectionItem, mode_label
from ui.helpers import pil_to_pixmap
from ui.widgets import ClickableLabel


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
        toolbar.addStretch(1)
        self.jump_button = QPushButton("엑셀 파형 바로가기")
        self.jump_button.setObjectName("inputActionBtn")
        self.jump_button.setToolTip(
            "하이퍼링크 모드가 켜져 있으면 이 Excel의 해당 시트·셀로 이동합니다."
        )
        self.jump_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.jump_button.setEnabled(bool(jump_enabled) and jump_callback is not None)
        self.jump_button.clicked.connect(self._on_jump_clicked)
        toolbar.addWidget(self.jump_button)
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


def create_preview_pane(
    default_title: str,
) -> Tuple[QWidget, QLabel, QLabel, ClickableLabel, QPushButton]:
    """셀 주소 옆에 엑셀 바로가기 버튼이 있는 미리보기 칸을 만든다."""
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
• Excel Ref: 기준 파일
• Excel 비교A / 비교B / 비교C: 나란히 볼 비교 파일
시트 수나 탭 순서가 달라도 사용할 수 있습니다. 프로그램은 시트 이름으로 연결합니다.
메뉴 [파일] → [Ref Excel 선택...] (Ctrl+O)으로 Ref 파일만 고를 수도 있습니다.


3. 이미지 불러오기
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
Queue 생성 뒤 Sheet 역할 관계와 Source Image ID를 검사하며, 누락·중복·역할 오배치가 있으면 Review를 시작하지 않습니다.


4. 메인 화면에서 넘기기
불러온 뒤 같은 위치의 이미지가 한 화면에 나란히 보입니다.
• [처음] [이전] [다음] [마지막] 버튼
• 왼쪽 방향키, PageUp: 이전
• 오른쪽 방향키, PageDown, Space: 다음
• Home: 처음 / End: 마지막
• 상단 [시트] 목록: 시트별 [REF + A], [REF ONLY], [C ONLY] 상태를 보고 이동
• 슬라이더: 전체 위치 중 원하는 곳으로 바로 이동
미리보기의 시트명·셀주소는 가운데에, [엑셀 파형 바로가기]는 오른쪽 끝에 있습니다.
한쪽에만 이미지가 없어도 위치는 건너뛰지 않습니다.


5. 리스트 창
[리스트실행]을 누르면 목록과 미리보기가 함께 있는 창이 열립니다.
• 위: 전체 / 시트 / No. / Ref 셀 / 비교A 셀 (Triple이면 비교B 셀, Quadra이면 비교C 셀까지)
• 아래: 선택한 행의 이미지. Double·Triple은 가로로, Quadra는 Ref|A / B|C 2×2
• 미리보기: 시트명·셀주소는 가운데, [엑셀 파형 바로가기]는 오른쪽 끝
• 한 번 클릭 또는 ↑/↓: 미리보기만 바뀝니다. Excel은 열리지 않습니다.
• 상단 [시트 선택]: 전체 시트 또는 특정 시트만 목록에 표시
제목 표시줄을 더블클릭하면 최대화할 수 있습니다.


6. 하이퍼링크 모드
메인 화면의 [시트] 옆과 리스트 창의 [시트 선택] 옆에 같은 [하이퍼링크 모드] 체크박스가 있습니다. 기본값은 꺼짐입니다.
한쪽에서 켜거나 끄면 다른 쪽에도 같이 적용되고, 다음에 프로그램을 열 때도 유지됩니다.
• 꺼짐: 더블클릭이나 바로가기 버튼을 눌러도 Excel을 열지 않습니다. 바로가기 버튼은 비활성입니다.
• 켜짐: 아래 방법으로 해당 Excel을 열고 그 시트·셀로 이동합니다.
  - 리스트의 Ref 셀 / 비교A 셀 / 비교B 셀 / 비교C 셀 더블클릭
  - 메인·리스트 미리보기 오른쪽 끝의 [엑셀 파형 바로가기]
  - 확대 팝업 오른쪽 위의 [엑셀 파형 바로가기]
  - 전체, 시트, No. 칸을 더블클릭해도 Excel은 열리지 않습니다.
  - '시트 없음'과 '이미지 없음'은 열지 않습니다.
  - 이미 Excel이 열려 있으면 새 창을 또 열지 않고, 그 창에서 해당 파일·칸으로 이동합니다.
  - 두 번째 이후에도 Excel 창이 맨 앞으로 올라옵니다.
목록을 넘기는 속도는 하이퍼링크를 켜도 거의 같습니다. Excel은 더블클릭하거나 바로가기 버튼을 누른 순간에만 엽니다.


7. 이미지 확대
메인 화면이나 리스트 창에서 미리보기 이미지를 클릭하면 확대 창이 열립니다.
• [확대 +] / [축소 −] 또는 + / − 키
• [창에 맞춤] 또는 0 키
• [실제 크기]
• 오른쪽 위 [엑셀 파형 바로가기]: 하이퍼링크 모드가 켜져 있으면 그 이미지의 Excel 칸으로 이동합니다.
창 크기를 바꾸면 맞춤 모드일 때 이미지도 다시 맞춰집니다.


8. 참고
• 이 프로그램은 유사도 점수, PASS/FAIL, 리포트를 만들지 않습니다.
• 정상 로딩 결과는 원본 이미지 누락 0, 중복 0을 내부 검증한 결과입니다.
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


class ImageListWindow(QDialog):
    """Excel 위치 목록과 해당 Double/Triple 이미지를 동시에 표시한다."""

    excel_jump_failed = pyqtSignal(str)
    hyperlink_mode_changed = pyqtSignal(bool)
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
        self._workbook_paths = dict(workbook_paths)
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
            "목록에서 ↑/↓ 키 또는 마우스로 이동하면 이미지가 함께 바뀝니다."
        )
        preview_hint.setObjectName("keyboardHint")
        preview_layout.addWidget(preview_hint)

        (
            self.ref_container,
            self.ref_title,
            self.ref_meta,
            self.ref_image,
            self.ref_jump_button,
        ) = self._create_preview_pane("Excel Ref")
        (
            self.a_container,
            self.a_title,
            self.a_meta,
            self.a_image,
            self.a_jump_button,
        ) = self._create_preview_pane("Excel 비교A")
        (
            self.b_container,
            self.b_title,
            self.b_meta,
            self.b_image,
            self.b_jump_button,
        ) = self._create_preview_pane("Excel 비교B")
        (
            self.c_container,
            self.c_title,
            self.c_meta,
            self.c_image,
            self.c_jump_button,
        ) = self._create_preview_pane("Excel 비교C")
        self.ref_image.clicked.connect(lambda: self._enlarge("ref"))
        self.a_image.clicked.connect(lambda: self._enlarge("a"))
        self.b_image.clicked.connect(lambda: self._enlarge("b"))
        self.c_image.clicked.connect(lambda: self._enlarge("c"))
        self.ref_jump_button.clicked.connect(lambda: self._jump_from_preview("ref"))
        self.a_jump_button.clicked.connect(lambda: self._jump_from_preview("a"))
        self.b_jump_button.clicked.connect(lambda: self._jump_from_preview("b"))
        self.c_jump_button.clicked.connect(lambda: self._jump_from_preview("c"))
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
        self._populate_table()
        self.set_current_index(initial_index)

    def _create_preview_pane(
        self, default_title: str
    ) -> Tuple[QWidget, QLabel, QLabel, ClickableLabel, QPushButton]:
        return create_preview_pane(default_title)

    @staticmethod
    def _cell_text(extracted: Optional[ExtractedImage]) -> str:
        if extracted is None:
            return "이미지 없음"
        if extracted.is_null:
            return extracted.placeholder_text
        return extracted.cell_address

    def _populate_sheet_combo(self) -> None:
        counts: Dict[int, int] = {}
        names: Dict[int, str] = {}
        statuses: Dict[int, str] = {}
        empty_sheets = set()
        for item in self._all_items:
            counts[item.sheet_index] = counts.get(item.sheet_index, 0) + 1
            names.setdefault(item.sheet_index, item.sheet_name)
            statuses.setdefault(item.sheet_index, item.sheet_status_text)
            if item.is_empty_sheet:
                empty_sheets.add(item.sheet_index)
        blocker = QSignalBlocker(self.sheet_combo)
        self.sheet_combo.clear()
        self.sheet_combo.addItem(
            f"전체 시트 ({len(counts)}개) · Review {len(self._all_items)}개", None
        )
        for sheet_index in sorted(counts):
            review_count = 0 if sheet_index in empty_sheets else counts[sheet_index]
            self.sheet_combo.addItem(
                f"{sheet_index}. {names[sheet_index]} "
                f"[{statuses[sheet_index]}] ({review_count}개)",
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
            f"시트 {item.sheet_index}. {item.sheet_name} [{item.sheet_status_text}] · "
            f"시트 내 {item.index} · 기준 위치 {item.cell_address}"
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
        ImageViewerDialog(
            image,
            f"{role} · {extracted.location_text}",
            self,
            jump_enabled=self.hyperlink_check.isChecked()
            and self._preview_jump_available(side),
            jump_callback=lambda: self._jump_from_preview(side),
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
