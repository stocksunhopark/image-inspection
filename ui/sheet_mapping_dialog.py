"""이미지 로딩 전에 시트 순서와 비교 그룹을 편집하는 대화상자."""

from __future__ import annotations

from typing import Dict, List, Optional

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QBrush, QPalette
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from models import ROLE_LABELS
from sheet_mapping import (
    SheetMappingGroup,
    SheetMappingPlan,
    SheetMappingSource,
    validate_sheet_mapping_plan,
)
from ui.theme import apply_theme


class _TableScrollComboBox(QComboBox):
    """표 스크롤 휠이 실수로 시트 선택을 바꾸지 않게 한다."""

    def __init__(self, scroll_target: QTableWidget):
        super().__init__()
        self._scroll_target = scroll_target

    def wheelEvent(self, event) -> None:
        # 셀 위젯에서 무시된 휠 이벤트가 항상 표까지 전파되지는 않는다.
        # 같은 이벤트를 표의 스크롤 처리기로 직접 넘겨 선택 변경을 막는다.
        self._scroll_target.wheelEvent(event)


class SheetMappingDialog(QDialog):
    """자동 매핑 전체를 행으로 보여 주고 사용자 지정 연결을 받는다."""

    def __init__(
        self,
        current_plan: SheetMappingPlan,
        automatic_plan: SheetMappingPlan,
        parent=None,
    ):
        super().__init__(parent)
        apply_theme(self)
        self.setWindowTitle("시트 순서 설정")
        self.setWindowFlag(Qt.WindowType.WindowMaximizeButtonHint, True)
        self.setWindowFlag(Qt.WindowType.WindowMinimizeButtonHint, True)
        self.setSizeGripEnabled(True)
        self.setMinimumSize(900, 560)
        self._automatic_plan = automatic_plan.clone()
        self._plan = current_plan.clone()
        self._custom_groups: List[SheetMappingGroup] = []
        self.result_plan: Optional[SheetMappingPlan] = None
        self._updating = False
        self._source_by_id: Dict[str, SheetMappingSource] = {
            source.source_sheet_id: source for source in self._plan.sources
        }
        self._sources_by_role: Dict[str, List[SheetMappingSource]] = {
            role: [
                source for source in self._plan.sources if source.role == role
            ]
            for role in self._plan.roles
        }
        self._build_ui()
        self._render_table()
        self._apply_initial_window_size()

    @property
    def plan(self) -> SheetMappingPlan:
        return self._plan.clone()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        title = QLabel("현재 시트 매핑과 실행 순서")
        title.setObjectName("titleLabel")
        layout.addWidget(title)
        guide = QLabel(
            "같은 이름으로 자동 연결된 시트와 ONLY 시트를 모두 표시합니다. "
            "새 비교 그룹에서 시트를 선택하면 적용 시 해당 그룹으로 이동합니다."
        )
        guide.setWordWrap(True)
        guide.setObjectName("subTitleLabel")
        layout.addWidget(guide)

        toolbar = QHBoxLayout()
        self.restore_button = QPushButton("자동 매핑 복원")
        self.add_group_button = QPushButton("비교 그룹 추가")
        self.delete_group_button = QPushButton("추가 그룹 삭제")
        self.enable_all_button = QPushButton("전체 활성")
        self.disable_all_button = QPushButton("전체 비활성")
        self.invert_enabled_button = QPushButton("선택 반전")
        self.move_up_button = QPushButton("위로")
        self.move_down_button = QPushButton("아래로")
        self.restore_button.clicked.connect(self._restore_automatic_mapping)
        self.add_group_button.clicked.connect(self._add_group)
        self.delete_group_button.clicked.connect(self._delete_selected_custom_group)
        self.enable_all_button.clicked.connect(
            lambda: self._set_all_groups_enabled(True)
        )
        self.disable_all_button.clicked.connect(
            lambda: self._set_all_groups_enabled(False)
        )
        self.invert_enabled_button.clicked.connect(self._invert_group_enabled)
        self.move_up_button.clicked.connect(lambda: self._move_selected_group(-1))
        self.move_down_button.clicked.connect(lambda: self._move_selected_group(1))
        self.delete_group_button.setToolTip(
            "사용자가 추가한 비교 그룹만 삭제할 수 있습니다."
        )
        toolbar.addWidget(self.restore_button)
        toolbar.addWidget(self.add_group_button)
        toolbar.addWidget(self.delete_group_button)
        toolbar.addSpacing(12)
        toolbar.addWidget(self.enable_all_button)
        toolbar.addWidget(self.disable_all_button)
        toolbar.addWidget(self.invert_enabled_button)
        toolbar.addSpacing(12)
        toolbar.addWidget(self.move_up_button)
        toolbar.addWidget(self.move_down_button)
        toolbar.addStretch(1)
        self.summary_label = QLabel()
        self.summary_label.setObjectName("locationLabel")
        toolbar.addWidget(self.summary_label)
        layout.addLayout(toolbar)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table.setWordWrap(False)
        self.table.setTextElideMode(Qt.TextElideMode.ElideRight)
        self.table.setMinimumHeight(340)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setMinimumSectionSize(40)
        self.table.verticalHeader().setDefaultSectionSize(42)
        self.table.cellClicked.connect(lambda row, _column: self.table.selectRow(row))
        self.table.itemChanged.connect(self._on_table_item_changed)
        self.table.itemSelectionChanged.connect(self._update_delete_button)
        layout.addWidget(self.table, stretch=1)

        bottom = QHBoxLayout()
        note = QLabel(
            "사용 체크를 끈 행은 매핑을 유지한 채 이미지 로딩에서 제외됩니다. "
            "자동 매핑 복원으로 기본 구성과 전체 활성 상태를 되돌릴 수 있습니다."
        )
        note.setWordWrap(True)
        note.setObjectName("subTitleLabel")
        bottom.addWidget(note, stretch=1)
        cancel_button = QPushButton("취소")
        apply_button = QPushButton("적용")
        cancel_button.clicked.connect(self.reject)
        apply_button.clicked.connect(self._accept_mapping)
        bottom.addWidget(cancel_button)
        bottom.addWidget(apply_button)
        layout.addLayout(bottom)

    def _apply_initial_window_size(self) -> None:
        screen = self.screen() or QApplication.primaryScreen()
        if screen is None:
            self.resize(1180, 720)
            return
        available = screen.availableGeometry()
        width = min(1400, max(self.minimumWidth(), int(available.width() * 0.88)))
        height = min(900, max(self.minimumHeight(), int(available.height() * 0.84)))
        self.resize(width, height)

    def _render_table(self, selected_row: Optional[int] = None) -> None:
        self._updating = True
        try:
            headers = ["사용", "순서", "비교 그룹"] + [
                ROLE_LABELS[role] for role in self._plan.roles
            ] + ["상태"]
            self.table.clear()
            self.table.setColumnCount(len(headers))
            self.table.setHorizontalHeaderLabels(headers)
            self.table.setRowCount(len(self._plan.groups))

            for row, group in enumerate(self._plan.groups):
                enabled_item = QTableWidgetItem(
                    "활성" if group.enabled else "비활성"
                )
                enabled_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                enabled_item.setFlags(
                    (enabled_item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                    & ~Qt.ItemFlag.ItemIsEditable
                )
                enabled_item.setCheckState(
                    Qt.CheckState.Checked
                    if group.enabled
                    else Qt.CheckState.Unchecked
                )
                enabled_item.setToolTip(
                    "체크를 끄면 이 비교 그룹은 이미지 로딩에서 제외됩니다."
                )
                self.table.setItem(row, 0, enabled_item)

                order_item = QTableWidgetItem(str(row + 1))
                order_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                order_item.setFlags(order_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row, 1, order_item)

                name_edit = QLineEdit(group.display_name)
                name_edit.setMinimumHeight(34)
                name_edit.setToolTip("Review 화면에 표시할 비교 그룹 이름")
                name_edit.setEnabled(group.enabled)
                name_edit.editingFinished.connect(
                    lambda edit=name_edit, target=group: self._set_group_name(
                        target, edit.text()
                    )
                )
                self.table.setCellWidget(row, 2, name_edit)

                for offset, role in enumerate(self._plan.roles, start=3):
                    combo = _TableScrollComboBox(self.table)
                    combo.setMinimumHeight(34)
                    combo.setEnabled(group.enabled)
                    combo.addItem("— 없음 —", None)
                    for source in self._sources_by_role[role]:
                        combo.addItem(
                            f"{source.original_index}. {source.name}",
                            source.source_sheet_id,
                        )
                    current_id = group.role_sheet_ids.get(role)
                    current_index = combo.findData(current_id)
                    combo.setCurrentIndex(max(0, current_index))
                    current_source = self._source_by_id.get(current_id)
                    current_text = (
                        current_source.name if current_source is not None else "없음"
                    )
                    combo.setToolTip(
                        f"{ROLE_LABELS[role]} 시트: {current_text}\n"
                        "새 비교 그룹에서 선택한 시트는 적용할 때 이 그룹으로 이동합니다."
                    )
                    popup_width = max(
                        combo.fontMetrics().horizontalAdvance(combo.itemText(index))
                        for index in range(combo.count())
                    ) + 48
                    combo.view().setMinimumWidth(min(420, popup_width))
                    combo.currentIndexChanged.connect(
                        lambda _index, target=group, selected_role=role, widget=combo:
                        self._assign_source(
                            target, selected_role, widget.currentData()
                        )
                    )
                    self.table.setCellWidget(row, offset, combo)

                status_item = QTableWidgetItem(self._status_text(group))
                status_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                status_item.setFlags(status_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row, len(headers) - 1, status_item)

                if not group.enabled:
                    muted_brush = QBrush(
                        self.palette().color(
                            QPalette.ColorGroup.Disabled,
                            QPalette.ColorRole.Text,
                        )
                    )
                    enabled_item.setForeground(muted_brush)
                    order_item.setForeground(muted_brush)
                    status_item.setForeground(muted_brush)

            header = self.table.horizontalHeader()
            header.setMinimumSectionSize(54)
            header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
            for column in range(3, 3 + len(self._plan.roles)):
                header.setSectionResizeMode(column, QHeaderView.ResizeMode.Stretch)
            header.setSectionResizeMode(
                len(headers) - 1, QHeaderView.ResizeMode.ResizeToContents
            )
            if self._plan.groups:
                row = 0 if selected_row is None else max(
                    0, min(selected_row, len(self._plan.groups) - 1)
                )
                self.table.selectRow(row)
            self._update_delete_button()
            enabled_count = sum(group.enabled for group in self._plan.groups)
            self.summary_label.setText(
                f"활성 {enabled_count}/{len(self._plan.groups)}개 · "
                f"원본 시트 {len(self._plan.sources)}개"
            )
        finally:
            self._updating = False

    def _status_text(self, group: SheetMappingGroup) -> str:
        labels = [
            ROLE_LABELS[role]
            for role in self._plan.roles
            if role in group.role_sheet_ids
        ]
        if len(labels) == 1:
            return f"{labels[0]} ONLY"
        return " + ".join(labels)

    def _set_group_name(self, group: SheetMappingGroup, text: str) -> None:
        group.display_name = str(text).strip()

    def _on_table_item_changed(self, item: QTableWidgetItem) -> None:
        if self._updating or item.column() != 0:
            return
        row = item.row()
        if not (0 <= row < len(self._plan.groups)):
            return
        self._plan.groups[row].enabled = (
            item.checkState() == Qt.CheckState.Checked
        )
        self._render_table(row)

    def _set_all_groups_enabled(self, enabled: bool) -> None:
        selected_row = self.table.currentRow()
        for group in self._plan.groups:
            group.enabled = bool(enabled)
        self._render_table(selected_row)

    def _invert_group_enabled(self) -> None:
        selected_row = self.table.currentRow()
        for group in self._plan.groups:
            group.enabled = not group.enabled
        self._render_table(selected_row)

    def _update_delete_button(self) -> None:
        if not hasattr(self, "delete_group_button"):
            return
        row = self.table.currentRow()
        can_delete = (
            0 <= row < len(self._plan.groups)
            and self._plan.groups[row].is_custom
        )
        self.delete_group_button.setEnabled(can_delete)

    def _delete_selected_custom_group(self) -> None:
        row = self.table.currentRow()
        if not (0 <= row < len(self._plan.groups)):
            return
        group = self._plan.groups[row]
        if not group.is_custom:
            return
        self._plan.groups.pop(row)
        self._custom_groups = [
            custom for custom in self._custom_groups if custom is not group
        ]
        self._render_table(min(row, len(self._plan.groups) - 1))

    def _assign_source(
        self,
        target_group: SheetMappingGroup,
        role: str,
        source_sheet_id: Optional[str],
    ) -> None:
        if self._updating:
            return
        previous_id = target_group.role_sheet_ids.get(role)
        if previous_id == source_sheet_id:
            return

        if self._is_custom_group(target_group):
            # 새 비교 그룹은 적용 전까지 기존 자동 매핑과 분리된 초안이다.
            # 다른 새 그룹에서 같은 역할/시트를 선택한 경우에만 그 선택을 옮긴다.
            if source_sheet_id is not None:
                for group in self._custom_groups:
                    if (
                        group is not target_group
                        and group.role_sheet_ids.get(role) == source_sheet_id
                    ):
                        group.role_sheet_ids.pop(role, None)
            target_group.role_sheet_ids.pop(role, None)
            if source_sheet_id is not None:
                target_group.role_sheet_ids[role] = source_sheet_id
            selected_row = self._plan.groups.index(target_group)
            self._render_table(selected_row)
            return

        if source_sheet_id is not None:
            for group in self._plan.groups:
                if group.role_sheet_ids.get(role) == source_sheet_id:
                    group.role_sheet_ids.pop(role, None)
                    group.matching_warning = ""
        target_group.role_sheet_ids.pop(role, None)
        if source_sheet_id is not None:
            target_group.role_sheet_ids[role] = source_sheet_id
            target_group.matching_warning = ""

        # 편집 도중에는 비어 있는 행이나 잠시 미배정된 시트를 정리하지 않는다.
        # 각 콤보 선택 사이의 중간 상태를 ONLY로 복구하면 사용자가 새 그룹에
        # 모으는 시트가 별도 행으로 다시 생길 수 있다. 최종 정리는 적용 시 한 번만 한다.
        try:
            selected_row = self._plan.groups.index(target_group)
        except ValueError:
            selected_row = max(0, len(self._plan.groups) - 1)
        self._render_table(selected_row)

    def _is_custom_group(self, target: SheetMappingGroup) -> bool:
        return any(group is target for group in self._custom_groups)

    def _reconcile_custom_groups(self) -> None:
        """새 그룹이 선택한 시트만 기존 그룹에서 제거해 최종 매핑을 만든다."""
        active_custom_groups = [
            custom
            for custom in self._custom_groups
            if any(group is custom for group in self._plan.groups)
        ]
        for custom in active_custom_groups:
            for role, source_id in tuple(custom.role_sheet_ids.items()):
                for group in self._plan.groups:
                    if group is custom:
                        continue
                    if group.role_sheet_ids.get(role) == source_id:
                        group.role_sheet_ids.pop(role, None)
                        group.matching_warning = ""

    def _remove_empty_groups(
        self, except_group: Optional[SheetMappingGroup] = None
    ) -> None:
        self._plan.groups = [
            group
            for group in self._plan.groups
            if group.role_sheet_ids or group is except_group
        ]

    def _append_unassigned_as_only_groups(self) -> None:
        used = {
            source_id
            for group in self._plan.groups
            for source_id in group.role_sheet_ids.values()
        }
        for source in self._plan.sources:
            if source.source_sheet_id not in used:
                self._plan.groups.append(
                    SheetMappingGroup(
                        display_name=source.name,
                        role_sheet_ids={source.role: source.source_sheet_id},
                    )
                )
                used.add(source.source_sheet_id)

    def _add_group(self) -> None:
        number = 1 + sum(
            group.display_name.startswith("새 비교 그룹")
            for group in self._plan.groups
        )
        group = SheetMappingGroup(
            display_name=f"새 비교 그룹 {number}",
            is_custom=True,
        )
        self._plan.groups.append(group)
        self._custom_groups.append(group)
        self._render_table(len(self._plan.groups) - 1)

    def _move_selected_group(self, offset: int) -> None:
        row = self.table.currentRow()
        if not (0 <= row < len(self._plan.groups)):
            return
        destination = max(0, min(len(self._plan.groups) - 1, row + int(offset)))
        if destination == row:
            return
        group = self._plan.groups.pop(row)
        self._plan.groups.insert(destination, group)
        self._render_table(destination)

    def _restore_automatic_mapping(self) -> None:
        self._plan = self._automatic_plan.clone()
        for group in self._plan.groups:
            group.enabled = True
        self._custom_groups = []
        self._source_by_id = {
            source.source_sheet_id: source for source in self._plan.sources
        }
        self._sources_by_role = {
            role: [source for source in self._plan.sources if source.role == role]
            for role in self._plan.roles
        }
        self._render_table()

    def _accept_mapping(self) -> None:
        self._reconcile_custom_groups()
        self._remove_empty_groups()
        self._append_unassigned_as_only_groups()
        for group in self._plan.groups:
            if not group.display_name.strip() and group.role_sheet_ids:
                first_id = next(iter(group.role_sheet_ids.values()))
                group.display_name = self._source_by_id[first_id].name
        try:
            validate_sheet_mapping_plan(self._plan)
        except ValueError as exc:
            QMessageBox.warning(self, "시트 매핑 확인", str(exc))
            self._render_table(self.table.currentRow())
            return
        self.result_plan = self._plan.clone()
        self.accept()
