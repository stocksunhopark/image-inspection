"""행 기반 시트 매핑 설정 창 테스트."""

from PyQt6.QtCore import QPoint, QPointF, Qt
from PyQt6.QtGui import QWheelEvent
from PyQt6.QtWidgets import QApplication

from sheet_mapping import build_automatic_sheet_mapping
from ui.sheet_mapping_dialog import SheetMappingDialog


def _automatic_plan():
    return build_automatic_sheet_mapping(
        {
            "ref": [(1, "MAIN"), (2, "CLK")],
            "a": [(1, "MAIN"), (2, "Clock")],
            "b": [(1, "MAIN"), (2, "CLK_MAIN")],
        },
        ("ref", "a", "b"),
    )


def test_dialog_shows_auto_groups_exact_matches_and_only_rows(qapp):
    automatic = _automatic_plan()
    dialog = SheetMappingDialog(automatic, automatic)
    try:
        assert dialog.windowTitle() == "시트 순서 설정"
        assert dialog.table.rowCount() == 4
        assert dialog.table.columnCount() == 7
        assert dialog.table.horizontalHeaderItem(0).text() == "사용"
        assert dialog.table.item(0, 0).checkState() == Qt.CheckState.Checked
        assert dialog.table.item(0, 6).text() == "REF + A + B"
        assert dialog.table.item(1, 6).text() == "REF ONLY"
        assert dialog.table.item(2, 6).text() == "A ONLY"
        assert dialog.restore_button.text() == "자동 매핑 복원"
        assert dialog.add_group_button.text() == "비교 그룹 추가"
        assert dialog.enable_all_button.text() == "전체 활성"
        assert dialog.disable_all_button.text() == "전체 비활성"
        assert dialog.invert_enabled_button.text() == "선택 반전"
        assert dialog.windowFlags() & Qt.WindowType.WindowMaximizeButtonHint
        assert dialog.windowFlags() & Qt.WindowType.WindowMinimizeButtonHint
        assert dialog.isSizeGripEnabled()
        assert dialog.table.verticalHeader().defaultSectionSize() >= 42
        assert dialog.table.minimumHeight() >= 340
    finally:
        dialog.close()
        dialog.deleteLater()


def test_group_checkbox_disables_row_without_changing_mapping(qapp):
    automatic = _automatic_plan()
    dialog = SheetMappingDialog(automatic, automatic)
    try:
        dialog.show()
        qapp.processEvents()
        original_mapping = dict(dialog._plan.groups[1].role_sheet_ids)

        dialog.table.item(1, 0).setCheckState(Qt.CheckState.Unchecked)
        qapp.processEvents()

        assert dialog._plan.groups[1].enabled is False
        assert dialog._plan.groups[1].role_sheet_ids == original_mapping
        assert dialog.table.item(1, 0).text() == "비활성"
        assert dialog.table.cellWidget(1, 2).isEnabled() is False
        assert dialog.table.cellWidget(1, 3).isEnabled() is False
        assert "활성 3/4개" in dialog.summary_label.text()

        dialog.table.item(1, 0).setCheckState(Qt.CheckState.Checked)
        qapp.processEvents()
        assert dialog._plan.groups[1].enabled is True
        assert dialog._plan.groups[1].role_sheet_ids == original_mapping
    finally:
        dialog.close()
        dialog.deleteLater()


def test_bulk_activation_custom_delete_and_automatic_restore(qapp):
    automatic = _automatic_plan()
    dialog = SheetMappingDialog(automatic, automatic)
    try:
        assert dialog.delete_group_button.isEnabled() is False

        dialog._add_group()
        assert dialog._plan.groups[-1].is_custom is True
        assert dialog.delete_group_button.isEnabled() is True
        dialog.delete_group_button.click()
        qapp.processEvents()
        assert len(dialog._plan.groups) == len(automatic.groups)

        dialog.disable_all_button.click()
        assert all(group.enabled is False for group in dialog._plan.groups)
        dialog.invert_enabled_button.click()
        assert all(group.enabled is True for group in dialog._plan.groups)
        dialog._plan.groups[0].enabled = False
        dialog._add_group()

        dialog.restore_button.click()
        assert len(dialog._plan.groups) == len(automatic.groups)
        assert all(group.enabled is True for group in dialog._plan.groups)
        assert all(group.is_custom is False for group in dialog._plan.groups)
    finally:
        dialog.close()
        dialog.deleteLater()


def test_dialog_custom_group_moves_different_names_without_duplicates(qapp):
    automatic = _automatic_plan()
    dialog = SheetMappingDialog(automatic, automatic)
    try:
        dialog._add_group()
        target = dialog._plan.groups[-1]
        ids = {
            (source.role, source.name): source.source_sheet_id
            for source in dialog._plan.sources
        }
        dialog._assign_source(target, "ref", ids[("ref", "CLK")])
        dialog._assign_source(target, "a", ids[("a", "Clock")])
        dialog._assign_source(target, "b", ids[("b", "CLK_MAIN")])

        assert len(dialog._plan.groups) == 5
        assert target.display_name == "새 비교 그룹 1"
        assert target.role_sheet_ids == {
            "ref": ids[("ref", "CLK")],
            "a": ids[("a", "Clock")],
            "b": ids[("b", "CLK_MAIN")],
        }
        dialog._accept_mapping()
        assert dialog.result_plan is not None
        assert len(dialog.result_plan.groups) == 2
        assert dialog.result_plan.groups[-1].display_name == "새 비교 그룹 1"
    finally:
        dialog.close()
        dialog.deleteLater()


def test_added_group_name_changes_only_when_user_edits_it(qapp):
    automatic = _automatic_plan()
    dialog = SheetMappingDialog(automatic, automatic)
    try:
        dialog._add_group()
        target = dialog._plan.groups[-1]
        selected = next(
            source
            for source in dialog._plan.sources
            if source.role == "ref" and source.name == "CLK"
        )

        dialog._assign_source(target, "ref", selected.source_sheet_id)
        assert target.display_name == "새 비교 그룹 1"

        dialog._set_group_name(target, "클럭 비교")
        dialog._assign_source(target, "a", None)
        assert target.display_name == "클럭 비교"
    finally:
        dialog.close()
        dialog.deleteLater()


def test_dialog_reorders_and_restores_automatic_mapping(qapp):
    automatic = _automatic_plan()
    dialog = SheetMappingDialog(automatic, automatic)
    try:
        dialog.table.selectRow(2)
        dialog._move_selected_group(-1)
        assert [group.display_name for group in dialog._plan.groups][:3] == [
            "MAIN",
            "Clock",
            "CLK",
        ]

        dialog._restore_automatic_mapping()
        assert [group.display_name for group in dialog._plan.groups] == [
            "MAIN",
            "CLK",
            "Clock",
            "CLK_MAIN",
        ]
    finally:
        dialog.close()
        dialog.deleteLater()


def test_many_sheet_rows_scroll_and_dialog_can_maximize(qapp):
    names = [(index, f"SHEET_{index:02d}_LONG_DISPLAY_NAME") for index in range(1, 51)]
    automatic = build_automatic_sheet_mapping(
        {"ref": names, "a": names}, ("ref", "a")
    )
    dialog = SheetMappingDialog(automatic, automatic)
    try:
        dialog.show()
        qapp.processEvents()
        assert dialog.table.rowCount() == 50
        assert dialog.table.verticalScrollBar().maximum() > 0
        assert dialog.table.rowHeight(0) >= 40
        first_combo = dialog.table.cellWidget(0, 3)
        assert first_combo.minimumHeight() >= 34

        dialog.showMaximized()
        qapp.processEvents()
        assert dialog.isMaximized()
    finally:
        dialog.close()
        dialog.deleteLater()


def test_mouse_wheel_over_sheet_combo_does_not_change_selection(qapp):
    names = [(index, f"SHEET_{index:02d}") for index in range(1, 41)]
    automatic = build_automatic_sheet_mapping(
        {"ref": names, "a": names}, ("ref", "a")
    )
    dialog = SheetMappingDialog(automatic, automatic)
    try:
        dialog.show()
        qapp.processEvents()
        combo = dialog.table.cellWidget(0, 3)
        combo.setCurrentIndex(1)
        original_index = combo.currentIndex()
        scrollbar = dialog.table.verticalScrollBar()
        assert scrollbar.maximum() > 0
        assert scrollbar.value() == 0
        event = QWheelEvent(
            QPointF(5, 5),
            QPointF(combo.mapToGlobal(QPoint(5, 5))),
            QPoint(0, 0),
            QPoint(0, -120),
            Qt.MouseButton.NoButton,
            Qt.KeyboardModifier.NoModifier,
            Qt.ScrollPhase.ScrollUpdate,
            False,
        )

        QApplication.sendEvent(combo, event)

        assert combo.currentIndex() == original_index
        assert scrollbar.value() > 0
    finally:
        dialog.close()
        dialog.deleteLater()


def test_real_combo_edits_accumulate_four_roles_in_one_added_group(qapp):
    common_names = [(index, f"COMMON_{index:02d}") for index in range(1, 18)]
    automatic = build_automatic_sheet_mapping(
        {
            "ref": common_names,
            "a": common_names,
            "b": common_names,
            "c": common_names,
        },
        ("ref", "a", "b", "c"),
    )
    dialog = SheetMappingDialog(automatic, automatic)
    try:
        dialog.show()
        qapp.processEvents()
        dialog._add_group()
        qapp.processEvents()
        target = dialog._plan.groups[-1]
        chosen = {
            "ref": next(
                source.source_sheet_id
                for source in dialog._plan.sources
                if source.role == "ref" and source.name == "COMMON_01"
            ),
            "a": next(
                source.source_sheet_id
                for source in dialog._plan.sources
                if source.role == "a" and source.name == "COMMON_02"
            ),
            "b": next(
                source.source_sheet_id
                for source in dialog._plan.sources
                if source.role == "b" and source.name == "COMMON_03"
            ),
            "c": next(
                source.source_sheet_id
                for source in dialog._plan.sources
                if source.role == "c" and source.name == "COMMON_04"
            ),
        }
        for role, source_id in chosen.items():
            row = dialog._plan.groups.index(target)
            column = 3 + dialog._plan.roles.index(role)
            combo = dialog.table.cellWidget(row, column)
            combo.setCurrentIndex(combo.findData(source_id))
            qapp.processEvents()

        assert len(dialog._plan.groups) == 18
        assert target.role_sheet_ids == chosen
        assert sum(
            source_id in group.role_sheet_ids.values()
            for group in dialog._plan.groups
            for source_id in chosen.values()
        ) == 8
        assert [group.display_name for group in dialog._plan.groups[18:]] == []

        dialog._accept_mapping()
        assert dialog.result_plan is not None
        assert sum(
            source_id in group.role_sheet_ids.values()
            for group in dialog.result_plan.groups
            for source_id in chosen.values()
        ) == 4
    finally:
        dialog.close()
        dialog.deleteLater()


def test_only_sources_do_not_spawn_rows_while_added_group_is_being_edited(qapp):
    common = [(index, f"COMMON_{index:02d}") for index in range(1, 14)]
    automatic = build_automatic_sheet_mapping(
        {
            "ref": common + [(14, "REF_DIFFERENT")],
            "a": common + [(14, "A_DIFFERENT")],
            "b": common + [(14, "B_DIFFERENT")],
            "c": common + [(14, "C_DIFFERENT")],
        },
        ("ref", "a", "b", "c"),
    )
    assert len(automatic.groups) == 17
    dialog = SheetMappingDialog(automatic, automatic)
    try:
        original_groups = [
            (group.display_name, dict(group.role_sheet_ids))
            for group in dialog._plan.groups
        ]
        dialog.show()
        qapp.processEvents()
        dialog._add_group()
        qapp.processEvents()
        target = dialog._plan.groups[-1]
        chosen_names = {
            "ref": "REF_DIFFERENT",
            "a": "A_DIFFERENT",
            "b": "B_DIFFERENT",
            "c": "C_DIFFERENT",
        }
        chosen_ids = {}
        for role, name in chosen_names.items():
            source = next(
                source
                for source in dialog._plan.sources
                if source.role == role and source.name == name
            )
            chosen_ids[role] = source.source_sheet_id
            row = dialog._plan.groups.index(target)
            column = 3 + dialog._plan.roles.index(role)
            combo = dialog.table.cellWidget(row, column)
            combo.setCurrentIndex(combo.findData(source.source_sheet_id))
            qapp.processEvents()

            assert len(dialog._plan.groups) == 18
            assert dialog._plan.groups.index(target) == 17
            assert len(dialog._plan.groups[18:]) == 0

        assert target.role_sheet_ids == chosen_ids
        assert sum(
            source_id in group.role_sheet_ids.values()
            for group in dialog._plan.groups
            for source_id in chosen_ids.values()
        ) == 8
        assert [
            (group.display_name, dict(group.role_sheet_ids))
            for group in dialog._plan.groups[:17]
        ] == original_groups
        assert target.display_name == "새 비교 그룹 1"

        dialog._accept_mapping()
        assert dialog.result_plan is not None
        assert len(dialog.result_plan.groups) == 14
        assert dialog.result_plan.groups[-1].role_sheet_ids == chosen_ids
    finally:
        dialog.close()
        dialog.deleteLater()
