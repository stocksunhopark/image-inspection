"""Navigation state for Double/Triple visual inspection items."""

from __future__ import annotations

import pytest

from app_state import AppState
from models import ExtractedImage, InspectionItem


def _image(
    sheet_index: int,
    cell: str,
    row: int,
    col: int,
    *,
    null: bool = False,
) -> ExtractedImage:
    path = None if null else f"sheet-{sheet_index}-{cell}.png"
    return ExtractedImage(
        sheet_index=sheet_index,
        sheet_name=f"Sheet {sheet_index}",
        cell_address=cell,
        merged_range="",
        anchor_row=row,
        anchor_col=col,
        is_null=null,
        source_path=path,
        preview_path=path,
    )


def _item(
    sheet_index: int,
    index: int,
    cell: str,
    *,
    triple: bool = False,
) -> InspectionItem:
    row = index - 1
    return InspectionItem(
        sheet_index=sheet_index,
        index=index,
        image_ref=_image(sheet_index, cell, row, 0),
        image_a=_image(sheet_index, cell, row, 0),
        image_b=_image(sheet_index, cell, row, 0) if triple else None,
    )


def test_empty_state_has_no_current_item_and_navigation_is_safe():
    state = AppState()

    assert state.mode == "double"
    assert state.items_by_sheet == {}
    assert state.flat_items == []
    assert state.current_index == -1
    assert state.current_item is None
    assert state.current_sheet_position() == (0, 0)
    assert state.move(1) is False
    assert state.go_first() is False
    assert state.go_last() is False
    assert state.select_sheet(1) is False


def test_set_items_orders_sheet_tabs_and_retains_job_context():
    s1_a = _item(1, 1, "A2")
    s1_b = _item(1, 2, "D10")
    s2_a = _item(2, 1, "C3")
    state = AppState()

    state.set_items(
        {2: [s2_a], 3: [], 1: [s1_a, s1_b]},
        mode="double",
        workbook_paths={"ref": "ref.xlsx", "a": "a.xlsx"},
        preview_temp_dir="preview-double",
    )

    assert list(state.items_by_sheet) == [1, 2, 3]
    assert state.flat_items == [s1_a, s1_b, s2_a]
    assert state.current_index == 0
    assert state.current_item is s1_a
    assert state.current_sheet_position() == (1, 2)
    assert state.workbook_paths == {"ref": "ref.xlsx", "a": "a.xlsx"}
    assert state.preview_temp_dir == "preview-double"


def test_move_clamps_at_bounds_and_select_sheet_targets_its_first_item():
    s1_a = _item(1, 1, "A2")
    s1_b = _item(1, 2, "D10")
    s2_a = _item(2, 1, "C3")
    state = AppState()
    state.set_items({1: [s1_a, s1_b], 2: [s2_a], 3: []}, mode="double")

    assert state.move(-100) is False
    assert state.current_index == 0
    assert state.current_sheet_position() == (1, 2)

    assert state.move(1) is True
    assert state.current_item is s1_b
    assert state.current_sheet_position() == (2, 2)

    assert state.move(100) is True
    assert state.current_index == 2
    assert state.current_item is s2_a
    assert state.current_sheet_position() == (1, 1)
    assert state.move(1) is False
    assert state.current_index == 2

    assert state.select_sheet(1) is True
    assert state.current_index == 0
    assert state.current_item is s1_a
    assert state.select_sheet(1) is False

    # Empty or unknown sheets leave the current selection intact.
    assert state.select_sheet(3) is False
    assert state.select_sheet(999) is False
    assert state.current_item is s1_a

    assert state.go_last() is True
    assert state.current_item is s2_a
    assert state.go_last() is False
    assert state.go_first() is True
    assert state.current_item is s1_a


def test_triple_replacement_and_clear_reset_selection_and_context():
    old = _item(1, 1, "A2")
    triple = _item(7, 1, "C3", triple=True)
    state = AppState()
    state.set_items({1: [old]}, mode="double", preview_temp_dir="old-preview")

    state.set_items(
        {7: [triple]},
        mode="triple",
        workbook_paths={"ref": "r.xlsx", "a": "a.xlsx", "b": "b.xlsx"},
        preview_temp_dir="triple-preview",
    )
    assert state.mode == "triple"
    assert state.flat_items == [triple]
    assert state.current_index == 0
    assert state.current_item is triple
    assert triple.image_b is not None
    assert state.workbook_paths["b"] == "b.xlsx"
    assert state.preview_temp_dir == "triple-preview"

    state.clear()
    assert state.items_by_sheet == {}
    assert state.flat_items == []
    assert state.current_index == -1
    assert state.current_item is None
    assert state.current_sheet_position() == (0, 0)
    assert state.workbook_paths == {}
    assert state.preview_temp_dir is None


def test_invalid_mode_is_rejected_without_mutating_existing_state():
    existing = _item(1, 1, "A2")
    replacement = _item(2, 1, "B2")
    state = AppState()
    state.set_items(
        {1: [existing]},
        mode="double",
        workbook_paths={"ref": "ref.xlsx", "a": "a.xlsx"},
        preview_temp_dir="preview",
    )

    with pytest.raises(ValueError, match="모드"):
        state.set_items({2: [replacement]}, mode="single")

    assert state.mode == "double"
    assert state.flat_items == [existing]
    assert state.current_item is existing
    assert state.workbook_paths == {"ref": "ref.xlsx", "a": "a.xlsx"}
    assert state.preview_temp_dir == "preview"

