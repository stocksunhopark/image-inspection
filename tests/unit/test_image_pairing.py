"""엑셀 이미지 앵커 짝짓기 단위 테스트."""
from __future__ import annotations

from types import SimpleNamespace

from excel_manager import merged_origin_0based, pair_image_positions


class _Merge:
    def __init__(self, min_row, min_col, max_row, max_col):
        self.min_row = min_row
        self.min_col = min_col
        self.max_row = max_row
        self.max_col = max_col


def _ws(merges):
    return SimpleNamespace(merged_cells=SimpleNamespace(ranges=merges))


def test_exact_positions_pair_one_to_one():
    objs_a = {(10, 3): "a", (20, 3): "a2"}
    objs_b = {(10, 3): "b", (20, 3): "b2"}
    pairs = pair_image_positions(objs_a, objs_b)
    assert pairs == [((10, 3), (10, 3)), ((20, 3), (20, 3))]


def test_same_merged_origin_pairs_shifted_cell_inside_merge():
    # 같은 병합 D271:J290 (1-based row 271-290, col 4-10) 안에서 한 칸 내려간 앵커.
    merge = _Merge(271, 4, 290, 10)
    ws = _ws([merge])
    objs_a = {(270, 3): "a"}  # D271
    objs_b = {(271, 3): "b"}  # D272
    pairs = pair_image_positions(objs_a, objs_b, ws, ws)
    assert pairs == [((270, 3), (271, 3))]


def test_plus_one_row_pairs_title_row_to_block():
    # LOW는 제목 행 D270:J270, ROOM은 본문 D271:J290. 병합 원점은 다르지만 1행 차.
    ws_a = _ws([_Merge(270, 4, 270, 10)])
    ws_b = _ws([_Merge(271, 4, 290, 10)])
    objs_a = {(69, 3): "keep", (269, 3): "shifted"}  # D70, D270
    objs_b = {(69, 3): "keep", (270, 3): "shifted"}  # D70, D271
    pairs = pair_image_positions(objs_a, objs_b, ws_a, ws_b)
    assert ((69, 3), (69, 3)) in pairs
    assert ((269, 3), (270, 3)) in pairs
    assert len(pairs) == 2


def test_does_not_steal_exact_match_for_neighbor():
    objs_a = {(10, 3): "a", (11, 3): "a2"}
    objs_b = {(10, 3): "b"}
    pairs = pair_image_positions(objs_a, objs_b)
    assert ((10, 3), (10, 3)) in pairs
    assert ((11, 3), None) in pairs
    assert len(pairs) == 2


def test_unmatched_remain_one_sided():
    objs_a = {(10, 3): "a"}
    objs_b = {(50, 8): "b"}
    pairs = pair_image_positions(objs_a, objs_b)
    assert ((10, 3), None) in pairs
    assert (None, (50, 8)) in pairs


def test_merged_origin_helper():
    ws = _ws([_Merge(271, 4, 290, 10)])
    assert merged_origin_0based(ws, 270, 3) == (270, 3)
    assert merged_origin_0based(ws, 275, 5) == (270, 3)
    assert merged_origin_0based(ws, 0, 0) == (0, 0)
