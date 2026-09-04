"""Double/Triple Excel loading for the side-by-side visual inspector."""

from __future__ import annotations

import shutil
import tempfile
from pathlib import Path
from typing import Dict, Tuple

import pytest
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image

from excel_manager import load_extracted_pil
from inspection_service import load_inspection_images
from models import ExtractedImage, InspectionItem


Color = Tuple[int, int, int]
WorkbookSpec = Dict[str, Dict[str, object]]


REF_COLORS: Dict[str, Color] = {
    "A2": (220, 20, 20),
    "D10": (20, 220, 20),
    "F20": (20, 20, 220),
    "J30": (180, 20, 180),
    "P60": (20, 180, 180),
    "SECOND!C3": (120, 80, 40),
}
A_COLORS: Dict[str, Color] = {
    "A2": (210, 40, 30),
    "D11": (30, 210, 40),
    "H22": (30, 40, 210),
    "L40": (210, 160, 30),
    "P60": (40, 170, 170),
    "SECOND!C3": (100, 90, 50),
}
B_COLORS: Dict[str, Color] = {
    "A2": (200, 50, 40),
    "D9": (40, 200, 50),
    "G21": (40, 50, 200),
    "J30": (170, 40, 170),
    "L41": (200, 150, 40),
    "N50": (80, 80, 80),
    "SECOND!C3": (80, 100, 60),
}


def _write_workbook(path: Path, spec: WorkbookSpec) -> Path:
    asset_dir = path.parent / f"{path.stem}-images"
    asset_dir.mkdir()
    workbook = Workbook()

    for sheet_number, (sheet_name, sheet_spec) in enumerate(spec.items()):
        worksheet = workbook.active if sheet_number == 0 else workbook.create_sheet()
        worksheet.title = sheet_name
        for merged_range in sheet_spec.get("merges", ()):  # type: ignore[union-attr]
            worksheet.merge_cells(str(merged_range))
        images = sheet_spec.get("images", {})  # type: ignore[union-attr]
        for cell, color in images.items():  # type: ignore[union-attr]
            image_path = asset_dir / f"{sheet_name}-{cell}.png"
            Image.new("RGB", (18, 12), color).save(image_path, format="PNG")
            worksheet.add_image(XLImage(str(image_path)), str(cell))

    workbook.save(path)
    workbook.close()
    return path


@pytest.fixture
def inspection_books(tmp_path: Path) -> Dict[str, Path]:
    common_merge = ["F20:H22"]
    ref = _write_workbook(
        tmp_path / "ref.xlsx",
        {
            "MAIN": {
                "merges": common_merge,
                # Deliberately not in coordinate order.
                "images": {
                    "P60": REF_COLORS["P60"],
                    "J30": REF_COLORS["J30"],
                    "F20": REF_COLORS["F20"],
                    "D10": REF_COLORS["D10"],
                    "A2": REF_COLORS["A2"],
                },
            },
            "SECOND": {"images": {"C3": REF_COLORS["SECOND!C3"]}},
        },
    )
    compare_a = _write_workbook(
        tmp_path / "a.xlsx",
        {
            "MAIN": {
                "merges": common_merge,
                "images": {
                    "L40": A_COLORS["L40"],
                    "H22": A_COLORS["H22"],
                    "P60": A_COLORS["P60"],
                    "D11": A_COLORS["D11"],
                    "A2": A_COLORS["A2"],
                },
            },
            "SECOND": {"images": {"C3": A_COLORS["SECOND!C3"]}},
        },
    )
    compare_b = _write_workbook(
        tmp_path / "b.xlsx",
        {
            "MAIN": {
                "merges": common_merge,
                "images": {
                    "N50": B_COLORS["N50"],
                    "L41": B_COLORS["L41"],
                    "J30": B_COLORS["J30"],
                    "G21": B_COLORS["G21"],
                    "D9": B_COLORS["D9"],
                    "A2": B_COLORS["A2"],
                },
            },
            "SECOND": {"images": {"C3": B_COLORS["SECOND!C3"]}},
        },
    )
    return {"ref": ref, "a": compare_a, "b": compare_b}


def _all_items(items_by_sheet: Dict[int, list[InspectionItem]]) -> list[InspectionItem]:
    return [
        item
        for sheet_index in sorted(items_by_sheet)
        for item in items_by_sheet[sheet_index]
    ]


def _assert_real_lazy_image(
    extracted: ExtractedImage,
    preview_dir: str,
    expected_color: Color,
) -> None:
    assert extracted.is_null is False
    assert extracted.image is None
    assert extracted.source_path
    assert extracted.preview_path == extracted.source_path
    source_path = Path(extracted.source_path).resolve()
    assert source_path.is_file()
    assert Path(preview_dir).resolve() in source_path.parents
    loaded = load_extracted_pil(extracted)
    assert loaded is not None
    assert loaded.mode == "RGB"
    assert loaded.getpixel((0, 0)) == expected_color


def _assert_null_image(extracted: ExtractedImage, expected_cell: str) -> None:
    assert extracted.is_null is True
    assert extracted.cell_address == expected_cell
    assert extracted.image is None
    assert extracted.source_path is None
    assert extracted.preview_path is None
    assert load_extracted_pil(extracted) is None


def test_double_matches_exact_adjacent_and_merged_positions_and_keeps_missing_sides(
    inspection_books: Dict[str, Path],
):
    items_by_sheet, preview_dir = load_inspection_images(
        str(inspection_books["ref"]), str(inspection_books["a"])
    )
    try:
        assert list(items_by_sheet) == [1, 2]
        main = items_by_sheet[1]
        assert [item.index for item in main] == [1, 2, 3, 4, 5, 6]
        assert [item.cell_address for item in main] == [
            "A2",
            "D10",
            "F20",
            "J30",
            "L40",
            "P60",
        ]

        # Exact anchor, one-row tolerance, then same merged block even though
        # the anchors differ by two rows and two columns.
        assert (main[0].image_ref.cell_address, main[0].image_a.cell_address) == (
            "A2",
            "A2",
        )
        assert (main[1].image_ref.cell_address, main[1].image_a.cell_address) == (
            "D10",
            "D11",
        )
        assert (main[2].image_ref.cell_address, main[2].image_a.cell_address) == (
            "F20",
            "H22",
        )
        assert main[2].image_ref.merged_range == "F20:H22"
        assert main[2].image_a.merged_range == "F20:H22"

        _assert_null_image(main[3].image_a, "J30")
        _assert_null_image(main[4].image_ref, "L40")
        assert main[3].image_ref.cell_address == "J30"
        assert main[4].image_a.cell_address == "L40"
        assert all(item.image_b is None for item in _all_items(items_by_sheet))

        assert [item.cell_address for item in items_by_sheet[2]] == ["C3"]
        assert items_by_sheet[2][0].index == 1
    finally:
        shutil.rmtree(preview_dir, ignore_errors=True)


def test_triple_aligns_all_sides_persists_lazy_sources_and_reports_progress(
    inspection_books: Dict[str, Path],
):
    progress = []
    statuses = []
    items_by_sheet, preview_dir = load_inspection_images(
        str(inspection_books["ref"]),
        str(inspection_books["a"]),
        str(inspection_books["b"]),
        progress_cb=lambda current, total, cell: progress.append(
            (current, total, cell)
        ),
        status_cb=statuses.append,
    )
    try:
        main = items_by_sheet[1]
        assert [item.cell_address for item in main] == [
            "A2",
            "D10",
            "F20",
            "J30",
            "L40",
            "N50",
            "P60",
        ]
        assert [
            (
                item.image_ref.cell_address,
                item.image_a.cell_address,
                item.image_b.cell_address if item.image_b is not None else None,
            )
            for item in main[:3]
        ] == [
            ("A2", "A2", "A2"),
            ("D10", "D11", "D9"),
            ("F20", "H22", "G21"),
        ]

        # A is absent, Ref is absent, then both Ref/A are absent.
        assert main[3].image_b is not None
        _assert_null_image(main[3].image_a, "J30")
        _assert_null_image(main[4].image_ref, "L40")
        assert main[4].image_b is not None
        assert main[4].image_b.cell_address == "L41"
        _assert_null_image(main[5].image_ref, "N50")
        _assert_null_image(main[5].image_a, "N50")
        assert main[5].image_b is not None
        assert main[5].image_b.cell_address == "N50"
        assert main[6].image_b is not None
        _assert_null_image(main[6].image_b, "P60")

        first = main[0]
        assert first.image_b is not None
        _assert_real_lazy_image(first.image_ref, preview_dir, REF_COLORS["A2"])
        _assert_real_lazy_image(first.image_a, preview_dir, A_COLORS["A2"])
        _assert_real_lazy_image(first.image_b, preview_dir, B_COLORS["A2"])
        _assert_real_lazy_image(main[4].image_a, preview_dir, A_COLORS["L40"])
        assert main[4].image_b is not None
        _assert_real_lazy_image(main[4].image_b, preview_dir, B_COLORS["L41"])

        all_items = _all_items(items_by_sheet)
        assert len(all_items) == 8
        assert [current for current, _total, _cell in progress] == list(range(1, 9))
        assert {total for _current, total, _cell in progress} == {8}
        assert all(str(cell).strip() for _current, _total, cell in progress)
        assert statuses and all(str(message).strip() for message in statuses)

        # Sources have been copied lazily, so all input workbooks can be moved
        # after the service returns (important for Windows file handles).
        for role, workbook_path in inspection_books.items():
            moved = workbook_path.with_name(f"{role}-moved.xlsx")
            workbook_path.rename(moved)
            assert moved.is_file()
        assert load_extracted_pil(first.image_ref) is not None
    finally:
        shutil.rmtree(preview_dir, ignore_errors=True)


def test_cancellation_raises_and_removes_partial_preview_directory(
    inspection_books: Dict[str, Path],
    tmp_path: Path,
    monkeypatch,
):
    service_temp_root = tmp_path / "service-temp"
    service_temp_root.mkdir()
    monkeypatch.setattr(tempfile, "tempdir", str(service_temp_root))
    progress = []

    with pytest.raises(InterruptedError, match="취소"):
        load_inspection_images(
            str(inspection_books["ref"]),
            str(inspection_books["a"]),
            str(inspection_books["b"]),
            progress_cb=lambda current, total, cell: progress.append(
                (current, total, cell)
            ),
            cancel_cb=lambda: bool(progress),
        )

    assert len(progress) == 1
    assert progress[0][:2] == (1, 8)
    assert list(service_temp_root.iterdir()) == []


def test_mismatched_sheet_tabs_are_rejected(tmp_path: Path):
    ref = _write_workbook(
        tmp_path / "ref-mismatch.xlsx",
        {"MAIN": {"images": {"A1": (10, 20, 30)}}, "SECOND": {"images": {}}},
    )
    compare = _write_workbook(
        tmp_path / "a-mismatch.xlsx",
        {"MAIN": {"images": {"A1": (10, 20, 30)}}, "OTHER": {"images": {}}},
    )

    with pytest.raises(ValueError, match="시트 탭 구성"):
        load_inspection_images(str(ref), str(compare))


def test_duplicate_image_anchor_is_rejected_instead_of_dropping_an_image(
    tmp_path: Path,
):
    source_a = tmp_path / "duplicate-a.png"
    source_b = tmp_path / "duplicate-b.png"
    Image.new("RGB", (12, 8), (255, 0, 0)).save(source_a)
    Image.new("RGB", (12, 8), (0, 255, 0)).save(source_b)
    ref_book = Workbook()
    ref_sheet = ref_book.active
    ref_sheet.title = "MAIN"
    ref_sheet.add_image(XLImage(str(source_a)), "A1")
    ref_sheet.add_image(XLImage(str(source_b)), "A1")
    ref = tmp_path / "ref-duplicate.xlsx"
    ref_book.save(ref)
    ref_book.close()

    compare = _write_workbook(
        tmp_path / "a-duplicate.xlsx",
        {"MAIN": {"images": {"A1": (10, 20, 30)}}},
    )

    with pytest.raises(ValueError, match="이미지가 2개 이상"):
        load_inspection_images(str(ref), str(compare))
