"""Double/Triple Excel에서 육안 검사용 이미지 묶음을 만드는 서비스."""

import os
import shutil
import tempfile
from typing import Callable, Dict, List, Optional, Tuple

from openpyxl import load_workbook

from excel_manager import (
    Position,
    extract_image_meta,
    extract_image_objects_by_position,
    make_null_image,
    pair_image_positions,
    persist_source_from_obj,
)
from models import ExtractedImage, InspectionItem

ProgressCallback = Callable[[int, int, str], None]
StatusCallback = Callable[[str], None]
CancelCallback = Callable[[], bool]
AlignedPositions = Tuple[Optional[Position], Optional[Position], Optional[Position]]


def _align_positions(
    objects_ref: Dict[Position, object],
    objects_a: Dict[Position, object],
    objects_b: Optional[Dict[Position, object]],
    worksheet_ref=None,
    worksheet_a=None,
    worksheet_b=None,
    cancel_cb: Optional[CancelCallback] = None,
) -> List[AlignedPositions]:
    """Ref/A 또는 Ref/A/B 위치를 한 화면 행 단위로 정렬한다."""
    ref_a_pairs = pair_image_positions(
        objects_ref,
        objects_a,
        worksheet_ref,
        worksheet_a,
        cancel_cb=cancel_cb,
    )
    groups: List[List[Optional[Position]]] = [
        [position_ref, position_a, None]
        for position_ref, position_a in ref_a_pairs
    ]
    if objects_b is None:
        return [(group[0], group[1], None) for group in groups]

    group_by_ref = {
        group[0]: group for group in groups if group[0] is not None
    }
    unmatched_b: Dict[Position, object] = {}
    for position_ref, position_b in pair_image_positions(
        objects_ref,
        objects_b,
        worksheet_ref,
        worksheet_b,
        cancel_cb=cancel_cb,
    ):
        if position_ref is not None:
            group_by_ref[position_ref][2] = position_b
        elif position_b is not None:
            unmatched_b[position_b] = objects_b[position_b]

    a_only = {
        group[1]: objects_a[group[1]]
        for group in groups
        if group[0] is None and group[1] is not None
    }
    group_by_a = {
        group[1]: group
        for group in groups
        if group[0] is None and group[1] is not None
    }
    for position_a, position_b in pair_image_positions(
        a_only,
        unmatched_b,
        worksheet_a,
        worksheet_b,
        cancel_cb=cancel_cb,
    ):
        if position_a is not None:
            group_by_a[position_a][2] = position_b
        elif position_b is not None:
            groups.append([None, None, position_b])

    groups.sort(
        key=lambda group: next(
            (position for position in group if position is not None),
            (10**9, 10**9),
        )
    )
    return [(group[0], group[1], group[2]) for group in groups]


def _validated_path(path: str, label: str) -> str:
    normalized = os.path.abspath(os.path.expanduser(str(path).strip()))
    if not os.path.isfile(normalized):
        raise FileNotFoundError(f"{label} Excel 파일을 찾을 수 없습니다: {normalized}")
    if os.path.splitext(normalized)[1].lower() not in {".xlsx", ".xlsm"}:
        raise ValueError(f"{label}에는 .xlsx 또는 .xlsm 파일을 선택해 주세요.")
    return normalized


def load_inspection_images(
    path_ref: str,
    path_a: str,
    path_b: Optional[str] = None,
    *,
    progress_cb: Optional[ProgressCallback] = None,
    status_cb: Optional[StatusCallback] = None,
    cancel_cb: Optional[CancelCallback] = None,
) -> Tuple[Dict[int, List[InspectionItem]], str]:
    """자동 판정 없이, 위치가 맞춰진 이미지 묶음과 임시폴더를 반환한다."""
    paths = [
        _validated_path(path_ref, "Ref"),
        _validated_path(path_a, "비교A"),
    ]
    if path_b:
        paths.append(_validated_path(path_b, "비교B"))

    preview_dir = tempfile.mkdtemp(prefix="excel_image_inspector_")
    workbooks = []
    keep_preview_dir = False

    def raise_if_cancelled() -> None:
        if cancel_cb is not None and cancel_cb():
            raise InterruptedError("사용자 취소")

    try:
        for index, path in enumerate(paths, start=1):
            raise_if_cancelled()
            if status_cb is not None:
                status_cb(f"Excel 파일 여는 중... ({index}/{len(paths)})")
            workbooks.append(load_workbook(path, data_only=True))
        raise_if_cancelled()

        reference_sheets = list(workbooks[0].sheetnames)
        role_names = ("Ref", "비교A", "비교B")
        for workbook_index, workbook in enumerate(workbooks[1:], start=1):
            if list(workbook.sheetnames) != reference_sheets:
                raise ValueError(
                    f"{role_names[workbook_index]}의 시트 탭 구성 또는 순서가 Ref와 다릅니다.\n"
                    f"Ref: {reference_sheets}\n"
                    f"{role_names[workbook_index]}: {list(workbook.sheetnames)}\n"
                    "같은 Excel 양식의 파일을 선택해 주세요."
                )

        sheet_count = max(len(workbook.worksheets) for workbook in workbooks)
        sheet_contexts = []
        total = 0
        for sheet_index in range(1, sheet_count + 1):
            raise_if_cancelled()
            if status_cb is not None:
                status_cb(f"이미지 위치 확인 중... ({sheet_index}/{sheet_count})")
            worksheets = [
                workbook.worksheets[sheet_index - 1]
                if sheet_index <= len(workbook.worksheets)
                else None
                for workbook in workbooks
            ]
            object_maps = [
                extract_image_objects_by_position(worksheet)
                for worksheet in worksheets
            ]
            groups = _align_positions(
                object_maps[0],
                object_maps[1],
                object_maps[2] if len(object_maps) == 3 else None,
                worksheets[0],
                worksheets[1],
                worksheets[2] if len(worksheets) == 3 else None,
                cancel_cb=cancel_cb,
            )
            sheet_contexts.append((sheet_index, worksheets, object_maps, groups))
            total += len(groups)

        output: Dict[int, List[InspectionItem]] = {}
        processed = 0
        for sheet_index, worksheets, object_maps, groups in sheet_contexts:
            raise_if_cancelled()
            sheet_name = next(
                (worksheet.title for worksheet in worksheets if worksheet is not None),
                f"Sheet{sheet_index}",
            )
            if status_cb is not None:
                status_cb(f"이미지 추출 중: {sheet_name}")
            rows: List[InspectionItem] = []
            for row_index, positions in enumerate(groups, start=1):
                raise_if_cancelled()
                canonical = next(
                    (position for position in positions if position is not None),
                    None,
                )
                if canonical is None:
                    continue
                extracted_sides: List[ExtractedImage] = []
                for side_index, side in enumerate(("ref", "a", "b")[: len(workbooks)]):
                    worksheet = worksheets[side_index]
                    position = positions[side_index]
                    image_obj = (
                        object_maps[side_index].get(position)
                        if position is not None
                        else None
                    )
                    side_sheet_name = (
                        worksheet.title
                        if worksheet is not None
                        else f"Sheet{sheet_index}"
                    )
                    if worksheet is not None and image_obj is not None:
                        extracted = extract_image_meta(
                            worksheet, sheet_index, image_obj
                        )
                        persist_source_from_obj(
                            extracted, image_obj, preview_dir, side
                        )
                    else:
                        extracted = make_null_image(
                            sheet_index,
                            side_sheet_name,
                            canonical[0],
                            canonical[1],
                        )
                    extracted_sides.append(extracted)

                rows.append(
                    InspectionItem(
                        sheet_index=sheet_index,
                        index=row_index,
                        image_ref=extracted_sides[0],
                        image_a=extracted_sides[1],
                        image_b=extracted_sides[2]
                        if len(extracted_sides) == 3
                        else None,
                    )
                )
                processed += 1
                if progress_cb is not None:
                    progress_cb(processed, total, rows[-1].cell_address)
                raise_if_cancelled()
            output[sheet_index] = rows

        raise_if_cancelled()
        keep_preview_dir = True
        return output, preview_dir
    finally:
        for workbook in workbooks:
            try:
                workbook.close()
            except Exception:
                # 다른 workbook과 부분 추출 폴더 정리를 계속 수행한다.
                pass
        if not keep_preview_dir:
            shutil.rmtree(preview_dir, ignore_errors=True)
