"""Excel 시트 합집합을 기반으로 육안 검사용 Review Queue를 만든다."""

from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass, field
import os
import shutil
import tempfile
from typing import Callable, Dict, List, Optional, Sequence, Tuple

from openpyxl import load_workbook

from excel_manager import (
    Position,
    extract_image_meta,
    extract_image_objects_by_position,
    make_null_image,
    pair_image_positions,
    persist_source_from_obj,
)
from models import (
    IntegrityReport,
    InspectionItem,
    InspectionLoadResult,
    MISSING_IMAGE,
    MISSING_SHEET,
    ROLE_LABELS,
    ROLE_ORDER,
    SheetInfo,
)

ProgressCallback = Callable[[int, int, str], None]
StatusCallback = Callable[[str], None]
CancelCallback = Callable[[], bool]
AlignedPositions = Tuple[
    Optional[Position],
    Optional[Position],
    Optional[Position],
    Optional[Position],
]


@dataclass(frozen=True)
class _SourceSheet:
    role: str
    original_index: int
    name: str
    worksheet: object
    source_sheet_id: str


@dataclass
class _SheetGroup:
    sources: Dict[str, _SourceSheet] = field(default_factory=dict)
    matching_warning: str = ""


@dataclass(frozen=True)
class _SourceImageRecord:
    role: str
    source_sheet_id: str
    position: Position


def _attach_extra_side(
    groups: List[List[Optional[Position]]],
    extra_index: int,
    objects_ref: Dict[Position, object],
    objects_extra: Dict[Position, object],
    objects_a: Dict[Position, object],
    worksheet_ref=None,
    worksheet_extra=None,
    worksheet_a=None,
    objects_fallback: Optional[Dict[Position, object]] = None,
    fallback_index: Optional[int] = None,
    worksheet_fallback=None,
    cancel_cb: Optional[CancelCallback] = None,
) -> None:
    """기존 행에 B 또는 C 위치를 붙이고, 남는 위치는 새 행으로 추가한다."""
    group_by_ref = {
        group[0]: group for group in groups if group[0] is not None
    }
    leftover: Dict[Position, object] = {}
    for position_ref, position_extra in pair_image_positions(
        objects_ref,
        objects_extra,
        worksheet_ref,
        worksheet_extra,
        cancel_cb=cancel_cb,
    ):
        if position_ref is not None:
            group_by_ref[position_ref][extra_index] = position_extra
        elif position_extra is not None:
            leftover[position_extra] = objects_extra[position_extra]

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
    still_left: Dict[Position, object] = {}
    for position_a, position_extra in pair_image_positions(
        a_only,
        leftover,
        worksheet_a,
        worksheet_extra,
        cancel_cb=cancel_cb,
    ):
        if position_a is not None:
            group_by_a[position_a][extra_index] = position_extra
        elif position_extra is not None:
            still_left[position_extra] = leftover[position_extra]

    if (
        objects_fallback is not None
        and fallback_index is not None
        and still_left
    ):
        fallback_objects = {
            group[fallback_index]: objects_fallback[group[fallback_index]]
            for group in groups
            if group[fallback_index] is not None
            and group[extra_index] is None
            and all(group[index] is None for index in range(fallback_index))
        }
        group_by_fallback = {
            group[fallback_index]: group
            for group in groups
            if group[fallback_index] is not None
            and group[extra_index] is None
            and all(group[index] is None for index in range(fallback_index))
        }
        remaining: Dict[Position, object] = {}
        for position_fallback, position_extra in pair_image_positions(
            fallback_objects,
            still_left,
            worksheet_fallback,
            worksheet_extra,
            cancel_cb=cancel_cb,
        ):
            if position_fallback is not None:
                group_by_fallback[position_fallback][extra_index] = position_extra
            elif position_extra is not None:
                remaining[position_extra] = still_left[position_extra]
        still_left = remaining

    for position_extra in sorted(still_left):
        group: List[Optional[Position]] = [None, None, None, None]
        group[extra_index] = position_extra
        groups.append(group)


def _align_positions(
    objects_ref: Dict[Position, object],
    objects_a: Dict[Position, object],
    objects_b: Optional[Dict[Position, object]],
    objects_c: Optional[Dict[Position, object]] = None,
    worksheet_ref=None,
    worksheet_a=None,
    worksheet_b=None,
    worksheet_c=None,
    cancel_cb: Optional[CancelCallback] = None,
) -> List[AlignedPositions]:
    """한 통합 시트 안에서 Ref/A/B/C 이미지 위치를 행 단위로 정렬한다."""
    ref_a_pairs = pair_image_positions(
        objects_ref,
        objects_a,
        worksheet_ref,
        worksheet_a,
        cancel_cb=cancel_cb,
    )
    groups: List[List[Optional[Position]]] = [
        [position_ref, position_a, None, None]
        for position_ref, position_a in ref_a_pairs
    ]
    if objects_b is not None:
        _attach_extra_side(
            groups,
            2,
            objects_ref,
            objects_b,
            objects_a,
            worksheet_ref,
            worksheet_b,
            worksheet_a,
            cancel_cb=cancel_cb,
        )
    if objects_c is not None:
        _attach_extra_side(
            groups,
            3,
            objects_ref,
            objects_c,
            objects_a,
            worksheet_ref,
            worksheet_c,
            worksheet_a,
            objects_fallback=objects_b,
            fallback_index=2,
            worksheet_fallback=worksheet_b,
            cancel_cb=cancel_cb,
        )

    groups.sort(
        key=lambda group: next(
            (position for position in group if position is not None),
            (10**9, 10**9),
        )
    )
    return [(group[0], group[1], group[2], group[3]) for group in groups]


def _validated_path(path: str, label: str) -> str:
    normalized = os.path.abspath(os.path.expanduser(str(path).strip()))
    if not os.path.isfile(normalized):
        raise FileNotFoundError(f"{label} Excel 파일을 찾을 수 없습니다: {normalized}")
    if os.path.splitext(normalized)[1].lower() not in {".xlsx", ".xlsm"}:
        raise ValueError(f"{label}에는 .xlsx 또는 .xlsm 파일을 선택해 주세요.")
    return normalized


def _normalized_sheet_name(name: str) -> str:
    """정확 일치가 없을 때만 사용하는 최소 보정 키."""
    return str(name).strip().casefold()


def _make_source_sheet_id(role: str, original_index: int, name: str) -> str:
    return f"{role}|sheet:{original_index}|name:{name!r}"


def _build_sheet_groups(
    workbooks: Sequence[object], roles: Tuple[str, ...]
) -> Tuple[List[_SheetGroup], List[_SourceSheet], Tuple[str, ...]]:
    """정확 이름 우선, 명확한 정규화 이름 보조로 Sheet 합집합을 만든다."""
    groups: List[_SheetGroup] = []
    exact_groups: Dict[str, _SheetGroup] = {}
    source_sheets: List[_SourceSheet] = []

    # 역할 순서대로 발견하므로 Ref 순서가 먼저 유지되고 A/B/C 전용 시트가
    # 각 원본 탭 순서대로 뒤에 추가된다.
    for role, workbook in zip(roles, workbooks):
        for original_index, worksheet in enumerate(workbook.worksheets, start=1):
            source = _SourceSheet(
                role=role,
                original_index=original_index,
                name=worksheet.title,
                worksheet=worksheet,
                source_sheet_id=_make_source_sheet_id(
                    role, original_index, worksheet.title
                ),
            )
            source_sheets.append(source)
            exact_group = exact_groups.get(source.name)
            if exact_group is not None and role not in exact_group.sources:
                exact_group.sources[role] = source
                continue
            group = _SheetGroup(sources={role: source})
            groups.append(group)
            exact_groups.setdefault(source.name, group)

    buckets: Dict[str, List[_SheetGroup]] = defaultdict(list)
    for group in groups:
        first_source = next(iter(group.sources.values()))
        buckets[_normalized_sheet_name(first_source.name)].append(group)

    removed_group_ids = set()
    warnings: List[str] = []
    for normalized_name, bucket in buckets.items():
        if len(bucket) <= 1:
            continue
        role_counts = Counter(
            role for group in bucket for role in group.sources
        )
        if all(count == 1 for count in role_counts.values()):
            target = bucket[0]
            for other in bucket[1:]:
                target.sources.update(other.sources)
                removed_group_ids.add(id(other))
            continue

        details = []
        for role in roles:
            names = [
                source.name
                for group in bucket
                for source_role, source in group.sources.items()
                if source_role == role
            ]
            if names:
                details.append(
                    f"{ROLE_LABELS[role]}: " + ", ".join(repr(name) for name in names)
                )
        warning = (
            f"시트 이름 '{normalized_name}'의 공백/대소문자 보정 결과가 "
            f"모호하여 자동 매칭하지 않았습니다 ({'; '.join(details)})."
        )
        warnings.append(warning)
        for group in bucket:
            group.matching_warning = warning

    return (
        [group for group in groups if id(group) not in removed_group_ids],
        source_sheets,
        tuple(warnings),
    )


def _make_sheet_info(
    sheet_index: int,
    group: _SheetGroup,
    roles: Tuple[str, ...],
) -> SheetInfo:
    display_source = next(group.sources[role] for role in roles if role in group.sources)
    return SheetInfo(
        sheet_index=sheet_index,
        display_name=display_source.name,
        selected_roles=roles,
        role_sheet_names={
            role: group.sources[role].name if role in group.sources else None
            for role in roles
        },
        role_sheet_ids={
            role: (
                group.sources[role].source_sheet_id
                if role in group.sources
                else None
            )
            for role in roles
        },
        matching_warning=group.matching_warning,
    )


def _make_source_image_ids(
    source: Optional[_SourceSheet],
    objects: Dict[Position, object],
) -> Dict[Position, str]:
    if source is None:
        return {}
    return {
        position: (
            f"{source.source_sheet_id}|image:{number}"
            f"|r:{position[0] + 1}|c:{position[1] + 1}"
        )
        for number, position in enumerate(sorted(objects), start=1)
    }


def _difference_text(counter: Counter[str], limit: int = 8) -> str:
    values = list(counter.elements())
    preview = ", ".join(values[:limit])
    if len(values) > limit:
        preview += f" 외 {len(values) - limit}개"
    return preview


def _validate_aligned_sheet(
    info: SheetInfo,
    groups: Sequence[AlignedPositions],
    source_image_ids: Sequence[Dict[Position, str]],
    roles: Tuple[str, ...],
) -> None:
    """원본 bytes를 읽기 전에 위치 정렬의 누락·중복을 먼저 차단한다."""
    errors: List[str] = []
    for side_index, role in enumerate(roles):
        id_by_position = source_image_ids[side_index]
        expected = Counter(id_by_position.values())
        actual = Counter(
            id_by_position[position]
            for positions in groups
            for position in (positions[side_index],)
            if position is not None
        )
        missing = expected - actual
        duplicate = actual - expected
        if missing:
            errors.append(
                f"{info.display_name} {ROLE_LABELS[role]} 누락 Source Image ID: "
                f"{_difference_text(missing)}"
            )
        if duplicate:
            errors.append(
                f"{info.display_name} {ROLE_LABELS[role]} 중복/추가 Source Image ID: "
                f"{_difference_text(duplicate)}"
            )
    if errors:
        raise ValueError(
            "Review Queue 사전 무결성 검사에 실패했습니다. 검토를 시작하지 않습니다.\n"
            + "\n".join(f"- {error}" for error in errors)
        )


def _validate_review_queue(
    output: Dict[int, List[InspectionItem]],
    sheet_infos: Dict[int, SheetInfo],
    source_sheets: Sequence[_SourceSheet],
    source_images: Dict[str, _SourceImageRecord],
    roles: Tuple[str, ...],
) -> IntegrityReport:
    """Sheet 역할 관계와 Source Image ID의 정확히 한 번 배치를 검증한다."""
    errors: List[str] = []
    if set(output) != set(sheet_infos):
        errors.append("원본 통합 Sheet 목록과 Review Queue의 Sheet 목록이 다릅니다.")

    source_sheet_by_id = {source.source_sheet_id: source for source in source_sheets}
    expected_sheet_ids = Counter(source_sheet_by_id.keys())
    queue_sheet_ids: Counter[str] = Counter()

    for sheet_index, info in sheet_infos.items():
        rows = output.get(sheet_index, [])
        if info.selected_roles != roles:
            errors.append(
                f"{sheet_index}. {info.display_name}: 선택 모드의 역할 구성이 다릅니다."
            )
        if not info.present_roles:
            errors.append(
                f"{sheet_index}. {info.display_name}: 대응하는 원본 Sheet가 없습니다."
            )
        if not rows:
            errors.append(f"{sheet_index}. {info.display_name}: Review 항목이 없습니다.")
        for role in roles:
            source_sheet_id = info.role_sheet_ids.get(role)
            source_sheet_name = info.role_sheet_names.get(role)
            if source_sheet_id is None:
                if source_sheet_name is not None:
                    errors.append(
                        f"{sheet_index}. {info.display_name}: {ROLE_LABELS[role]}의 "
                        "Sheet ID 없이 이름만 등록되었습니다."
                    )
                continue
            queue_sheet_ids[source_sheet_id] += 1
            source = source_sheet_by_id.get(source_sheet_id)
            if source is None:
                errors.append(
                    f"{sheet_index}. {info.display_name}: 알 수 없는 원본 Sheet ID "
                    f"{source_sheet_id}"
                )
                continue
            if source.role != role:
                errors.append(
                    f"{sheet_index}. {info.display_name}: {source.name} Sheet가 "
                    f"{ROLE_LABELS[source.role]}에서 {ROLE_LABELS[role]}로 잘못 배치되었습니다."
                )
            if source.name != source_sheet_name:
                errors.append(
                    f"{sheet_index}. {info.display_name}: 원본 Sheet 이름 관계가 다릅니다."
                )
            if _normalized_sheet_name(source.name) != _normalized_sheet_name(
                info.display_name
            ):
                errors.append(
                    f"{sheet_index}. {info.display_name}: 이름이 다른 원본 Sheet "
                    f"{source.name!r}이 연결되었습니다."
                )

    missing_sheets = expected_sheet_ids - queue_sheet_ids
    duplicate_sheets = queue_sheet_ids - expected_sheet_ids
    if missing_sheets:
        errors.append(f"누락 Sheet ID: {_difference_text(missing_sheets)}")
    if duplicate_sheets:
        errors.append(f"중복/추가 Sheet ID: {_difference_text(duplicate_sheets)}")

    expected_image_ids = Counter(source_images.keys())
    queue_image_ids: Counter[str] = Counter()
    for sheet_index, rows in output.items():
        info = sheet_infos.get(sheet_index)
        if info is None:
            continue
        for row in rows:
            if row.sheet_index != sheet_index:
                errors.append(
                    f"{info.display_name}: Review 항목의 Sheet 번호가 잘못되었습니다."
                )
            if row.sheet_info is not info:
                errors.append(
                    f"{info.display_name}: Review 항목의 Sheet 정보 연결이 잘못되었습니다."
                )
            row_roles = tuple(role for role, _image in row.side_images())
            if row_roles != roles:
                errors.append(
                    f"{info.display_name}: Review 역할 열이 {row_roles}로 잘못 구성되었습니다."
                )
            for role, image in row.side_images():
                source_sheet_id = info.role_sheet_ids.get(role)
                if image.source_role != role:
                    errors.append(
                        f"{info.display_name}: placeholder/이미지 역할 정보가 "
                        f"{ROLE_LABELS[role]} 열과 다릅니다."
                    )
                if image.is_null:
                    expected_reason = (
                        MISSING_IMAGE
                        if source_sheet_id is not None
                        else MISSING_SHEET
                    )
                    if image.missing_reason != expected_reason:
                        errors.append(
                            f"{info.display_name} {ROLE_LABELS[role]} {image.cell_address}: "
                            f"placeholder가 {expected_reason}이 아닙니다."
                        )
                    if image.source_image_id is not None:
                        errors.append(
                            f"{info.display_name} {ROLE_LABELS[role]}: placeholder에 "
                            "Source Image ID가 등록되었습니다."
                        )
                    if image.source_sheet_id != source_sheet_id:
                        errors.append(
                            f"{info.display_name} {ROLE_LABELS[role]}: placeholder의 "
                            "Source Sheet 관계가 잘못되었습니다."
                        )
                    continue

                source_image_id = image.source_image_id
                if not source_image_id:
                    errors.append(
                        f"{info.display_name} {ROLE_LABELS[role]} {image.cell_address}: "
                        "Source Image ID가 없습니다."
                    )
                    continue
                queue_image_ids[source_image_id] += 1
                source_record = source_images.get(source_image_id)
                if source_record is None:
                    errors.append(f"알 수 없는 Source Image ID: {source_image_id}")
                    continue
                if source_record.role != role or image.source_role != role:
                    errors.append(
                        f"{source_image_id}: {ROLE_LABELS[source_record.role]} 이미지가 "
                        f"{ROLE_LABELS[role]} 열에 잘못 배치되었습니다."
                    )
                if (
                    source_sheet_id is None
                    or source_record.source_sheet_id != source_sheet_id
                    or image.source_sheet_id != source_sheet_id
                ):
                    errors.append(
                        f"{source_image_id}: 잘못된 통합 Sheet에 배치되었습니다."
                    )
                if source_record.position != (image.anchor_row, image.anchor_col):
                    errors.append(
                        f"{source_image_id}: 원본과 다른 셀 위치에 배치되었습니다."
                    )

    missing_images = expected_image_ids - queue_image_ids
    duplicate_images = queue_image_ids - expected_image_ids
    if missing_images:
        errors.append(f"누락 Source Image ID: {_difference_text(missing_images)}")
    if duplicate_images:
        errors.append(
            f"중복/추가 Source Image ID: {_difference_text(duplicate_images)}"
        )

    if errors:
        details = "\n".join(f"- {error}" for error in errors[:30])
        if len(errors) > 30:
            details += f"\n- 그 외 오류 {len(errors) - 30}개"
        raise ValueError(
            "Review Queue 무결성 검사에 실패했습니다. 검토를 시작하지 않습니다.\n"
            + details
        )

    return IntegrityReport(
        role_sheet_counts=tuple(
            (
                role,
                sum(1 for source in source_sheets if source.role == role),
            )
            for role in roles
        ),
        review_sheet_count=len(sheet_infos),
        source_image_count=sum(expected_image_ids.values()),
        queue_image_count=sum(queue_image_ids.values()),
        review_item_count=sum(len(rows) for rows in output.values()),
    )


def load_inspection_images(
    path_ref: str,
    path_a: str,
    path_b: Optional[str] = None,
    path_c: Optional[str] = None,
    *,
    progress_cb: Optional[ProgressCallback] = None,
    status_cb: Optional[StatusCallback] = None,
    cancel_cb: Optional[CancelCallback] = None,
) -> InspectionLoadResult:
    """모든 역할의 Sheet 합집합과 누락 없는 Review Queue를 반환한다."""
    if path_c and not path_b:
        raise ValueError("비교C를 쓰려면 비교B Excel도 함께 선택해 주세요.")
    paths = [
        _validated_path(path_ref, "Ref"),
        _validated_path(path_a, "비교A"),
    ]
    if path_b:
        paths.append(_validated_path(path_b, "비교B"))
    if path_c:
        paths.append(_validated_path(path_c, "비교C"))
    roles = ROLE_ORDER[: len(paths)]

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

        if status_cb is not None:
            status_cb("시트 이름 기준으로 전체 Review 목록 구성 중...")
        sheet_groups, source_sheets, warnings = _build_sheet_groups(
            workbooks, roles
        )

        sheet_contexts = []
        sheet_infos: Dict[int, SheetInfo] = {}
        source_images: Dict[str, _SourceImageRecord] = {}
        total = 0
        sheet_count = len(sheet_groups)
        for sheet_index, group in enumerate(sheet_groups, start=1):
            raise_if_cancelled()
            info = _make_sheet_info(sheet_index, group, roles)
            sheet_infos[sheet_index] = info
            if status_cb is not None:
                status_cb(
                    f"이미지 위치 확인 중... ({sheet_index}/{sheet_count}) "
                    f"{info.display_name}"
                )
            sources = [group.sources.get(role) for role in roles]
            worksheets = [
                source.worksheet if source is not None else None
                for source in sources
            ]
            object_maps = [
                extract_image_objects_by_position(worksheet)
                for worksheet in worksheets
            ]
            source_image_ids = [
                _make_source_image_ids(source, objects)
                for source, objects in zip(sources, object_maps)
            ]
            for role, source, id_map in zip(roles, sources, source_image_ids):
                if source is None:
                    continue
                for position, source_image_id in id_map.items():
                    if source_image_id in source_images:
                        raise ValueError(
                            f"Source Image ID가 중복 생성되었습니다: {source_image_id}"
                        )
                    source_images[source_image_id] = _SourceImageRecord(
                        role=role,
                        source_sheet_id=source.source_sheet_id,
                        position=position,
                    )

            groups = _align_positions(
                object_maps[0],
                object_maps[1],
                object_maps[2] if len(object_maps) >= 3 else None,
                object_maps[3] if len(object_maps) == 4 else None,
                worksheets[0],
                worksheets[1],
                worksheets[2] if len(worksheets) >= 3 else None,
                worksheets[3] if len(worksheets) == 4 else None,
                cancel_cb=cancel_cb,
            )
            _validate_aligned_sheet(info, groups, source_image_ids, roles)
            sheet_contexts.append(
                (
                    sheet_index,
                    info,
                    sources,
                    worksheets,
                    object_maps,
                    source_image_ids,
                    groups,
                )
            )
            total += max(1, len(groups))

        output: Dict[int, List[InspectionItem]] = {}
        processed = 0
        for (
            sheet_index,
            info,
            sources,
            worksheets,
            object_maps,
            source_image_ids,
            groups,
        ) in sheet_contexts:
            raise_if_cancelled()
            if status_cb is not None:
                status_cb(f"이미지 추출 중: {info.display_name}")
            rows: List[InspectionItem] = []
            position_rows: Sequence[AlignedPositions] = groups or [
                (None, None, None, None)
            ]
            empty_sheet = not groups
            for row_index, positions in enumerate(position_rows, start=1):
                raise_if_cancelled()
                canonical = next(
                    (position for position in positions if position is not None),
                    None,
                )
                extracted_sides = []
                for side_index, role in enumerate(roles):
                    source = sources[side_index]
                    worksheet = worksheets[side_index]
                    position = positions[side_index]
                    image_obj = (
                        object_maps[side_index].get(position)
                        if position is not None
                        else None
                    )
                    if source is not None and image_obj is not None:
                        source_image_id = source_image_ids[side_index][position]
                        extracted = extract_image_meta(
                            worksheet,
                            sheet_index,
                            image_obj,
                            source_role=role,
                            source_sheet_id=source.source_sheet_id,
                            source_image_id=source_image_id,
                        )
                        persist_source_from_obj(
                            extracted, image_obj, preview_dir, role
                        )
                    else:
                        placeholder_position = canonical if not empty_sheet else None
                        extracted = make_null_image(
                            sheet_index,
                            source.name if source is not None else info.display_name,
                            placeholder_position[0]
                            if placeholder_position is not None
                            else None,
                            placeholder_position[1]
                            if placeholder_position is not None
                            else None,
                            source_role=role,
                            source_sheet_id=(
                                source.source_sheet_id
                                if source is not None
                                else None
                            ),
                            missing_reason=(
                                MISSING_IMAGE
                                if source is not None
                                else MISSING_SHEET
                            ),
                        )
                    extracted_sides.append(extracted)

                rows.append(
                    InspectionItem(
                        sheet_index=sheet_index,
                        index=row_index,
                        image_ref=extracted_sides[0],
                        image_a=extracted_sides[1],
                        image_b=(
                            extracted_sides[2]
                            if len(extracted_sides) >= 3
                            else None
                        ),
                        image_c=(
                            extracted_sides[3]
                            if len(extracted_sides) >= 4
                            else None
                        ),
                        sheet_info=info,
                        is_empty_sheet=empty_sheet,
                    )
                )
                processed += 1
                if progress_cb is not None:
                    progress_cb(processed, total, rows[-1].cell_address)
                raise_if_cancelled()
            output[sheet_index] = rows

        raise_if_cancelled()
        if status_cb is not None:
            status_cb("Review Queue 무결성 검사 중...")
        report = _validate_review_queue(
            output,
            sheet_infos,
            source_sheets,
            source_images,
            roles,
        )
        if status_cb is not None:
            status_cb(
                f"무결성 검사 완료: 원본 이미지 {report.source_image_count}개 · "
                "누락 0 · 중복 0"
            )
        keep_preview_dir = True
        return InspectionLoadResult(
            items_by_sheet=output,
            preview_dir=preview_dir,
            sheet_infos=sheet_infos,
            integrity_report=report,
            warnings=warnings,
        )
    finally:
        for workbook in workbooks:
            try:
                workbook.close()
            except Exception:
                # 다른 workbook과 부분 추출 폴더 정리를 계속 수행한다.
                pass
        if not keep_preview_dir:
            shutil.rmtree(preview_dir, ignore_errors=True)
