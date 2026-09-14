"""Excel 시트 자동 매핑과 사용자 지정 매핑을 표현하는 순수 모델."""

from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass, field
import os
from typing import Dict, List, Mapping, Optional, Sequence, Tuple

from excel_manager import read_sheet_index_names
from models import ROLE_LABELS


@dataclass(frozen=True)
class SheetMappingSource:
    """한 Excel 역할에 속한 원본 시트."""

    role: str
    original_index: int
    name: str
    source_sheet_id: str


@dataclass
class SheetMappingGroup:
    """Review 시트 한 행에 배치할 역할별 원본 시트 관계."""

    display_name: str
    role_sheet_ids: Dict[str, str] = field(default_factory=dict)
    matching_warning: str = ""
    enabled: bool = True
    is_custom: bool = False


@dataclass
class SheetMappingPlan:
    """선택된 Excel 전체의 시트 순서와 비교 그룹."""

    roles: Tuple[str, ...]
    sources: Tuple[SheetMappingSource, ...]
    groups: List[SheetMappingGroup]
    warnings: Tuple[str, ...] = ()

    def clone(self) -> "SheetMappingPlan":
        return SheetMappingPlan(
            roles=tuple(self.roles),
            sources=tuple(self.sources),
            groups=[
                SheetMappingGroup(
                    display_name=group.display_name,
                    role_sheet_ids=dict(group.role_sheet_ids),
                    matching_warning=group.matching_warning,
                    enabled=group.enabled,
                    is_custom=group.is_custom,
                )
                for group in self.groups
            ],
            warnings=tuple(self.warnings),
        )


def make_source_sheet_id(role: str, original_index: int, name: str) -> str:
    return f"{role}|sheet:{int(original_index)}|name:{str(name)!r}"


def _normalized_sheet_name(name: str) -> str:
    return str(name).strip().casefold()


def sources_from_sheet_names(
    sheet_names_by_role: Mapping[str, Sequence[Tuple[int, str]]],
    roles: Sequence[str],
) -> Tuple[SheetMappingSource, ...]:
    sources: List[SheetMappingSource] = []
    for role in roles:
        for original_index, name in sheet_names_by_role.get(role, ()):
            normalized_name = str(name)
            sources.append(
                SheetMappingSource(
                    role=role,
                    original_index=int(original_index),
                    name=normalized_name,
                    source_sheet_id=make_source_sheet_id(
                        role, int(original_index), normalized_name
                    ),
                )
            )
    return tuple(sources)


def build_automatic_sheet_mapping(
    sheet_names_by_role: Mapping[str, Sequence[Tuple[int, str]]],
    roles: Sequence[str],
) -> SheetMappingPlan:
    """현재 프로그램 규칙으로 자동 시트 매핑과 합집합 순서를 만든다."""
    selected_roles = tuple(roles)
    sources = sources_from_sheet_names(sheet_names_by_role, selected_roles)
    groups: List[SheetMappingGroup] = []
    exact_groups: Dict[str, SheetMappingGroup] = {}

    # roles와 각 파일 탭 순서대로 처리하므로 REF 순서가 먼저 유지된다.
    for source in sources:
        exact_group = exact_groups.get(source.name)
        if exact_group is not None and source.role not in exact_group.role_sheet_ids:
            exact_group.role_sheet_ids[source.role] = source.source_sheet_id
            continue
        group = SheetMappingGroup(
            display_name=source.name,
            role_sheet_ids={source.role: source.source_sheet_id},
        )
        groups.append(group)
        exact_groups.setdefault(source.name, group)

    source_by_id = {source.source_sheet_id: source for source in sources}
    buckets: Dict[str, List[SheetMappingGroup]] = defaultdict(list)
    for group in groups:
        first_id = next(iter(group.role_sheet_ids.values()))
        buckets[_normalized_sheet_name(source_by_id[first_id].name)].append(group)

    removed_group_ids = set()
    warnings: List[str] = []
    for normalized_name, bucket in buckets.items():
        if len(bucket) <= 1:
            continue
        role_counts = Counter(
            role for group in bucket for role in group.role_sheet_ids
        )
        if all(count == 1 for count in role_counts.values()):
            target = bucket[0]
            for other in bucket[1:]:
                target.role_sheet_ids.update(other.role_sheet_ids)
                removed_group_ids.add(id(other))
            continue

        details = []
        for role in selected_roles:
            names = [
                source_by_id[source_id].name
                for group in bucket
                for source_role, source_id in group.role_sheet_ids.items()
                if source_role == role
            ]
            if names:
                details.append(
                    f"{ROLE_LABELS[role]}: "
                    + ", ".join(repr(name) for name in names)
                )
        warning = (
            f"시트 이름 '{normalized_name}'의 공백/대소문자 보정 결과가 "
            f"모호하여 자동 매칭하지 않았습니다 ({'; '.join(details)})."
        )
        warnings.append(warning)
        for group in bucket:
            group.matching_warning = warning

    plan = SheetMappingPlan(
        roles=selected_roles,
        sources=sources,
        groups=[
            group for group in groups if id(group) not in removed_group_ids
        ],
        warnings=tuple(warnings),
    )
    validate_sheet_mapping_plan(plan, sources, selected_roles)
    return plan


def build_sheet_mapping_from_paths(
    paths_by_role: Mapping[str, str], roles: Sequence[str]
) -> SheetMappingPlan:
    """이미지와 스타일은 열지 않고 파일별 시트명만 읽어 자동 매핑한다."""
    selected_roles = tuple(roles)
    sheet_names_by_role: Dict[str, List[Tuple[int, str]]] = {}
    for role in selected_roles:
        path = os.path.abspath(os.path.expanduser(str(paths_by_role.get(role, "")).strip()))
        if not os.path.isfile(path):
            raise FileNotFoundError(
                f"{ROLE_LABELS[role]} Excel 파일을 찾을 수 없습니다: {path}"
            )
        if os.path.splitext(path)[1].lower() not in {".xlsx", ".xlsm"}:
            raise ValueError(
                f"{ROLE_LABELS[role]}에는 .xlsx 또는 .xlsm 파일을 선택해 주세요."
            )
        sheet_names_by_role[role] = read_sheet_index_names(path)
    return build_automatic_sheet_mapping(sheet_names_by_role, selected_roles)


def sheet_mapping_input_signature(
    paths_by_role: Mapping[str, str], roles: Sequence[str]
) -> Tuple[Tuple[str, str], ...]:
    """선택 역할이나 파일 경로가 바뀌었는지 판별하는 안정적인 키."""
    signature = []
    for role in roles:
        path = os.path.abspath(os.path.expanduser(str(paths_by_role.get(role, "")).strip()))
        signature.append((role, os.path.normcase(path)))
    return tuple(signature)


def validate_sheet_mapping_plan(
    plan: SheetMappingPlan,
    expected_sources: Optional[Sequence[SheetMappingSource]] = None,
    expected_roles: Optional[Sequence[str]] = None,
) -> None:
    """모든 원본 시트가 올바른 역할에 정확히 한 번 배치됐는지 확인한다."""
    roles = tuple(expected_roles) if expected_roles is not None else tuple(plan.roles)
    sources = tuple(expected_sources) if expected_sources is not None else tuple(plan.sources)
    errors: List[str] = []
    if tuple(plan.roles) != roles:
        errors.append("선택 모드의 Excel 역할 구성이 다릅니다.")

    source_by_id = {source.source_sheet_id: source for source in sources}
    if len(source_by_id) != len(sources):
        errors.append("원본 Sheet ID가 중복되었습니다.")

    used_ids: Counter[str] = Counter()
    if plan.groups and not any(group.enabled for group in plan.groups):
        errors.append("활성화된 비교 그룹이 하나도 없습니다.")
    for index, group in enumerate(plan.groups, start=1):
        if not str(group.display_name).strip():
            errors.append(f"{index}번 비교 그룹 이름이 비어 있습니다.")
        if not group.role_sheet_ids:
            errors.append(f"{index}번 비교 그룹에 시트가 없습니다.")
        for role, source_id in group.role_sheet_ids.items():
            if role not in roles:
                errors.append(f"{index}번 비교 그룹에 알 수 없는 역할 {role!r}이 있습니다.")
                continue
            source = source_by_id.get(source_id)
            if source is None:
                errors.append(f"{index}번 비교 그룹에 알 수 없는 Sheet ID가 있습니다: {source_id}")
                continue
            if source.role != role:
                errors.append(
                    f"{index}번 비교 그룹의 {source.name!r} 시트 역할이 "
                    f"{ROLE_LABELS[source.role]}에서 {ROLE_LABELS[role]}로 바뀌었습니다."
                )
            used_ids[source_id] += 1

    expected_ids = Counter(source_by_id.keys())
    missing = expected_ids - used_ids
    duplicate = used_ids - expected_ids
    if missing:
        errors.append("매핑되지 않은 시트가 있습니다: " + ", ".join(missing))
    if duplicate:
        errors.append("두 그룹 이상에 중복된 시트가 있습니다: " + ", ".join(duplicate))
    if errors:
        raise ValueError("시트 매핑이 올바르지 않습니다.\n" + "\n".join(f"- {error}" for error in errors))
