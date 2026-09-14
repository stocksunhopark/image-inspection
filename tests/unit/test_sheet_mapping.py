"""자동/사용자 지정 시트 매핑 모델 테스트."""

from pathlib import Path

from openpyxl import Workbook

from sheet_mapping import (
    SheetMappingGroup,
    build_automatic_sheet_mapping,
    build_sheet_mapping_from_paths,
    sheet_mapping_input_signature,
    validate_sheet_mapping_plan,
)


def test_automatic_mapping_keeps_ref_order_and_shows_exact_and_only_groups():
    plan = build_automatic_sheet_mapping(
        {
            "ref": [(1, "MAIN"), (2, "CLK")],
            "a": [(1, "MAIN"), (2, "Clock")],
            "b": [(1, "MAIN"), (2, "CLK_MAIN")],
        },
        ("ref", "a", "b"),
    )

    assert [group.display_name for group in plan.groups] == [
        "MAIN",
        "CLK",
        "Clock",
        "CLK_MAIN",
    ]
    assert [tuple(group.role_sheet_ids) for group in plan.groups] == [
        ("ref", "a", "b"),
        ("ref",),
        ("a",),
        ("b",),
    ]


def test_automatic_mapping_preserves_existing_unambiguous_name_normalization():
    plan = build_automatic_sheet_mapping(
        {"ref": [(1, "ROOM")], "a": [(1, " room ")]},
        ("ref", "a"),
    )

    assert len(plan.groups) == 1
    assert tuple(plan.groups[0].role_sheet_ids) == ("ref", "a")


def test_custom_mapping_allows_different_names_but_rejects_duplicates():
    plan = build_automatic_sheet_mapping(
        {"ref": [(1, "CLK")], "a": [(1, "Clock")]},
        ("ref", "a"),
    )
    ref_id = plan.groups[0].role_sheet_ids["ref"]
    a_id = plan.groups[1].role_sheet_ids["a"]
    plan.groups = [
        SheetMappingGroup(
            display_name="CLK 비교",
            role_sheet_ids={"ref": ref_id, "a": a_id},
        )
    ]
    validate_sheet_mapping_plan(plan)

    plan.groups.append(
        SheetMappingGroup(display_name="중복", role_sheet_ids={"a": a_id})
    )
    try:
        validate_sheet_mapping_plan(plan)
    except ValueError as exc:
        assert "중복" in str(exc)
    else:
        raise AssertionError("중복 시트 매핑이 허용되었습니다.")


def test_mapping_clone_preserves_activation_and_custom_group_flags():
    plan = build_automatic_sheet_mapping(
        {"ref": [(1, "MAIN")], "a": [(1, "MAIN")]},
        ("ref", "a"),
    )
    plan.groups[0].enabled = False
    plan.groups[0].is_custom = True

    cloned = plan.clone()

    assert cloned.groups[0].enabled is False
    assert cloned.groups[0].is_custom is True
    assert cloned.groups[0] is not plan.groups[0]


def test_mapping_rejects_plan_without_an_active_group():
    plan = build_automatic_sheet_mapping(
        {"ref": [(1, "MAIN")], "a": [(1, "MAIN")]},
        ("ref", "a"),
    )
    plan.groups[0].enabled = False

    try:
        validate_sheet_mapping_plan(plan)
    except ValueError as exc:
        assert "활성화된 비교 그룹" in str(exc)
    else:
        raise AssertionError("활성 비교 그룹이 없는 매핑이 허용되었습니다.")


def test_build_mapping_from_paths_reads_real_excel_sheet_names(tmp_path: Path):
    paths = {}
    for role, names in {"ref": ["MAIN", "CLK"], "a": ["MAIN", "Clock"]}.items():
        workbook = Workbook()
        workbook.active.title = names[0]
        for name in names[1:]:
            workbook.create_sheet(name)
        path = tmp_path / f"{role}.xlsx"
        workbook.save(path)
        workbook.close()
        paths[role] = str(path)

    plan = build_sheet_mapping_from_paths(paths, ("ref", "a"))
    assert [group.display_name for group in plan.groups] == [
        "MAIN",
        "CLK",
        "Clock",
    ]


def test_mapping_signature_survives_file_metadata_changes(tmp_path: Path):
    ref = tmp_path / "ref.xlsx"
    compare = tmp_path / "a.xlsx"
    ref.write_bytes(b"first")
    compare.write_bytes(b"second")
    paths = {"ref": str(ref), "a": str(compare)}
    before = sheet_mapping_input_signature(paths, ("ref", "a"))

    ref.write_bytes(b"changed-size-and-time")
    compare.write_bytes(b"also-changed")
    after = sheet_mapping_input_signature(paths, ("ref", "a"))

    assert after == before
