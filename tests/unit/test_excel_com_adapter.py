"""RMS/IRM detection, Public conversion, and loader routing tests."""

from pathlib import Path
import zipfile

from PIL import Image

import excel_com_adapter
import excel_manager
from excel_com_adapter import (
    DEFAULT_PUBLIC_LABEL,
    OLE_COMPOUND_SIGNATURE,
    ProtectedWorkbookError,
    convert_rms_workbook_to_public,
    is_rms_protected_workbook,
    read_public_label_template,
)
from excel_manager import InspectionWorkbookLoader, persist_source_from_obj
from models import ExtractedImage


def test_rms_detector_requires_ole_signature_and_drm_marker(tmp_path: Path):
    protected = tmp_path / "protected.xlsx"
    protected.write_bytes(
        OLE_COMPOUND_SIGNATURE
        + b"\0" * 32
        + "DRMEncryptedDataSpace".encode("utf-16le")
    )
    zip_like = tmp_path / "normal.xlsx"
    zip_like.write_bytes(
        b"PK\x03\x04" + "DRMEncryptedDataSpace".encode("utf-16le")
    )

    assert is_rms_protected_workbook(str(protected)) is True
    assert is_rms_protected_workbook(str(zip_like)) is False
    assert is_rms_protected_workbook(str(tmp_path / "missing.xlsx")) is False


def test_sheet_name_reader_converts_rms_before_reading_xml(monkeypatch):
    calls = []

    monkeypatch.setattr(excel_manager, "is_rms_protected_workbook", lambda _p: True)
    monkeypatch.setattr(
        excel_manager,
        "convert_rms_workbook_to_public",
        lambda path: calls.append(path),
    )
    monkeypatch.setattr(
        excel_manager,
        "_sheet_names_from_workbook_xml",
        lambda _path: ["ROOM", "HOT"],
    )

    assert excel_manager.read_sheet_index_names("protected.xlsx") == [
        (1, "ROOM"),
        (2, "HOT"),
    ]
    assert calls == ["protected.xlsx"]


def test_loader_reuses_one_excel_session_for_conversions(monkeypatch):
    calls = []

    class FakeSession:
        def close(self):
            calls.append("closed")

    monkeypatch.setattr(excel_manager, "is_rms_protected_workbook", lambda _p: True)
    monkeypatch.setattr(excel_manager, "ExcelComSession", FakeSession)
    monkeypatch.setattr(
        excel_manager,
        "convert_rms_workbook_to_public",
        lambda path, label, session: calls.append((path, label, session)),
    )
    monkeypatch.setattr(
        excel_manager,
        "load_workbook",
        lambda path, data_only: f"workbook:{path}",
    )
    loader = InspectionWorkbookLoader()

    assert loader.open("ref.xlsx") == "workbook:ref.xlsx"
    assert loader.open("a.xlsx") == "workbook:a.xlsx"
    loader.close()

    assert calls[0][0:2] == ("ref.xlsx", DEFAULT_PUBLIC_LABEL)
    assert calls[1][0:2] == ("a.xlsx", DEFAULT_PUBLIC_LABEL)
    assert calls[0][2] is calls[1][2]
    assert calls[2] == "closed"


def test_read_public_label_template_from_ooxml(tmp_path: Path):
    path = tmp_path / "public.xlsx"
    custom_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
 <property name="MSIP_Label_{DEFAULT_PUBLIC_LABEL.label_id}_Enabled"><vt:lpwstr>true</vt:lpwstr></property>
 <property name="MSIP_Label_{DEFAULT_PUBLIC_LABEL.label_id}_Name"><vt:lpwstr>{DEFAULT_PUBLIC_LABEL.label_name}</vt:lpwstr></property>
 <property name="MSIP_Label_{DEFAULT_PUBLIC_LABEL.label_id}_SiteId"><vt:lpwstr>{DEFAULT_PUBLIC_LABEL.site_id}</vt:lpwstr></property>
 <property name="MSIP_Label_{DEFAULT_PUBLIC_LABEL.label_id}_ContentBits"><vt:lpwstr>0</vt:lpwstr></property>
</Properties>"""
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("docProps/custom.xml", custom_xml)

    assert read_public_label_template(str(path)) == DEFAULT_PUBLIC_LABEL


def test_public_conversion_keeps_new_file_and_removes_backup(
    tmp_path: Path, monkeypatch
):
    path = tmp_path / "protected.xlsx"
    path.write_bytes(b"internal")

    class FakeSession:
        def convert_workbook_to_public(self, target, label):
            assert label == DEFAULT_PUBLIC_LABEL
            Path(target).write_bytes(b"public")

    monkeypatch.setattr(
        excel_com_adapter, "is_rms_protected_workbook", lambda _path: True
    )
    monkeypatch.setattr(
        excel_com_adapter,
        "_validate_converted_public_workbook",
        lambda _path, _label: None,
    )

    assert convert_rms_workbook_to_public(
        str(path), session=FakeSession()
    ) is True
    assert path.read_bytes() == b"public"
    assert list(tmp_path.glob("*.internal_backup_*.tmp")) == []


def test_public_conversion_restores_original_after_failure(
    tmp_path: Path, monkeypatch
):
    path = tmp_path / "protected.xlsx"
    original = b"internal-original"
    path.write_bytes(original)

    class FakeSession:
        def convert_workbook_to_public(self, target, _label):
            Path(target).write_bytes(b"damaged")
            raise RuntimeError("save failed")

    monkeypatch.setattr(
        excel_com_adapter, "is_rms_protected_workbook", lambda _path: True
    )

    try:
        convert_rms_workbook_to_public(str(path), session=FakeSession())
    except ProtectedWorkbookError:
        pass
    else:
        raise AssertionError("conversion failure must be reported")

    assert path.read_bytes() == original
    assert list(tmp_path.glob("*.internal_backup_*.tmp")) == []


def test_persist_source_uses_direct_png_export_for_com_image(tmp_path: Path):
    class FakeComImage:
        def export_png(self, path):
            Image.new("RGB", (3, 2), "red").save(path, format="PNG")

    extracted = ExtractedImage(
        sheet_index=1,
        sheet_name="ROOM",
        cell_address="B2",
        merged_range="",
        anchor_row=1,
        anchor_col=1,
    )

    raw = persist_source_from_obj(extracted, FakeComImage(), str(tmp_path), "ref")

    assert raw.startswith(b"\x89PNG\r\n\x1a\n")
    assert extracted.source_path == extracted.preview_path
    assert extracted.source_path.endswith("_ref.png")
    with Image.open(extracted.source_path) as image:
        assert image.size == (3, 2)
