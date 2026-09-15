"""Microsoft Excel COM support for RMS/IRM-protected workbooks.

Protected workbooks are converted in place to the organization's Public label
before the normal OOXML/openpyxl loading path is used.  A same-directory backup
is kept until the converted file has passed structural and label validation.
"""

from __future__ import annotations

import ctypes
from dataclasses import dataclass
from datetime import datetime
import gc
import os
import re
import shutil
import tempfile
import time
from typing import Iterable, List, Optional, Tuple
import uuid
import xml.etree.ElementTree as ET
import zipfile

from PIL import Image, ImageGrab

try:  # Imported eagerly so PyInstaller includes the Windows COM modules.
    import pythoncom
    import win32com.client
except ImportError:  # pragma: no cover - exercised on non-Windows installations.
    pythoncom = None
    win32com = None


OLE_COMPOUND_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_RMS_DATA_SPACE_MARKER = "DRMEncryptedDataSpace".encode("utf-16le")
_PICTURE_SHAPE_TYPES = frozenset({11, 13, 28, 29})
_XL_SCREEN = 1
_XL_BITMAP = 2
_XL_PICTURE = -4147
_MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
_MSO_ASSIGNMENT_METHOD_PRIVILEGED = 1
_PUBLIC_CONVERSION_JUSTIFICATION = (
    "Image Inspection 프로그램에서 이미지를 검사하기 위한 Public 자동 변환"
)
_CUSTOM_PROPERTY_PATTERN = re.compile(
    r"^MSIP_Label_([0-9a-fA-F-]+)_(Enabled|SetDate|Method|Name|SiteId|ActionId|ContentBits)$"
)
_CUSTOM_PROPERTY_NAMESPACE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
)


@dataclass(frozen=True)
class PublicLabelTemplate:
    """Tenant-specific fields needed to apply the Public sensitivity label."""

    label_id: str
    label_name: str
    site_id: str
    content_bits: int = 0


# SMI Public label. ``discover_public_label_template`` refreshes these values
# from a selected Public workbook whenever one is available.
DEFAULT_PUBLIC_LABEL = PublicLabelTemplate(
    label_id="20c036c9-0c34-4148-8d16-ce4edc3b8f9e",
    label_name="20c036c9-0c34-4148-8d16-ce4edc3b8f9e",
    site_id="ef41154a-eda0-4128-a9af-266c62d702f7",
)


class ProtectedWorkbookError(RuntimeError):
    """Raised when Excel cannot authorize or expose a protected workbook."""


def _clipboard_sequence_number() -> Optional[int]:
    try:
        return int(ctypes.windll.user32.GetClipboardSequenceNumber())
    except (AttributeError, OSError):
        return None


def _file_contains(path: str, needle: bytes, chunk_size: int = 64 * 1024) -> bool:
    overlap = max(0, len(needle) - 1)
    tail = b""
    with open(path, "rb") as file_handle:
        while True:
            chunk = file_handle.read(chunk_size)
            if not chunk:
                return False
            combined = tail + chunk
            if needle in combined:
                return True
            tail = combined[-overlap:] if overlap else b""


def is_rms_protected_workbook(path: str) -> bool:
    """Return whether *path* is an Office IRM data-space container."""
    try:
        with open(path, "rb") as file_handle:
            if file_handle.read(len(OLE_COMPOUND_SIGNATURE)) != OLE_COMPOUND_SIGNATURE:
                return False
        return _file_contains(path, _RMS_DATA_SPACE_MARKER)
    except OSError:
        return False


def _custom_property_text(property_element) -> str:
    for child in property_element:
        return child.text or ""
    return ""


def read_public_label_template(path: str) -> Optional[PublicLabelTemplate]:
    """Read an enabled, non-encrypting sensitivity label from an OOXML file."""
    try:
        with zipfile.ZipFile(path) as archive:
            custom_xml = archive.read("docProps/custom.xml")
    except (OSError, KeyError, zipfile.BadZipFile):
        return None

    try:
        root = ET.fromstring(custom_xml)
    except ET.ParseError:
        return None

    labels = {}
    for property_element in root.findall(
        f"{{{_CUSTOM_PROPERTY_NAMESPACE}}}property"
    ):
        match = _CUSTOM_PROPERTY_PATTERN.match(
            property_element.attrib.get("name", "")
        )
        if match is None:
            continue
        label_id, field = match.groups()
        labels.setdefault(label_id.lower(), {})[field] = _custom_property_text(
            property_element
        )

    for label_id, fields in labels.items():
        if fields.get("Enabled", "").strip().casefold() != "true":
            continue
        try:
            content_bits = int(fields.get("ContentBits", "0"))
        except ValueError:
            continue
        if content_bits != 0:
            continue
        return PublicLabelTemplate(
            label_id=label_id,
            label_name=fields.get("Name") or label_id,
            site_id=fields.get("SiteId") or DEFAULT_PUBLIC_LABEL.site_id,
            content_bits=content_bits,
        )
    return None


def discover_public_label_template(paths: Iterable[str]) -> PublicLabelTemplate:
    """Prefer Public label metadata already present in the selected files."""
    candidates = []
    for path in paths:
        template = read_public_label_template(path)
        if template is None:
            continue
        if template.label_id.casefold() == DEFAULT_PUBLIC_LABEL.label_id.casefold():
            return template
        candidates.append(template)
    return candidates[0] if candidates else DEFAULT_PUBLIC_LABEL


def _is_valid_ooxml_workbook(path: str) -> bool:
    try:
        with zipfile.ZipFile(path) as archive:
            archive.read("[Content_Types].xml")
            archive.read("xl/workbook.xml")
        return True
    except (OSError, KeyError, zipfile.BadZipFile):
        return False


def _validate_converted_public_workbook(
    path: str, expected_label: PublicLabelTemplate
) -> None:
    if is_rms_protected_workbook(path) or not _is_valid_ooxml_workbook(path):
        raise ProtectedWorkbookError(
            "Public 변환 후 Excel 파일 구조 검증에 실패했습니다."
        )
    actual_label = read_public_label_template(path)
    if (
        actual_label is None
        or actual_label.label_id.casefold() != expected_label.label_id.casefold()
    ):
        raise ProtectedWorkbookError(
            "Public 변환 후 민감도 레이블 검증에 실패했습니다."
        )


def _make_backup(path: str) -> str:
    directory = os.path.dirname(path) or os.curdir
    basename = os.path.basename(path)
    descriptor, backup_path = tempfile.mkstemp(
        prefix=f".{basename}.internal_backup_",
        suffix=".tmp",
        dir=directory,
    )
    os.close(descriptor)
    try:
        shutil.copy2(path, backup_path)
    except Exception:
        try:
            os.remove(backup_path)
        except OSError:
            pass
        raise
    return backup_path


def convert_rms_workbook_to_public(
    path: str,
    public_label: PublicLabelTemplate = DEFAULT_PUBLIC_LABEL,
    *,
    session: Optional["ExcelComSession"] = None,
) -> bool:
    """Convert an RMS workbook in place and restore the original on failure."""
    normalized = os.path.abspath(path)
    if not is_rms_protected_workbook(normalized):
        return False

    try:
        backup_path = _make_backup(normalized)
    except Exception as exc:
        raise ProtectedWorkbookError(
            "Internal Excel의 안전 백업을 만들 수 없습니다. 파일 쓰기 권한과 "
            "디스크 여유 공간을 확인해 주세요."
        ) from exc

    owns_session = session is None
    conversion_session = session or ExcelComSession()
    try:
        conversion_session.convert_workbook_to_public(normalized, public_label)
        _validate_converted_public_workbook(normalized, public_label)
    except Exception as exc:
        if owns_session:
            conversion_session.close()
        try:
            os.replace(backup_path, normalized)
        except Exception as restore_exc:
            raise ProtectedWorkbookError(
                "Public 변환에 실패했고 원본 자동 복구도 완료하지 못했습니다. "
                f"복구용 파일: {backup_path}"
            ) from restore_exc
        if isinstance(exc, ProtectedWorkbookError):
            raise
        raise ProtectedWorkbookError(
            "Internal Excel을 Public으로 변환하지 못했습니다. 파일이 다른 "
            "프로그램에서 열려 있지 않은지 확인해 주세요."
        ) from exc
    else:
        if owns_session:
            conversion_session.close()
        try:
            os.remove(backup_path)
        except OSError:
            pass
        return True


@dataclass(frozen=True)
class ExcelComMergedRange:
    min_row: int
    min_col: int
    max_row: int
    max_col: int
    address: str

    def __str__(self) -> str:
        return self.address


@dataclass(frozen=True)
class _MergedCells:
    ranges: Tuple[ExcelComMergedRange, ...]


@dataclass(frozen=True)
class _AnchorMarker:
    row: int
    col: int


@dataclass(frozen=True)
class _Anchor:
    _from: _AnchorMarker


class ExcelComImage:
    """Small openpyxl-image-compatible wrapper around an Excel Shape."""

    def __init__(self, worksheet: "ExcelComWorksheet", shape, row: int, col: int):
        self._worksheet = worksheet
        self._shape = shape
        self.anchor = _Anchor(_AnchorMarker(row=row, col=col))

    def export_png(self, path: str) -> None:
        self._worksheet.export_shape_png(self._shape, path)

    def release(self) -> None:
        self._shape = None
        self._worksheet = None


class ExcelComWorksheet:
    """Worksheet surface consumed by the existing inspection service."""

    def __init__(self, workbook: "ExcelComWorkbook", com_worksheet):
        self._workbook = workbook
        self._com_worksheet = com_worksheet
        self.title = str(com_worksheet.Name)
        self._images: List[ExcelComImage] = []
        merged_ranges = {}

        shapes = com_worksheet.Shapes
        for shape_index in range(1, int(shapes.Count) + 1):
            shape = shapes.Item(shape_index)
            try:
                shape_type = int(shape.Type)
                if shape_type not in _PICTURE_SHAPE_TYPES:
                    continue
                cell = shape.TopLeftCell
                row = int(cell.Row) - 1
                col = int(cell.Column) - 1
                merged = self._merged_range_from_cell(cell)
                if merged is not None:
                    merged_ranges[
                        (merged.min_row, merged.min_col, merged.max_row, merged.max_col)
                    ] = merged
                self._images.append(ExcelComImage(self, shape, row, col))
                shape = None  # Ownership moved to ExcelComImage.
            finally:
                if shape is not None:
                    shape = None
        self.merged_cells = _MergedCells(tuple(merged_ranges.values()))

    @staticmethod
    def _merged_range_from_cell(cell) -> Optional[ExcelComMergedRange]:
        try:
            if not bool(cell.MergeCells):
                return None
            area = cell.MergeArea
            min_row = int(area.Row)
            min_col = int(area.Column)
            max_row = min_row + int(area.Rows.Count) - 1
            max_col = min_col + int(area.Columns.Count) - 1
            address = str(area.Address(False, False)).replace("$", "")
            return ExcelComMergedRange(
                min_row=min_row,
                min_col=min_col,
                max_row=max_row,
                max_col=max_col,
                address=address,
            )
        except Exception:
            return None

    def _activate(self) -> None:
        self._workbook._com_workbook.Activate()
        self._com_worksheet.Activate()

    def _copy_picture(self, shape, picture_format: int) -> None:
        last_error = None
        for attempt in range(3):
            try:
                self._activate()
                shape.CopyPicture(_XL_SCREEN, picture_format)
                if pythoncom is not None:
                    pythoncom.PumpWaitingMessages()
                return
            except Exception as exc:
                last_error = exc
                time.sleep(0.08 * (attempt + 1))
        raise ProtectedWorkbookError(
            "보안 Excel의 이미지를 복사할 수 없습니다. 현재 계정에 문서의 "
            "복사/추출 권한이 있는지 확인해 주세요."
        ) from last_error

    def export_shape_png(self, shape, path: str) -> None:
        """Export one picture without saving an unprotected workbook copy."""
        clipboard_sequence = _clipboard_sequence_number()
        self._copy_picture(shape, _XL_BITMAP)
        for _attempt in range(12):
            if pythoncom is not None:
                pythoncom.PumpWaitingMessages()
            if (
                clipboard_sequence is not None
                and _clipboard_sequence_number() == clipboard_sequence
            ):
                time.sleep(0.03)
                continue
            try:
                clipboard_value = ImageGrab.grabclipboard()
            except OSError:
                clipboard_value = None
            if isinstance(clipboard_value, Image.Image):
                try:
                    clipboard_value.convert("RGB").save(path, format="PNG")
                finally:
                    clipboard_value.close()
                return
            time.sleep(0.03)

        # Some Office/clipboard combinations expose only a metafile.  Excel can
        # still rasterize it through a temporary in-memory chart object.
        chart_object = None
        try:
            self._copy_picture(shape, _XL_PICTURE)
            chart_objects = self._com_worksheet.ChartObjects()
            chart_object = chart_objects.Add(
                0, 0, float(shape.Width), float(shape.Height)
            )
            chart_object.Activate()
            chart_object.Chart.Paste()
            exported = bool(chart_object.Chart.Export(path, "PNG"))
            if not exported or not os.path.isfile(path) or os.path.getsize(path) == 0:
                raise ProtectedWorkbookError(
                    "보안 Excel 이미지의 PNG 변환에 실패했습니다."
                )
        finally:
            if chart_object is not None:
                try:
                    chart_object.Delete()
                except Exception:
                    pass

    def release(self) -> None:
        for image in self._images:
            image.release()
        self._images.clear()
        self._com_worksheet = None
        self._workbook = None


class ExcelComWorkbook:
    """Workbook adapter exposing ``worksheets`` and ``close`` like openpyxl."""

    def __init__(self, session: "ExcelComSession", path: str, com_workbook):
        self._session = session
        self._path = path
        self._com_workbook = com_workbook
        self._closed = False
        self.worksheets = [
            ExcelComWorksheet(self, com_workbook.Worksheets(index))
            for index in range(1, int(com_workbook.Worksheets.Count) + 1)
        ]

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        for worksheet in self.worksheets:
            worksheet.release()
        self.worksheets.clear()
        try:
            self._com_workbook.Close(SaveChanges=False)
        finally:
            self._com_workbook = None
            if self._session is not None:
                self._session._forget(self)
            self._session = None


class ExcelComSession:
    """Own one isolated, hidden Excel instance on the calling thread."""

    def __init__(self):
        self._excel = None
        self._workbooks: List[ExcelComWorkbook] = []
        self._com_initialized = False

    def __enter__(self) -> "ExcelComSession":
        self._start()
        return self

    def __exit__(self, _exc_type, _exc, _traceback) -> None:
        self.close()

    def _start(self) -> None:
        if self._excel is not None:
            return
        if pythoncom is None or win32com is None:
            raise ProtectedWorkbookError(
                "보안 Excel 지원 구성요소(pywin32)가 설치되어 있지 않습니다."
            )
        pythoncom.CoInitialize()
        self._com_initialized = True
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = False
            try:
                excel.EnableEvents = False
                excel.AutomationSecurity = _MSO_AUTOMATION_SECURITY_FORCE_DISABLE
            except Exception:
                pass
            self._excel = excel
        except Exception as exc:
            self.close()
            raise ProtectedWorkbookError(
                "보안 Excel을 열려면 Windows용 Microsoft Excel이 설치되어 있어야 합니다."
            ) from exc

    def _open_raw(self, path: str, *, read_only: bool = True):
        self._start()
        try:
            return self._excel.Workbooks.Open(
                os.path.abspath(path),
                UpdateLinks=0,
                ReadOnly=read_only,
                IgnoreReadOnlyRecommended=True,
                Notify=False,
                AddToMru=False,
            )
        except Exception as exc:
            raise ProtectedWorkbookError(
                "보안 Excel을 열 수 없습니다. Microsoft Excel에 회사 계정으로 "
                "로그인한 뒤 이 파일의 읽기 권한을 확인해 주세요."
            ) from exc

    def convert_workbook_to_public(
        self, path: str, public_label: PublicLabelTemplate
    ) -> None:
        """Apply the Public label and save the workbook at its original path."""
        workbook = self._open_raw(path, read_only=False)
        try:
            if bool(workbook.ReadOnly):
                raise ProtectedWorkbookError(
                    "Internal Excel이 읽기 전용으로 열려 Public으로 저장할 수 없습니다. "
                    "다른 Excel 창에서 파일을 닫은 뒤 다시 시도해 주세요."
                )
            sensitivity_label = workbook.SensitivityLabel
            label_info = sensitivity_label.CreateLabelInfo()
            label_info.ActionId = str(uuid.uuid4())
            label_info.AssignmentMethod = _MSO_ASSIGNMENT_METHOD_PRIVILEGED
            label_info.ContentBits = int(public_label.content_bits)
            label_info.IsEnabled = True
            label_info.Justification = _PUBLIC_CONVERSION_JUSTIFICATION
            label_info.LabelId = public_label.label_id
            label_info.LabelName = public_label.label_name
            label_info.SetDate = datetime.now()
            label_info.SiteId = public_label.site_id
            # Office expects the caller context to be a COM object. Reusing the
            # LabelInfo object is the documented VBA/Interop-compatible pattern.
            sensitivity_label.SetLabel(label_info, label_info)
            if pythoncom is not None:
                pythoncom.PumpWaitingMessages()
            workbook.Save()
        except ProtectedWorkbookError:
            raise
        except Exception as exc:
            raise ProtectedWorkbookError(
                "Excel에서 Internal 민감도 레이블을 Public으로 변경하거나 "
                "저장하지 못했습니다."
            ) from exc
        finally:
            workbook.Close(SaveChanges=False)

    def read_sheet_names(self, path: str) -> List[str]:
        workbook = self._open_raw(path)
        try:
            return [
                str(workbook.Worksheets(index).Name)
                for index in range(1, int(workbook.Worksheets.Count) + 1)
            ]
        finally:
            workbook.Close(SaveChanges=False)

    def open_workbook(self, path: str) -> ExcelComWorkbook:
        raw_workbook = self._open_raw(path)
        try:
            workbook = ExcelComWorkbook(self, path, raw_workbook)
        except Exception:
            raw_workbook.Close(SaveChanges=False)
            raise
        self._workbooks.append(workbook)
        return workbook

    def _forget(self, workbook: ExcelComWorkbook) -> None:
        if workbook in self._workbooks:
            self._workbooks.remove(workbook)

    def close(self) -> None:
        for workbook in list(reversed(self._workbooks)):
            try:
                workbook.close()
            except Exception:
                pass
        self._workbooks.clear()
        if self._excel is not None:
            try:
                self._excel.CutCopyMode = False
            except Exception:
                pass
            try:
                self._excel.Quit()
            except Exception:
                pass
            self._excel = None
        gc.collect()
        if self._com_initialized and pythoncom is not None:
            pythoncom.CoUninitialize()
            self._com_initialized = False
