"""Excel 이미지 육안 검사기가 공유하는 순수 데이터 모델."""

from dataclasses import dataclass, field
from typing import Dict, Iterator, List, Optional, Tuple

from PIL import Image

SUPPORTED_MODES = frozenset({"double", "triple", "quadra"})
ROLE_ORDER = ("ref", "a", "b", "c")
ROLE_LABELS = {"ref": "REF", "a": "A", "b": "B", "c": "C"}
MISSING_SHEET = "sheet_missing"
MISSING_IMAGE = "image_missing"
MODE_LABELS = {
    "double": "Double",
    "triple": "Triple",
    "quadra": "Quadra",
}


def mode_label(mode: str) -> str:
    return MODE_LABELS.get(mode, mode)


@dataclass
class SheetInfo:
    """통합 Review 시트와 역할별 원본 시트의 존재 관계."""

    sheet_index: int
    display_name: str
    selected_roles: Tuple[str, ...]
    role_sheet_names: Dict[str, Optional[str]] = field(default_factory=dict)
    role_sheet_ids: Dict[str, Optional[str]] = field(default_factory=dict)
    matching_warning: str = ""

    @property
    def present_roles(self) -> Tuple[str, ...]:
        return tuple(
            role
            for role in self.selected_roles
            if self.role_sheet_ids.get(role) is not None
        )

    @property
    def status_text(self) -> str:
        labels = [ROLE_LABELS.get(role, role.upper()) for role in self.present_roles]
        if len(labels) == 1:
            return f"{labels[0]} ONLY"
        return " + ".join(labels)

    @property
    def display_text(self) -> str:
        return f"{self.display_name} [{self.status_text}]"


@dataclass(frozen=True)
class IntegrityReport:
    """Review Queue가 원본 Sheet/Image를 정확히 보존했다는 검사 결과."""

    role_sheet_counts: Tuple[Tuple[str, int], ...]
    review_sheet_count: int
    source_image_count: int
    queue_image_count: int
    review_item_count: int


@dataclass
class InspectionLoadResult:
    """서비스 로딩 결과. 두 값 구조분해와도 호환된다."""

    items_by_sheet: Dict[int, List["InspectionItem"]]
    preview_dir: str
    sheet_infos: Dict[int, SheetInfo]
    integrity_report: IntegrityReport
    warnings: Tuple[str, ...] = ()

    def __iter__(self) -> Iterator[object]:
        # 기존 ``items_by_sheet, preview_dir = ...`` 호출을 유지한다.
        yield self.items_by_sheet
        yield self.preview_dir


@dataclass
class ExtractedImage:
    """워크시트에 삽입된 이미지 한 장과 셀 위치 정보."""

    sheet_index: int
    sheet_name: str
    cell_address: str
    merged_range: str
    anchor_row: int
    anchor_col: int
    image: Optional[Image.Image] = None
    is_null: bool = False
    source_path: Optional[str] = None
    preview_path: Optional[str] = None
    source_role: Optional[str] = None
    source_sheet_id: Optional[str] = None
    source_image_id: Optional[str] = None
    missing_reason: Optional[str] = None

    @property
    def placeholder_text(self) -> str:
        if self.missing_reason == MISSING_SHEET:
            return "시트 없음"
        return "이미지 없음"

    @property
    def location_text(self) -> str:
        if self.is_null:
            return f"{self.sheet_name} · {self.placeholder_text}"
        merged = f" · 병합 {self.merged_range}" if self.merged_range else ""
        return f"{self.sheet_name} · {self.cell_address}{merged}"


@dataclass
class InspectionItem:
    """같은 위치에서 함께 보여 줄 Ref/A, Ref/A/B, Ref/A/B/C 이미지 묶음."""

    sheet_index: int
    index: int
    image_ref: ExtractedImage
    image_a: ExtractedImage
    image_b: Optional[ExtractedImage] = None
    image_c: Optional[ExtractedImage] = None
    sheet_info: Optional[SheetInfo] = None
    is_empty_sheet: bool = False

    def _side_extracted(self) -> Tuple[ExtractedImage, ...]:
        sides = [self.image_ref, self.image_a]
        if self.image_b is not None:
            sides.append(self.image_b)
        if self.image_c is not None:
            sides.append(self.image_c)
        return tuple(sides)

    @property
    def sheet_name(self) -> str:
        if self.sheet_info is not None:
            return self.sheet_info.display_name
        for image in self._side_extracted():
            if image is not None and not image.is_null and image.sheet_name:
                return image.sheet_name
        for image in self._side_extracted():
            if image is not None and image.sheet_name:
                return image.sheet_name
        return f"Sheet{self.sheet_index}"

    @property
    def present_roles(self) -> Tuple[str, ...]:
        if self.sheet_info is not None:
            return self.sheet_info.present_roles
        return tuple(
            role
            for role, image in self.side_images()
            if image.missing_reason != MISSING_SHEET
        )

    @property
    def sheet_status_text(self) -> str:
        if self.sheet_info is not None:
            return self.sheet_info.status_text
        labels = [ROLE_LABELS.get(role, role.upper()) for role in self.present_roles]
        if len(labels) == 1:
            return f"{labels[0]} ONLY"
        return " + ".join(labels)

    @property
    def cell_address(self) -> str:
        if self.is_empty_sheet:
            return "-"
        for image in self._side_extracted():
            if image is not None and not image.is_null:
                return image.cell_address
        return "-"

    def side_images(self) -> List[Tuple[str, ExtractedImage]]:
        images: List[Tuple[str, ExtractedImage]] = [
            ("ref", self.image_ref),
            ("a", self.image_a),
        ]
        if self.image_b is not None:
            images.append(("b", self.image_b))
        if self.image_c is not None:
            images.append(("c", self.image_c))
        return images

    def source_images(self) -> List[Tuple[str, ExtractedImage]]:
        return [
            (role, image)
            for role, image in self.side_images()
            if not image.is_null
        ]
