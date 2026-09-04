"""Excel 이미지 육안 검사기가 공유하는 순수 데이터 모델."""

from dataclasses import dataclass
from typing import List, Optional, Tuple

from PIL import Image


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

    @property
    def location_text(self) -> str:
        if self.is_null:
            return f"{self.sheet_name} · 이미지 없음"
        merged = f" · 병합 {self.merged_range}" if self.merged_range else ""
        return f"{self.sheet_name} · {self.cell_address}{merged}"


@dataclass
class InspectionItem:
    """같은 위치에서 함께 보여 줄 Ref/A 또는 Ref/A/B 이미지 묶음."""

    sheet_index: int
    index: int
    image_ref: ExtractedImage
    image_a: ExtractedImage
    image_b: Optional[ExtractedImage] = None

    @property
    def sheet_name(self) -> str:
        for image in (self.image_ref, self.image_a, self.image_b):
            if image is not None and not image.is_null and image.sheet_name:
                return image.sheet_name
        for image in (self.image_ref, self.image_a, self.image_b):
            if image is not None and image.sheet_name:
                return image.sheet_name
        return f"Sheet{self.sheet_index}"

    @property
    def cell_address(self) -> str:
        for image in (self.image_ref, self.image_a, self.image_b):
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
        return images
