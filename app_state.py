"""Qt에 의존하지 않는 Double/Triple/Quadra 이미지 탐색 상태."""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence

from models import IntegrityReport, SUPPORTED_MODES, InspectionItem, SheetInfo


@dataclass
class AppState:
    mode: str = "double"
    items_by_sheet: Dict[int, List[InspectionItem]] = field(default_factory=dict)
    current_index: int = -1
    workbook_paths: Dict[str, str] = field(default_factory=dict)
    preview_temp_dir: Optional[str] = None
    sheet_infos: Dict[int, SheetInfo] = field(default_factory=dict)
    integrity_report: Optional[IntegrityReport] = None
    load_warnings: List[str] = field(default_factory=list)
    _flat_items: List[InspectionItem] = field(
        default_factory=list, init=False, repr=False
    )

    @property
    def flat_items(self) -> List[InspectionItem]:
        return self._flat_items

    @property
    def current_item(self) -> Optional[InspectionItem]:
        items = self.flat_items
        if 0 <= self.current_index < len(items):
            return items[self.current_index]
        return None

    def set_items(
        self,
        items_by_sheet: Dict[int, List[InspectionItem]],
        *,
        mode: str,
        workbook_paths: Optional[Dict[str, str]] = None,
        preview_temp_dir: Optional[str] = None,
        sheet_infos: Optional[Dict[int, SheetInfo]] = None,
        integrity_report: Optional[IntegrityReport] = None,
        load_warnings: Optional[Sequence[str]] = None,
    ) -> None:
        if mode not in SUPPORTED_MODES:
            raise ValueError(f"지원하지 않는 검사 모드: {mode}")
        self.mode = mode
        self.items_by_sheet = {
            int(sheet_index): list(items)
            for sheet_index, items in sorted(items_by_sheet.items())
        }
        self.workbook_paths = dict(workbook_paths or {})
        self.preview_temp_dir = preview_temp_dir
        self.sheet_infos = dict(sheet_infos or {})
        self.integrity_report = integrity_report
        self.load_warnings = list(load_warnings or ())
        self._flat_items = [
            item
            for sheet_index in sorted(self.items_by_sheet)
            for item in self.items_by_sheet[sheet_index]
        ]
        self.current_index = 0 if self._flat_items else -1

    def clear(self) -> None:
        self.items_by_sheet = {}
        self.current_index = -1
        self.workbook_paths = {}
        self.preview_temp_dir = None
        self.sheet_infos = {}
        self.integrity_report = None
        self.load_warnings = []
        self._flat_items = []

    def move(self, offset: int) -> bool:
        items = self.flat_items
        if not items:
            return False
        new_index = max(0, min(len(items) - 1, self.current_index + int(offset)))
        changed = new_index != self.current_index
        self.current_index = new_index
        return changed

    def go_first(self) -> bool:
        if not self.flat_items:
            return False
        changed = self.current_index != 0
        self.current_index = 0
        return changed

    def go_last(self) -> bool:
        items = self.flat_items
        if not items:
            return False
        last = len(items) - 1
        changed = self.current_index != last
        self.current_index = last
        return changed

    def select_sheet(self, sheet_index: int) -> bool:
        target = int(sheet_index)
        for index, item in enumerate(self.flat_items):
            if item.sheet_index == target:
                changed = index != self.current_index
                self.current_index = index
                return changed
        return False

    def current_sheet_position(self) -> tuple[int, int]:
        current = self.current_item
        if current is None:
            return (0, 0)
        sheet_items = self.items_by_sheet.get(current.sheet_index, [])
        for index, item in enumerate(sheet_items, start=1):
            if item is current:
                return (index, len(sheet_items))
        return (0, len(sheet_items))
