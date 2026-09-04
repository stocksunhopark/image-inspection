"""Double/Triple Excel 로딩을 GUI 스레드 밖에서 실행하는 Qt worker."""

import threading
from typing import Optional

from PyQt6.QtCore import QObject, pyqtSignal

from inspection_service import load_inspection_images


class ExcelLoadWorker(QObject):
    finished = pyqtSignal(object)
    failed = pyqtSignal(str)
    progress = pyqtSignal(int, int, str)
    status = pyqtSignal(str)

    def __init__(self, path_ref: str, path_a: str, path_b: Optional[str] = None):
        super().__init__()
        self.path_ref = path_ref
        self.path_a = path_a
        self.path_b = path_b
        self._stop_event = threading.Event()

    def request_stop(self) -> None:
        self._stop_event.set()

    def is_cancelled(self) -> bool:
        return self._stop_event.is_set()

    def run(self) -> None:
        try:
            items_by_sheet, preview_dir = load_inspection_images(
                self.path_ref,
                self.path_a,
                self.path_b,
                progress_cb=lambda current, total, cell: self.progress.emit(
                    current, total, cell
                ),
                status_cb=self.status.emit,
                cancel_cb=self.is_cancelled,
            )
            mode = "triple" if self.path_b else "double"
            paths = {"ref": self.path_ref, "a": self.path_a}
            if self.path_b:
                paths["b"] = self.path_b
            self.finished.emit((items_by_sheet, preview_dir, mode, paths))
        except InterruptedError:
            self.failed.emit("사용자 취소")
        except Exception as exc:
            self.failed.emit(str(exc))
