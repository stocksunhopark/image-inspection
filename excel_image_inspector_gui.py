import ctypes
import os
import sys


_QT_DLL_DIRECTORY = None
_QT_DLL_HANDLES = []


def _prepare_frozen_qt_runtime() -> None:
    """Load bundled Qt DLLs explicitly before importing PyQt6 on Windows."""
    global _QT_DLL_DIRECTORY
    if sys.platform != "win32" or not getattr(sys, "frozen", False):
        return
    bundle_root = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    qt_bin = os.path.join(bundle_root, "PyQt6", "Qt6", "bin")
    if not os.path.isdir(qt_bin):
        return
    _QT_DLL_DIRECTORY = os.add_dll_directory(qt_bin)
    for filename in ("Qt6Core.dll", "Qt6Gui.dll", "Qt6Widgets.dll"):
        dll_path = os.path.join(qt_bin, filename)
        if os.path.isfile(dll_path):
            _QT_DLL_HANDLES.append(ctypes.WinDLL(dll_path))


_prepare_frozen_qt_runtime()

os.environ.setdefault("QT_OPENGL", "software")
os.environ.setdefault("QT_QUICK_BACKEND", "software")
os.environ.setdefault("QTWEBENGINE_DISABLE_GPU", "1")
os.environ.setdefault("QT_ANGLE_PLATFORM", "software")

from PyQt6.QtWidgets import QApplication

from ui.main_window import MainWindow


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("Excel Image Inspector")
    app.setApplicationDisplayName("OSC 파형 수동 비교기")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
