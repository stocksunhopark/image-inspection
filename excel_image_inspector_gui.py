import os
import sys

os.environ.setdefault("QT_OPENGL", "software")
os.environ.setdefault("QT_QUICK_BACKEND", "software")
os.environ.setdefault("QTWEBENGINE_DISABLE_GPU", "1")
os.environ.setdefault("QT_ANGLE_PLATFORM", "software")

from PyQt6.QtWidgets import QApplication

from ui.main_window import MainWindow


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("Excel Image Inspector")
    app.setApplicationDisplayName("Excel 이미지 육안 검사기")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
