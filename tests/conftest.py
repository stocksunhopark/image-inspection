"""Qt 테스트를 화면 없이 실행하기 위한 공통 설정."""

import os

import pytest
from PyQt6.QtWidgets import QApplication, QMessageBox

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


@pytest.fixture(autouse=True)
def auto_confirm_main_window_exit(monkeypatch):
    """테스트 정리 단계의 MainWindow.close()가 확인창에서 대기하지 않게 한다."""
    monkeypatch.setattr(
        QMessageBox,
        "question",
        lambda *_args, **_kwargs: QMessageBox.StandardButton.Yes,
    )


@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication(["excel-image-inspector-tests", "-platform", "offscreen"])
    return app
