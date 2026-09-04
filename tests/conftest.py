"""Qt 테스트를 화면 없이 실행하기 위한 공통 설정."""

import os

import pytest
from PyQt6.QtWidgets import QApplication

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication(["excel-image-inspector-tests", "-platform", "offscreen"])
    return app
