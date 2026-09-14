from pathlib import Path

from PyQt6.QtWidgets import QWidget


_COMBO_ARROW_PATH = (
    Path(__file__).resolve().parent / "assets" / "combo_down_arrow.svg"
).as_posix()


MAIN_WINDOW_STYLESHEET = """
QMainWindow, QWidget {
    background-color: #0b1220;
    color: #d7e3ff;
    font-size: 12px;
}
QLabel#titleLabel {
    color: #67e8f9;
    font-size: 24px;
    font-weight: 700;
}
QLabel#subTitleLabel, QLabel#keyboardHint {
    color: #94a3b8;
}
QLabel#positionLabel {
    color: #e0f2fe;
    font-size: 14px;
    font-weight: 700;
}
QLabel#locationLabel {
    color: #93c5fd;
    font-weight: 600;
}
QLabel#previewPane {
    background-color: #020617;
    border: 1px solid #334155;
    border-radius: 8px;
    color: #64748b;
}
QLabel#previewTitle {
    color: #e0f2fe;
    background-color: #1e293b;
    border: 1px solid #334155;
    border-radius: 6px;
    padding: 5px;
    font-weight: 700;
}
QLabel#imageMeta {
    color: #94a3b8;
    padding: 2px;
}
QLabel#statusBusy { color: #fbbf24; font-weight: 600; }
QLabel#statusIdle { color: #93c5fd; }
QGroupBox {
    background-color: #111a2b;
    border: 1px solid #243246;
    border-radius: 10px;
    margin-top: 10px;
    color: #93c5fd;
    font-weight: 600;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 6px;
}
QLineEdit, QComboBox, QPlainTextEdit {
    background-color: #0f172a;
    border: 1px solid #334155;
    border-radius: 8px;
    padding: 6px;
    color: #e2e8f0;
    selection-background-color: #2563eb;
}
QPlainTextEdit {
    padding: 8px;
    selection-background-color: #2563eb;
}
QComboBox {
    padding: 6px 36px 6px 8px;
}
QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 28px;
    background-color: #162033;
    border: none;
    border-left: 1px solid #334155;
    border-top-right-radius: 7px;
    border-bottom-right-radius: 7px;
}
QComboBox::drop-down:hover {
    background-color: #1e293b;
}
QComboBox::down-arrow {
    image: url("__COMBO_ARROW_PATH__");
    width: 9px;
    height: 6px;
}
QComboBox QAbstractItemView {
    background-color: #0f172a;
    color: #e2e8f0;
    selection-background-color: #1d4ed8;
}
QTableWidget {
    background-color: #0f172a;
    alternate-background-color: #111827;
    border: 1px solid #334155;
    border-radius: 8px;
    color: #e2e8f0;
    gridline-color: #243246;
    selection-background-color: #1d4ed8;
    selection-color: #ffffff;
}
QTableWidget::item { padding: 4px; }
QHeaderView::section {
    background-color: #1e293b;
    color: #7dd3fc;
    border: none;
    border-bottom: 1px solid #334155;
    padding: 6px;
    font-weight: 700;
}
QPushButton {
    background-color: #1d4ed8;
    border: 1px solid #2563eb;
    border-radius: 8px;
    color: white;
    padding: 7px 12px;
    font-weight: 600;
}
QPushButton:hover { background-color: #2563eb; }
QPushButton:pressed { background-color: #1e40af; }
QPushButton:disabled {
    background-color: #334155;
    border-color: #334155;
    color: #94a3b8;
}
QPushButton#memoButton {
    background-color: #334155;
    border-color: #475569;
    padding: 5px 9px;
}
QPushButton#memoButton[hasNote="true"] {
    background-color: #075985;
    border-color: #38bdf8;
    color: #e0f2fe;
}
QPushButton#memoDeleteButton {
    background-color: #7f1d1d;
    border-color: #ef4444;
}
QWidget#reviewPreviewContainer[defectMarked="true"] {
    border: 2px solid #ef4444;
    border-radius: 9px;
}
QWidget#reviewPreviewContainer[defectMarked="true"] QCheckBox#defectCheck {
    color: #fca5a5;
}
QLabel#reviewCountLabel {
    color: #7dd3fc;
    font-weight: 700;
}
QCheckBox {
    spacing: 6px;
    font-weight: 600;
    color: #e0f2fe;
}
QCheckBox::indicator {
    width: 13px;
    height: 13px;
    border: 1px solid #7dd3fc;
    border-radius: 3px;
    background-color: #0f172a;
}
QCheckBox::indicator:hover {
    border-color: #67e8f9;
    background-color: #1e293b;
}
QCheckBox::indicator:checked {
    background-color: #22d3ee;
    border-color: #67e8f9;
}
QCheckBox::indicator:disabled {
    border-color: #475569;
    background-color: #1e293b;
}
QProgressBar {
    background: #0f172a;
    border: 1px solid #334155;
    border-radius: 8px;
    color: #d7e3ff;
    min-height: 18px;
    text-align: center;
}
QProgressBar::chunk { background-color: #1d4ed8; border-radius: 7px; }
QSlider::groove:horizontal {
    height: 6px;
    background: #334155;
    border-radius: 3px;
}
QSlider::handle:horizontal {
    background: #67e8f9;
    width: 16px;
    margin: -5px 0;
    border-radius: 8px;
}
QMenuBar, QMenu { background-color: #111a2b; color: #d7e3ff; }
QMenuBar::item:selected, QMenu::item:selected {
    background-color: #1d4ed8;
    color: white;
}
QSplitter::handle { background-color: #243246; width: 2px; }
""".replace("__COMBO_ARROW_PATH__", _COMBO_ARROW_PATH)


def apply_theme(widget: QWidget) -> None:
    widget.setStyleSheet(MAIN_WINDOW_STYLESHEET)
