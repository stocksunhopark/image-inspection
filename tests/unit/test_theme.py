"""공통 테마의 커스텀 컨트롤 렌더링 회귀 테스트."""

from pathlib import Path

from PyQt6.QtGui import QPixmap
from PyQt6.QtWidgets import QComboBox, QStyle, QStyleOptionComboBox

from ui.theme import _COMBO_ARROW_PATH, apply_theme


def test_combo_drop_down_does_not_overlap_rounded_outer_border(qapp):
    combo = QComboBox()
    try:
        apply_theme(combo)
        combo.addItem("18. CLOCK 사용자 비교 [REF + A + B + C]")
        combo.resize(400, 34)
        combo.show()
        qapp.processEvents()

        option = QStyleOptionComboBox()
        combo.initStyleOption(option)
        arrow_rect = combo.style().subControlRect(
            QStyle.ComplexControl.CC_ComboBox,
            option,
            QStyle.SubControl.SC_ComboBoxArrow,
            combo,
        )
        edit_rect = combo.style().subControlRect(
            QStyle.ComplexControl.CC_ComboBox,
            option,
            QStyle.SubControl.SC_ComboBoxEditField,
            combo,
        )

        assert arrow_rect.right() < combo.rect().right()
        assert edit_rect.right() < arrow_rect.left()
        assert Path(_COMBO_ARROW_PATH).is_file()
        assert QPixmap(_COMBO_ARROW_PATH).isNull() is False
    finally:
        combo.close()
        combo.deleteLater()
