# color_slot_widget.py - Single filament color slot widget (status + combo + swatch)
from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QWidget, QHBoxLayout, QVBoxLayout, QLabel, QComboBox, QFrame,
)

from color_library import ColorLibrary
from models import (
    COLOR_BG_INPUT, COLOR_BG_HEADER, COLOR_TEXT_WHITE, COLOR_TEXT_MUTED,
    COLOR_GREEN, COLOR_RED, COLOR_AMBER, COLOR_BLUE,
)


class ColorSlotWidget(QWidget):
    """
    One color slot row: [status label / name combo | colored swatch box with slot number].

    Mirrors the per-slot UI built inside Build-PJob's foreach loop.
    Emits `status_changed` whenever the validity of the selected name changes.
    """

    status_changed = Signal()

    STATUS_MATCHED   = '[MATCHED]'
    STATUS_CHANGED   = '[CHANGED]'
    STATUS_UNMATCHED = '[UNMATCHED]'
    STATUS_VERIFIED  = '[VERIFIED]'

    def __init__(
        self,
        orig_hex: str,
        orig_name: str,
        color_lib: ColorLibrary,
        slot_index: int,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.orig_hex  = orig_hex    # "#RRGGBBFF"
        self.orig_name = orig_name   # matched name at load time (may be "")
        self._color_lib = color_lib
        self._slot_index = slot_index

        self._build_ui()
        # Set initial status
        if orig_name:
            self._set_status(self.STATUS_MATCHED)
        else:
            self._set_status(self.STATUS_UNMATCHED)

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self.setStyleSheet('background:transparent;')
        row = QHBoxLayout(self)
        row.setContentsMargins(0, 0, 0, 15)
        row.setSpacing(6)
        row.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        # Left side: status label + combo
        text_col = QVBoxLayout()
        text_col.setContentsMargins(0, 0, 10, 0)
        text_col.setSpacing(2)
        text_col.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        self._lbl_status = QLabel()
        self._lbl_status.setAlignment(Qt.AlignmentFlag.AlignRight)
        self._lbl_status.setStyleSheet(
            f'color:{COLOR_TEXT_MUTED}; font-size:10px; font-weight:bold; background:transparent;'
        )
        text_col.addWidget(self._lbl_status)

        self._combo = QComboBox()
        self._combo.setMinimumWidth(110)
        self._combo.setMaximumWidth(210)
        self._combo.setStyleSheet(f"""
            QComboBox {{
                background: {COLOR_BG_INPUT};
                color: {COLOR_TEXT_WHITE};
                border: 1px solid {COLOR_BLUE};
                border-radius: 3px;
                padding: 2px 4px;
            }}
            QComboBox::drop-down {{
                border: none;
                width: 20px;
            }}
            QComboBox::down-arrow {{
                image: none;
                width: 8px;
                height: 6px;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 6px solid {COLOR_TEXT_MUTED};
            }}
            QComboBox QAbstractItemView {{
                background: {COLOR_BG_INPUT};
                color: {COLOR_TEXT_WHITE};
                selection-background-color: {COLOR_BLUE};
                border: 1px solid {COLOR_BG_HEADER};
                outline: none;
            }}
        """)
        self._combo.addItem('Select Color...')
        for name in self._color_lib.all_names():
            self._combo.addItem(name)
        if self.orig_name:
            self._combo.setCurrentText(self.orig_name)
        else:
            self._combo.setCurrentIndex(0)
        self._combo.currentTextChanged.connect(self._on_name_changed)
        text_col.addWidget(self._combo)

        row.addLayout(text_col)

        # Right side: colored swatch square with slot number
        swatch_color = self.orig_hex if (self.orig_hex and len(self.orig_hex) >= 7) else '#333333'
        r, g, b = _hex_to_rgb(swatch_color)
        brightness = 0.299 * r + 0.587 * g + 0.114 * b
        num_color = '#000000' if brightness > 128 else '#FFFFFF'

        self._swatch = QFrame()
        self._swatch.setFixedSize(52, 52)
        self._swatch.setStyleSheet(
            f'background:{swatch_color[:7]}; border:1px solid {COLOR_BG_HEADER}; border-radius:2px;'
        )

        self._lbl_num = QLabel(str(self._slot_index))
        self._lbl_num.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._lbl_num.setStyleSheet(
            f'color:{num_color}; font-size:14px; font-weight:bold; background:transparent; border:none;'
        )
        self._lbl_num.setParent(self._swatch)
        self._lbl_num.setGeometry(0, 0, 52, 52)

        row.addWidget(self._swatch)

    # ── Public API ────────────────────────────────────────────────────────────

    def current_name(self) -> str:
        return self._combo.currentText()

    def is_matched(self) -> bool:
        """Return True if the current selection is valid (in library)."""
        return self._lbl_status.text() in (self.STATUS_MATCHED, self.STATUS_CHANGED, self.STATUS_VERIFIED)

    def set_verified(self) -> None:
        self._set_status(self.STATUS_VERIFIED)

    def set_enabled(self, enabled: bool) -> None:
        self._combo.setEnabled(enabled)

    # ── Internal ──────────────────────────────────────────────────────────────

    def _on_name_changed(self, text: str) -> None:
        if self._color_lib.contains_name(text):
            new_hex = self._color_lib.hex_for_name(text) or self.orig_hex
            self._update_swatch(new_hex)
            if text == self.orig_name and self.orig_name:
                self._set_status(self.STATUS_MATCHED)
            else:
                self._set_status(self.STATUS_CHANGED)
        else:
            self._set_status(self.STATUS_UNMATCHED)
        self.status_changed.emit()

    def _set_status(self, status: str) -> None:
        self._lbl_status.setText(status)
        colors = {
            self.STATUS_MATCHED:   COLOR_GREEN,
            self.STATUS_CHANGED:   COLOR_AMBER,
            self.STATUS_UNMATCHED: COLOR_RED,
            self.STATUS_VERIFIED:  COLOR_GREEN,
        }
        color = colors.get(status, COLOR_TEXT_MUTED)
        self._lbl_status.setStyleSheet(
            f'color:{color}; font-size:10px; font-weight:bold; background:transparent;'
        )

    def _update_swatch(self, hex_color: str) -> None:
        r, g, b = _hex_to_rgb(hex_color)
        brightness = 0.299 * r + 0.587 * g + 0.114 * b
        num_color = '#000000' if brightness > 128 else '#FFFFFF'
        self._swatch.setStyleSheet(
            f'background:{hex_color[:7]}; border:1px solid {COLOR_BG_HEADER}; border-radius:2px;'
        )
        self._lbl_num.setStyleSheet(
            f'color:{num_color}; font-size:14px; font-weight:bold; background:transparent; border:none;'
        )


# ── Helpers ───────────────────────────────────────────────────────────────────

def _hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    """Parse '#RRGGBB' or '#RRGGBBFF' to (r, g, b). Returns (51,51,51) on failure."""
    try:
        h = hex_color.lstrip('#')
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    except (ValueError, IndexError):
        return 51, 51, 51
