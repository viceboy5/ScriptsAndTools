# theme.py - Dark QSS stylesheet matching the PS1 WPF dark theme
from models import (
    COLOR_BG_DARK, COLOR_BG_PANEL, COLOR_BG_INPUT, COLOR_BG_HEADER,
    COLOR_TEXT_WHITE, COLOR_TEXT_MUTED, COLOR_TEXT_DIM,
    COLOR_GREEN, COLOR_RED, COLOR_AMBER, COLOR_BLUE,
)

DARK_QSS = f"""
/* ── Global ────────────────────────────────────────────────────────── */
QWidget {{
    background-color: {COLOR_BG_DARK};
    color: {COLOR_TEXT_WHITE};
    font-family: Segoe UI, Arial, sans-serif;
    font-size: 12px;
}}

QMainWindow, QDialog {{
    background-color: {COLOR_BG_DARK};
}}

/* ── Scroll area ─────────────────────────────────────────────────────*/
QScrollArea {{
    background-color: #0D0E10;
    border: none;
}}
QScrollArea > QWidget > QWidget {{
    background-color: #0D0E10;
}}

QScrollBar:vertical {{
    background: #1A1C22;
    width: 10px;
    margin: 0;
}}
QScrollBar::handle:vertical {{
    background: #3A3D4A;
    min-height: 20px;
    border-radius: 5px;
}}
QScrollBar::handle:vertical:hover {{
    background: {COLOR_BLUE};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0px;
}}

/* ── Header bar ──────────────────────────────────────────────────────*/
#HeaderBar {{
    background-color: {COLOR_BG_PANEL};
    border-bottom: 1px solid {COLOR_BG_HEADER};
}}

/* ── Buttons ─────────────────────────────────────────────────────────*/
QPushButton {{
    background-color: {COLOR_BG_HEADER};
    color: {COLOR_TEXT_WHITE};
    font-weight: bold;
    border: none;
    border-radius: 4px;
    padding: 5px 12px;
}}
QPushButton:hover {{
    background-color: #3A3D4A;
}}
QPushButton:pressed {{
    background-color: #222430;
}}
QPushButton:disabled {{
    background-color: #333333;
    color: #666666;
}}

/* ── Text inputs ─────────────────────────────────────────────────────*/
QLineEdit, QTextEdit {{
    background-color: {COLOR_BG_INPUT};
    color: {COLOR_TEXT_WHITE};
    border: 1px solid {COLOR_BG_HEADER};
    border-radius: 3px;
    padding: 3px 6px;
    selection-background-color: {COLOR_BLUE};
}}
QLineEdit:focus, QTextEdit:focus {{
    border: 1px solid {COLOR_BLUE};
}}

/* ── ComboBox ────────────────────────────────────────────────────────*/
QComboBox {{
    background-color: {COLOR_BG_INPUT};
    color: {COLOR_TEXT_WHITE};
    border: 1px solid {COLOR_BLUE};
    border-radius: 3px;
    padding: 3px 6px;
    min-width: 80px;
}}
QComboBox:focus {{
    border: 1px solid {COLOR_BLUE};
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
    background-color: {COLOR_BG_INPUT};
    color: {COLOR_TEXT_WHITE};
    selection-background-color: {COLOR_BLUE};
    border: 1px solid {COLOR_BG_HEADER};
    outline: none;
}}

/* ── CheckBox ────────────────────────────────────────────────────────*/
QCheckBox {{
    color: {COLOR_TEXT_WHITE};
    spacing: 6px;
}}
QCheckBox::indicator {{
    width: 14px;
    height: 14px;
    border: 1px solid {COLOR_TEXT_MUTED};
    border-radius: 2px;
    background: {COLOR_BG_INPUT};
}}
QCheckBox::indicator:checked {{
    background: {COLOR_BLUE};
    border-color: {COLOR_BLUE};
}}

/* ── Labels ──────────────────────────────────────────────────────────*/
QLabel {{
    color: {COLOR_TEXT_WHITE};
    background: transparent;
}}

/* ── Group / panel borders ───────────────────────────────────────────*/
QFrame[frameShape="4"],   /* HLine */
QFrame[frameShape="5"] {{ /* VLine */
    color: {COLOR_BG_HEADER};
}}

/* ── Tooltips ────────────────────────────────────────────────────────*/
QToolTip {{
    background-color: {COLOR_BG_PANEL};
    color: {COLOR_TEXT_WHITE};
    border: 1px solid {COLOR_BG_HEADER};
    padding: 4px;
}}

/* ── Message boxes ───────────────────────────────────────────────────*/
QMessageBox {{
    background-color: {COLOR_BG_PANEL};
}}
QMessageBox QLabel {{
    color: {COLOR_TEXT_WHITE};
}}
"""
