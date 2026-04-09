# gp_widget.py - GrandparentJob widget (one themed group of card folders)
# Mirrors PS1 Build-GpJob function
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QComboBox, QCheckBox, QFrame, QSizePolicy, QMessageBox,
    QFileDialog,
)

from models import (
    PRINTER_PREFIXES, GP_THEMES,
    COLOR_BG_PANEL, COLOR_BG_HEADER, COLOR_BG_INPUT,
    COLOR_TEXT_WHITE, COLOR_TEXT_MUTED, COLOR_TEXT_DIM,
    COLOR_GREEN, COLOR_RED, COLOR_AMBER, COLOR_BLUE, COLOR_STEEL_BLUE,
)
from color_library import ColorLibrary
from file_utils import detect_printer_prefix


class GpWidget(QFrame):
    """
    Container widget for one grandparent folder (a theme group).

    Layout (mirrors PS1 Build-GpJob):
      ┌─ Header bar ──────────────────────────────────────────────────┐
      │  [CurrentName]  Printer:[▾]  Theme:[▾]  [Skip]  → Preview  (#n)  [CombineTSV] [Remove] │
      └───────────────────────────────────────────────────────────────┘
      ┌─ Parent list (populated by add_parents) ──────────────────────┐
      │  PJobWidget ...                                                │
      └───────────────────────────────────────────────────────────────┘
    """

    def __init__(
        self,
        gp_path: str,
        parent_dict: dict[str, Path],
        color_lib: ColorLibrary,
        script_dir: Path,
        main_window,
    ) -> None:
        super().__init__()
        self.gp_path = gp_path
        self._color_lib = color_lib
        self._script_dir = script_dir
        self._main_window = main_window
        self._p_widgets: list = []   # list[PJobWidget]

        # Derive display name + initial printer prefix + theme
        if gp_path.startswith('ROOT_'):
            self._gp_name = '(No Parent Folder)'
            self._detected_prefix = ''
            self._gp_name_for_theme = ''
        else:
            di = Path(gp_path)
            self._gp_name = di.name
            self._detected_prefix = detect_printer_prefix(self._gp_name)
            parts = self._gp_name.split('_')
            if self._detected_prefix and len(parts) > 1:
                self._gp_name_for_theme = '_'.join(parts[1:])
            else:
                self._gp_name_for_theme = self._gp_name

        self._build_ui()
        self.add_parents(parent_dict)

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self.setStyleSheet(
            f'QFrame {{ background:{COLOR_BG_PANEL}; border:1px solid {COLOR_BG_HEADER}; border-radius:6px; }}'
        )
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        outer.addWidget(self._build_header())

        # Parent list area
        self._parent_list = QWidget()
        self._parent_list.setStyleSheet('background:transparent; border:none;')
        self._parent_layout = QVBoxLayout(self._parent_list)
        self._parent_layout.setContentsMargins(15, 15, 15, 15)
        self._parent_layout.setSpacing(10)
        outer.addWidget(self._parent_list)

    def _build_header(self) -> QWidget:
        header = QWidget()
        header.setStyleSheet(
            f'background:{COLOR_BG_HEADER}; border-radius:0; border:none;'
        )
        header.setFixedHeight(60)

        h = QHBoxLayout(header)
        h.setContentsMargins(15, 0, 15, 0)
        h.setSpacing(0)

        # Current folder name
        lbl_name = QLabel(self._gp_name)
        lbl_name.setStyleSheet(
            f'color:#CCCCCC; font-size:13px; font-weight:bold; background:transparent; border:none;'
        )
        h.addWidget(lbl_name)
        h.addSpacing(20)

        # Printer prefix label + combo
        lbl_pfx = QLabel('Printer: ')
        lbl_pfx.setStyleSheet(f'color:{COLOR_AMBER}; font-size:14px; font-weight:bold; background:transparent; border:none;')
        h.addWidget(lbl_pfx)

        self._cb_prefix = QComboBox()
        self._cb_prefix.setFixedWidth(85)
        self._cb_prefix.addItem('(none)')
        for pfx in PRINTER_PREFIXES:
            self._cb_prefix.addItem(pfx)
        # Select detected prefix
        if self._detected_prefix:
            idx = self._cb_prefix.findText(self._detected_prefix)
            if idx >= 0:
                self._cb_prefix.setCurrentIndex(idx)
        self._cb_prefix.currentIndexChanged.connect(self._on_prefix_changed)
        h.addWidget(self._cb_prefix)
        h.addSpacing(20)

        # Theme label + combo
        lbl_theme = QLabel('Theme: ')
        lbl_theme.setStyleSheet(f'color:{COLOR_AMBER}; font-size:14px; font-weight:bold; background:transparent; border:none;')
        h.addWidget(lbl_theme)

        self._cb_theme = QComboBox()
        self._cb_theme.setFixedWidth(175)
        for t in GP_THEMES:
            self._cb_theme.addItem(t)
        # Match theme from folder name
        matched = next(
            (t for t in GP_THEMES
             if re_strip(t) == re_strip(self._gp_name_for_theme)),
            None,
        )
        if matched:
            self._cb_theme.setCurrentText(matched)
        else:
            self._cb_theme.setCurrentIndex(-1)
        self._cb_theme.currentIndexChanged.connect(self._on_theme_changed)
        h.addWidget(self._cb_theme)
        h.addSpacing(15)

        # Skip rename checkbox
        self._chk_skip = QCheckBox("Don't rename folder")
        self._chk_skip.setStyleSheet(f'color:{COLOR_TEXT_WHITE}; background:transparent; border:none;')
        h.addWidget(self._chk_skip)
        h.addSpacing(20)

        # Live GP folder name preview
        self._lbl_preview = QLabel(self._make_preview_text())
        self._lbl_preview.setStyleSheet(
            f'color:{COLOR_STEEL_BLUE}; font-size:14px; font-weight:bold; background:transparent; border:none;'
        )
        h.addWidget(self._lbl_preview)
        h.addSpacing(15)

        # File count label
        self._lbl_count = QLabel('')
        self._lbl_count.setStyleSheet(f'color:{COLOR_TEXT_MUTED}; font-size:11px; background:transparent; border:none;')
        h.addWidget(self._lbl_count)

        h.addStretch(1)

        # Combine TSV button
        btn_tsv = QPushButton('Combine TSV Data')
        btn_tsv.setFixedSize(140, 30)
        btn_tsv.setStyleSheet('background:#7B4FBF; color:white; font-weight:bold; border:none; border-radius:4px;')
        btn_tsv.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_tsv.clicked.connect(self._on_combine_tsv)
        h.addWidget(btn_tsv)
        h.addSpacing(10)

        # Remove group button
        btn_remove = QPushButton('Remove Group')
        btn_remove.setFixedSize(140, 30)
        btn_remove.setStyleSheet(f'background:{COLOR_RED}; color:white; font-weight:bold; border:none; border-radius:4px;')
        btn_remove.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_remove.clicked.connect(self._on_remove)
        h.addWidget(btn_remove)

        return header

    # ── Public API ────────────────────────────────────────────────────────────

    def add_parents(self, parent_dict: dict[str, Path]) -> None:
        """Add one or more parent folders to this group."""
        from parent_widget import PJobWidget
        for folder_path, anchor_file in parent_dict.items():
            p_widget = PJobWidget(
                folder_path=Path(folder_path),
                anchor_file=anchor_file,
                color_lib=self._color_lib,
                script_dir=self._script_dir,
                gp_widget=self,
                main_window=self._main_window,
            )
            self._parent_layout.addWidget(p_widget)
            self._p_widgets.append(p_widget)
        self._update_count()

    def parent_widgets(self) -> list:
        return list(self._p_widgets)

    def remove_parent_widget(self, p_widget) -> None:
        self._parent_layout.removeWidget(p_widget)
        p_widget.deleteLater()
        if p_widget in self._p_widgets:
            self._p_widgets.remove(p_widget)
        self._update_count()

    def theme(self) -> str:
        """Return the currently selected theme string (alphanumeric only for filename use)."""
        return re_strip(self._cb_theme.currentText() or '')

    def theme_display(self) -> str:
        return self._cb_theme.currentText() or ''

    def prefix(self) -> str:
        txt = self._cb_prefix.currentText()
        return '' if txt == '(none)' else txt

    def skip_rename(self) -> bool:
        return self._chk_skip.isChecked()

    def update_all_previews(self) -> None:
        """Trigger preview refresh on every child PJobWidget."""
        for p_w in self._p_widgets:
            p_w.update_preview()
        self._lbl_preview.setText(self._make_preview_text())

    # ── Internal helpers ──────────────────────────────────────────────────────

    def _update_count(self) -> None:
        n = len(self._p_widgets)
        self._lbl_count.setText(f'({n} plate{"s" if n != 1 else ""})')

    def _make_preview_text(self) -> str:
        pf = self.prefix()
        th = re_strip(self._gp_name_for_theme)
        th_cb = re_strip(self._cb_theme.currentText() or '')
        th = th_cb if th_cb else th
        preview = f'{pf}_{th}' if pf and th else (pf or th)
        return f'\u2192 {preview}' if preview else ''

    def _on_prefix_changed(self) -> None:
        self.update_all_previews()
        self._main_window.update_process_all_button()

    def _on_theme_changed(self) -> None:
        self.update_all_previews()
        self._main_window.update_process_all_button()

    def _on_combine_tsv(self) -> None:
        """Combine all *_Data.tsv files in the grandparent folder."""
        if self.gp_path.startswith('ROOT_') or not Path(self.gp_path).exists():
            QMessageBox.warning(self, 'Combine TSV Data', 'Grandparent folder path is not valid.')
            return

        gp_dir = Path(self.gp_path)
        folder_name = gp_dir.name
        out_tsv = gp_dir / f'{folder_name}_Data.tsv'

        tsv_files = [
            f for f in gp_dir.rglob('*_Data.tsv')
            if not f.name.lower().endswith('_design_data.tsv')
        ]
        if not tsv_files:
            QMessageBox.warning(self, 'Nothing to Combine', f'No TSV data files found in:\n{gp_dir}')
            return

        combined: dict[str, str] = {}
        for tsv in tsv_files:
            if tsv == out_tsv:
                continue
            try:
                lines = tsv.read_text(encoding='utf-8', errors='replace').splitlines()
                if lines:
                    last = lines[-1].strip()
                    if last:
                        key = last.split('\t')[0]
                        combined[key] = last
            except OSError:
                pass

        if not combined:
            QMessageBox.warning(self, 'Nothing to Combine', 'No TSV data rows found to combine.')
            return

        out_tsv.write_text('\n'.join(combined.values()), encoding='utf-8')

        from PySide6.QtGui import QClipboard
        from PySide6.QtWidgets import QApplication
        QApplication.clipboard().setText('\r\n'.join(combined.values()))

        QMessageBox.information(
            self, 'Combine Complete',
            f'Combined {len(combined)} rows into:\n{folder_name}_Data.tsv\n\n'
            f'All {len(combined)} rows have been copied to your clipboard.',
        )

    def _on_remove(self) -> None:
        self._main_window.remove_gp_widget(self)


# ── Utility ───────────────────────────────────────────────────────────────────

def re_strip(text: str) -> str:
    """Strip all non-alphanumeric characters (mirrors PS1 -replace '[^a-zA-Z0-9]','')."""
    import re
    return re.sub(r'[^a-zA-Z0-9]', '', text)
