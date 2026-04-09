# parent_widget.py - ParentJob widget (one card folder row)
# Mirrors PS1 Build-PJob, Refresh-PJob, Enqueue-PJob, Start-NextProcess
from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QPixmap, QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QWidget, QFrame, QHBoxLayout, QVBoxLayout, QGridLayout,
    QLabel, QPushButton, QLineEdit, QComboBox, QCheckBox,
    QSizePolicy, QMessageBox, QFileDialog, QScrollArea,
)

from models import (
    ADJ_PRESETS,
    COLOR_BG_DARK, COLOR_BG_PANEL, COLOR_BG_INPUT, COLOR_BG_HEADER,
    COLOR_TEXT_WHITE, COLOR_TEXT_MUTED, COLOR_TEXT_DIM,
    COLOR_GREEN, COLOR_RED, COLOR_AMBER, COLOR_BLUE,
    COLOR_PURPLE, COLOR_PINK, COLOR_BLUE_GREY,
)
from models import PJobData, ColorSlotData, FileRowData
from color_library import ColorLibrary
from color_slot_widget import ColorSlotWidget
from file_utils import (
    parse_file, file_base_color, file_sort_key,
    extract_3mf_to_temp, read_3mf_colors, read_3mf_images,
    smart_fill,
)
from image_utils import randomize_pick_colors_from_bytes, bytes_to_qpixmap


# ── FileRowWidget ─────────────────────────────────────────────────────────────

class FileRowWidget(QWidget):
    """
    One row in the file list: [suffix badge] [old name] [-›] [new name] [Open] [X]
    Mirrors PS1 Add-FileRow.
    """

    def __init__(
        self,
        file_path: Path,
        pjob_widget: 'PJobWidget',
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._pjob = pjob_widget
        self._file_path = file_path

        parsed = parse_file(file_path.name)
        self.suffix = parsed['Suffix']
        self.extension = parsed['Extension']
        self.base_color = file_base_color(file_path.name)
        self.old_path = file_path
        self.target_name = ''

        self._build_ui(parsed)

    def _build_ui(self, parsed: dict) -> None:
        self.setStyleSheet(f'background:{COLOR_BG_DARK}; border-bottom:1px solid {COLOR_BG_HEADER};')
        self.setMinimumHeight(40)

        row = QHBoxLayout(self)
        row.setContentsMargins(10, 5, 5, 5)
        row.setSpacing(5)

        # Suffix badge (editable)
        self._suffix_box = QLineEdit(parsed['Suffix'])
        self._suffix_box.setFixedWidth(65)
        self._suffix_box.setStyleSheet(
            f'background:{COLOR_BG_INPUT}; color:{COLOR_AMBER}; border:none; padding:2px 4px;'
        )
        self._suffix_box.textChanged.connect(lambda: self._pjob.update_preview())
        row.addWidget(self._suffix_box)

        # Old name label
        self._lbl_old = QLabel(self._file_path.name)
        self._lbl_old.setStyleSheet(f'color:{COLOR_TEXT_DIM}; font-size:11px; background:transparent;')
        self._lbl_old.setWordWrap(True)
        row.addWidget(self._lbl_old, stretch=1)

        # Arrow
        lbl_arr = QLabel('-›')
        lbl_arr.setStyleSheet(f'color:{COLOR_TEXT_MUTED}; background:transparent;')
        row.addWidget(lbl_arr)

        # New name label (two-tone: base in blue-grey, suffix in type color)
        self._lbl_new_base = QLabel('')
        self._lbl_new_base.setStyleSheet(f'color:{COLOR_BLUE_GREY}; font-size:11px; font-weight:bold; background:transparent;')
        self._lbl_new_sfx = QLabel('')
        self._lbl_new_sfx.setStyleSheet(f'color:{self.base_color}; font-size:11px; font-weight:bold; background:transparent;')

        new_layout = QHBoxLayout()
        new_layout.setContentsMargins(5, 0, 5, 0)
        new_layout.setSpacing(0)
        new_layout.addWidget(self._lbl_new_base)
        new_layout.addWidget(self._lbl_new_sfx)
        new_layout.addStretch()
        row.addLayout(new_layout, stretch=1)

        # Open button
        btn_open = QPushButton('Open')
        btn_open.setFixedSize(40, 20)
        btn_open.setStyleSheet(f'background:{COLOR_BG_HEADER}; color:#A0C4FF; border:none; border-radius:2px;')
        btn_open.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_open.setToolTip(str(self._file_path))
        btn_open.clicked.connect(self._on_open)
        row.addWidget(btn_open)

        # Delete button
        btn_del = QPushButton('X')
        btn_del.setFixedSize(20, 20)
        btn_del.setStyleSheet(f'background:{COLOR_RED}; color:white; border:none; border-radius:2px;')
        btn_del.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_del.clicked.connect(self._on_delete)
        row.addWidget(btn_del)

    def get_suffix(self) -> str:
        return re.sub(r'[^a-zA-Z0-9]', '', self._suffix_box.text())

    def set_new_name(self, base: str, sfx_part: str, collision: bool) -> None:
        if collision:
            full = base + sfx_part
            self._lbl_new_base.setText(full)
            self._lbl_new_base.setStyleSheet(f'color:{COLOR_RED}; font-size:11px; font-weight:bold; background:transparent;')
            self._lbl_new_sfx.setText('')
            self._lbl_old.setStyleSheet(f'color:{COLOR_RED}; font-size:11px; background:transparent;')
        else:
            self._lbl_new_base.setText(base)
            self._lbl_new_base.setStyleSheet(f'color:{COLOR_BLUE_GREY}; font-size:11px; font-weight:bold; background:transparent;')
            self._lbl_new_sfx.setText(sfx_part)
            self._lbl_new_sfx.setStyleSheet(f'color:{self.base_color}; font-size:11px; font-weight:bold; background:transparent;')
            self._lbl_old.setStyleSheet(f'color:{COLOR_TEXT_DIM}; font-size:11px; background:transparent;')

    def _on_open(self) -> None:
        os.startfile(str(self._file_path))

    def _on_delete(self) -> None:
        res = QMessageBox.warning(
            self, 'Confirm',
            f'Permanently delete:\n{self._file_path.name}?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if res == QMessageBox.StandardButton.Yes:
            try:
                self._file_path.unlink()
            except OSError:
                pass
            self._pjob.remove_file_row(self)


# ── PJobWidget ────────────────────────────────────────────────────────────────

class PJobWidget(QFrame):
    """
    One card-folder row.  Mirrors PS1 Build-PJob + all associated handlers.

    Layout:
      [Left: card image (438) + pick image (438)] | [Right: controls (560px fixed)]
    """

    _CARD_SIZE = 438

    def __init__(
        self,
        folder_path: Path,
        anchor_file: Path,
        color_lib: ColorLibrary,
        script_dir: Path,
        gp_widget,
        main_window,
    ) -> None:
        super().__init__()
        self._color_lib = color_lib
        self._script_dir = script_dir
        self._gp_widget = gp_widget
        self._main_window = main_window

        # Data
        self.data = PJobData(
            folder_path=folder_path,
            anchor_file=anchor_file,
            temp_work=Path(),   # set after extraction
        )
        self._file_row_widgets: list[FileRowWidget] = []
        self._color_slot_widgets: list[ColorSlotWidget] = []
        self._active_pick_path: str | None = None  # for double-click merge map

        # Extract 3MF to temp
        self._extract_3mf()

        self.setStyleSheet(
            f'QFrame {{ background:{COLOR_BG_DARK}; border:1px solid {COLOR_BG_HEADER}; border-radius:4px; }}'
        )
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)

        self._build_ui()
        self.update_preview()

    # ── 3MF extraction ────────────────────────────────────────────────────────

    def _extract_3mf(self) -> None:
        try:
            temp_work = extract_3mf_to_temp(self.data.anchor_file)
            self.data.temp_work = temp_work
            color_slots = read_3mf_colors(temp_work, self._color_lib.hex_to_name)
            self.data.color_slots = color_slots
        except Exception:
            self.data.temp_work = Path(tempfile.mkdtemp(prefix='LiveCard_'))
            self.data.color_slots = []

    # ── UI Construction ───────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        outer = QHBoxLayout(self)
        outer.setContentsMargins(10, 10, 10, 10)
        outer.setSpacing(0)

        # LEFT: images (scaling, star-sized)
        left_widget = QWidget()
        left_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        left_layout = QHBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 15, 0)
        left_layout.setSpacing(0)
        left_layout.addWidget(self._build_card_panel())
        left_layout.addWidget(self._build_pick_panel())
        outer.addWidget(left_widget, stretch=1)

        # RIGHT: controls (fixed 560px)
        right_widget = QWidget()
        right_widget.setFixedWidth(560)
        right_widget.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Maximum)
        self._right_layout = QVBoxLayout(right_widget)
        self._right_layout.setContentsMargins(15, 0, 0, 0)
        self._right_layout.setSpacing(8)
        self._build_right_panel()
        outer.addWidget(right_widget)

    # ── Card image panel (left column 1) ──────────────────────────────────────

    def _build_card_panel(self) -> QWidget:
        S = self._CARD_SIZE
        panel = QWidget()
        panel.setFixedSize(S, S)
        panel.setAcceptDrops(True)
        panel.dragEnterEvent = self._card_drag_enter
        panel.dropEvent = self._card_drop

        # Stacked layout: everything at (0,0)
        layout = QGridLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Main plate image
        self._img_plate = QLabel()
        self._img_plate.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._img_plate.setScaledContents(True)
        self._img_plate.setCursor(Qt.CursorShape.PointingHandCursor)
        self._img_plate.mouseDoubleClickEvent = self._on_plate_dbl_click
        layout.addWidget(self._img_plate, 0, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)

        # [CURRENT] gcode thumbnail (top-left corner, 110x125)
        self._current_thumb = QWidget()
        self._current_thumb.setFixedSize(110, 125)
        ct_layout = QGridLayout(self._current_thumb)
        ct_layout.setContentsMargins(0, 0, 0, 0)
        self._img_current = QLabel()
        self._img_current.setScaledContents(True)
        ct_layout.addWidget(self._img_current, 0, 0)
        lbl_ct = QLabel('[CURRENT]')
        lbl_ct.setStyleSheet(f'color:{COLOR_AMBER}; font-size:8px; font-weight:bold; background:transparent;')
        ct_layout.addWidget(lbl_ct, 0, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self._current_thumb.setVisible(False)
        layout.addWidget(self._current_thumb, 0, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)

        # Character name label (top-right)
        self._lbl_char_card = QLabel('')
        self._lbl_char_card.setStyleSheet(
            f'color:{COLOR_AMBER}; font-size:20px; font-weight:bold; background:transparent;'
        )
        self._lbl_char_card.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
        layout.addWidget(self._lbl_char_card, 0, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)

        # Browse Image button (bottom-left)
        self._btn_browse_img = QPushButton('Browse Images')
        self._btn_browse_img.setFixedSize(110, 22)
        self._btn_browse_img.setStyleSheet(
            f'background:{COLOR_AMBER}; color:white; font-weight:bold; border:none; border-radius:2px;'
        )
        self._btn_browse_img.setCursor(Qt.CursorShape.PointingHandCursor)
        self._btn_browse_img.clicked.connect(self._on_browse_image)
        layout.addWidget(self._btn_browse_img, 0, 0, Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignLeft)

        # Color slots overlay (right-center)
        self._color_slots_container = QWidget()
        cs_layout = QVBoxLayout(self._color_slots_container)
        cs_layout.setContentsMargins(0, 0, 10, 0)
        cs_layout.setSpacing(0)
        cs_layout.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        for i, slot_data in enumerate(self.data.color_slots, start=1):
            slot_widget = ColorSlotWidget(
                orig_hex=slot_data.orig_hex,
                orig_name=slot_data.name,
                color_lib=self._color_lib,
                slot_index=i,
            )
            slot_widget.status_changed.connect(self._on_color_status_changed)
            cs_layout.addWidget(slot_widget)
            self._color_slot_widgets.append(slot_widget)
        layout.addWidget(self._color_slots_container, 0, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        # Processing overlay (amber border + semi-transparent amber fill)
        self._proc_overlay = QFrame()
        self._proc_overlay.setStyleSheet(
            'background:rgba(232,161,53,30); border:6px solid rgba(232,161,53,220); border-radius:0;'
        )
        self._proc_overlay.setVisible(False)

        self._lbl_card_status = QLabel('[ PROCESSING ]')
        self._lbl_card_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._lbl_card_status.setStyleSheet(
            'color:rgba(232,161,53,255); font-size:13px; font-weight:bold;'
            'background:rgba(0,0,0,180); padding:4px 5px 6px 5px; border:none;'
        )
        inner_v = QVBoxLayout(self._proc_overlay)
        inner_v.setContentsMargins(0, 0, 0, 0)
        inner_v.addStretch()
        inner_v.addWidget(self._lbl_card_status)
        layout.addWidget(self._proc_overlay, 0, 0)

        # Finished overlay (dark cover over both panels - we only add it to card panel for now)
        self._finished_overlay = QFrame()
        self._finished_overlay.setStyleSheet('background:rgba(16,17,23,230); border:none;')
        self._finished_overlay.setVisible(False)
        self._img_plate_finished = QLabel()
        self._img_plate_finished.setScaledContents(True)
        self._img_plate_finished.setCursor(Qt.CursorShape.PointingHandCursor)
        self._img_plate_finished.mouseDoubleClickEvent = self._on_finished_dbl_click
        fin_layout = QVBoxLayout(self._finished_overlay)
        fin_layout.setContentsMargins(0, 0, 0, 0)
        fin_layout.addWidget(self._img_plate_finished)
        layout.addWidget(self._finished_overlay, 0, 0)

        # Load initial image
        self._load_initial_plate_image()

        return panel

    # ── Pick image panel (left column 2) ──────────────────────────────────────

    def _build_pick_panel(self) -> QWidget:
        S = self._CARD_SIZE
        panel = QWidget()
        panel.setFixedSize(S, S)

        layout = QGridLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Pick layer image
        self._img_pick = QLabel()
        self._img_pick.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._img_pick.setScaledContents(True)
        self._img_pick.setCursor(Qt.CursorShape.PointingHandCursor)
        self._img_pick.mouseDoubleClickEvent = self._on_pick_dbl_click
        layout.addWidget(self._img_pick, 0, 0)

        # Merge detected banner (top of pick panel)
        nest_exists = any(
            p.name.lower().endswith('nest.3mf')
            for p in self.data.folder_path.iterdir()
            if p.is_file()
        ) if self.data.folder_path.exists() else False

        self._merge_banner = QLabel('MERGE DETECTED')
        self._merge_banner.setAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignTop)
        self._merge_banner.setStyleSheet(
            'color:#FFFFFF; font-size:12px; font-weight:bold;'
            'background:rgba(30,140,60,210); padding:5px; border:none;'
        )
        self._merge_banner.setVisible(nest_exists)
        layout.addWidget(self._merge_banner, 0, 0, Qt.AlignmentFlag.AlignTop)

        # Pick processing overlay
        self._pick_proc_overlay = QFrame()
        self._pick_proc_overlay.setStyleSheet(
            'background:rgba(232,161,53,30); border:6px solid rgba(232,161,53,220); border-radius:0;'
        )
        self._pick_proc_overlay.setVisible(False)

        self._lbl_pick_status = QLabel('[ PROCESSING ]')
        self._lbl_pick_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._lbl_pick_status.setStyleSheet(
            'color:rgba(232,161,53,255); font-size:13px; font-weight:bold;'
            'background:rgba(0,0,0,180); padding:4px 5px 6px 5px; border:none;'
        )
        inner_v = QVBoxLayout(self._pick_proc_overlay)
        inner_v.setContentsMargins(0, 0, 0, 0)
        inner_v.addStretch()
        inner_v.addWidget(self._lbl_pick_status)
        layout.addWidget(self._pick_proc_overlay, 0, 0)

        # Load pick image from existing gcode file
        self._load_pick_image()

        return panel

    # ── Right panel ───────────────────────────────────────────────────────────

    def _build_right_panel(self) -> None:
        rl = self._right_layout

        # Header: folder label + Refresh + Remove Folder buttons
        hdr = QHBoxLayout()
        self._lbl_folder = QLabel(f'Folder: {self.data.folder_path.name}')
        self._lbl_folder.setStyleSheet(f'color:{COLOR_TEXT_WHITE}; font-size:14px; font-weight:bold; background:transparent;')
        hdr.addWidget(self._lbl_folder, stretch=1)

        btn_refresh = QPushButton('Refresh')
        btn_refresh.setFixedSize(100, 25)
        btn_refresh.setStyleSheet(f'background:{COLOR_BG_HEADER}; color:white; border:none; border-radius:3px;')
        btn_refresh.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_refresh.clicked.connect(self._on_refresh)
        hdr.addWidget(btn_refresh)

        btn_remove_folder = QPushButton('Remove Folder')
        btn_remove_folder.setFixedSize(110, 25)
        btn_remove_folder.setStyleSheet(f'background:{COLOR_RED}; color:white; font-weight:bold; border:none; border-radius:3px;')
        btn_remove_folder.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_remove_folder.clicked.connect(self._on_remove_folder)
        hdr.addWidget(btn_remove_folder)
        rl.addLayout(hdr)

        # Revert merge button (hidden unless merge exists)
        self._btn_revert_merge = QPushButton('Revert Merge')
        self._btn_revert_merge.setFixedSize(120, 25)
        self._btn_revert_merge.setStyleSheet(f'background:{COLOR_AMBER}; color:white; font-weight:bold; border:none; border-radius:3px;')
        self._btn_revert_merge.setCursor(Qt.CursorShape.PointingHandCursor)
        self._btn_revert_merge.clicked.connect(self._on_revert_merge)
        nest_exists = any(
            p.name.lower().endswith('nest.3mf')
            for p in self.data.folder_path.iterdir()
            if p.is_file()
        ) if self.data.folder_path.exists() else False
        self._btn_revert_merge.setVisible(nest_exists)
        rl.addWidget(self._btn_revert_merge, alignment=Qt.AlignmentFlag.AlignRight)

        # Character + Adjective edit box
        edit_frame = QFrame()
        edit_frame.setStyleSheet(
            f'background:{COLOR_BG_PANEL}; border:1px solid {COLOR_BG_HEADER}; border-radius:3px;'
        )
        edit_h = QHBoxLayout(edit_frame)
        edit_h.setContentsMargins(10, 10, 10, 10)
        edit_h.setSpacing(20)

        # Character
        char_col = QVBoxLayout()
        char_col.addWidget(self._muted_label('Character *'))
        gp_name = self._gp_widget._gp_name if hasattr(self._gp_widget, '_gp_name') else ''
        fills = smart_fill(self.data.anchor_file.name, gp_name)
        self._tb_char = QLineEdit(fills['Char'])
        self._tb_char.setFixedWidth(200)
        self._tb_char.setStyleSheet(f'background:{COLOR_BG_INPUT}; color:white; border:1px solid {COLOR_BG_HEADER}; padding:3px 6px; border-radius:3px;')
        self._tb_char.textChanged.connect(self.update_preview)
        char_col.addWidget(self._tb_char)
        edit_h.addLayout(char_col)

        # Adjective
        adj_col = QVBoxLayout()
        adj_col.addWidget(self._muted_label('Adjective (Optional)'))
        self._cb_adj = QComboBox()
        self._cb_adj.setEditable(True)
        self._cb_adj.setFixedWidth(200)
        self._cb_adj.addItem('')
        for adj in ADJ_PRESETS:
            self._cb_adj.addItem(adj)
        self._cb_adj.setCurrentText(fills['Adj'])
        self._cb_adj.currentTextChanged.connect(self.update_preview)
        adj_col.addWidget(self._cb_adj)
        edit_h.addLayout(adj_col)
        edit_h.addStretch()
        rl.addWidget(edit_frame)

        # Processing options checkboxes
        opts_frame = QFrame()
        opts_frame.setStyleSheet(
            f'background:{COLOR_BG_PANEL}; border:1px solid {COLOR_BG_HEADER}; border-radius:3px;'
        )
        opts_h = QHBoxLayout(opts_frame)
        opts_h.setContentsMargins(10, 8, 10, 8)
        opts_h.setSpacing(15)

        self._chk_merge   = self._make_checkbox('Merge',   True)
        self._chk_slice   = self._make_checkbox('Slice',   True)
        self._chk_extract = self._make_checkbox('Extract', True)
        self._chk_image   = self._make_checkbox('Image',   False)
        self._chk_logs    = self._make_checkbox('Logs',    False)

        for chk in (self._chk_merge, self._chk_slice, self._chk_extract, self._chk_image, self._chk_logs):
            opts_h.addWidget(chk)
        opts_h.addStretch()
        rl.addWidget(opts_frame)

        # File list
        self._files_widget = QWidget()
        self._files_widget.setStyleSheet('background:transparent; border:none;')
        self._files_layout = QVBoxLayout(self._files_widget)
        self._files_layout.setContentsMargins(0, 5, 0, 5)
        self._files_layout.setSpacing(0)
        rl.addWidget(self._files_widget)

        self._populate_file_rows()

        # Action buttons
        action_row = QHBoxLayout()
        action_row.setAlignment(Qt.AlignmentFlag.AlignRight)

        self._btn_delete_logs = QPushButton('Delete Logs')
        self._btn_delete_logs.setFixedSize(100, 35)
        self._btn_delete_logs.setStyleSheet(f'background:#555555; color:white; font-weight:bold; border:none; border-radius:4px;')
        self._btn_delete_logs.setCursor(Qt.CursorShape.PointingHandCursor)
        self._btn_delete_logs.clicked.connect(self._on_delete_logs)
        action_row.addWidget(self._btn_delete_logs)
        action_row.addSpacing(15)

        self._btn_apply = QPushButton('Add to Queue')
        self._btn_apply.setFixedSize(150, 35)
        self._btn_apply.setStyleSheet(f'background:{COLOR_GREEN}; color:white; font-weight:bold; border:none; border-radius:4px;')
        self._btn_apply.setCursor(Qt.CursorShape.PointingHandCursor)
        self._btn_apply.clicked.connect(self._on_apply)
        action_row.addWidget(self._btn_apply)

        self._btn_revert_done = QPushButton('REVERT')
        self._btn_revert_done.setFixedSize(75, 35)
        self._btn_revert_done.setStyleSheet(f'background:{COLOR_RED}; color:white; font-weight:bold; border:none; border-radius:4px;')
        self._btn_revert_done.setCursor(Qt.CursorShape.PointingHandCursor)
        self._btn_revert_done.setVisible(False)
        self._btn_revert_done.clicked.connect(lambda: self._on_revert_done())
        action_row.addWidget(self._btn_revert_done)

        rl.addLayout(action_row)
        rl.addStretch()

    # ── Image loading ─────────────────────────────────────────────────────────

    def _load_initial_plate_image(self) -> None:
        # Prefer custom PNG in folder
        if self.data.folder_path.exists():
            pngs = sorted(
                [p for p in self.data.folder_path.glob('*.png')
                 if not p.name.lower().endswith('_slicepreview.png')],
                key=lambda p: p.stat().st_mtime, reverse=True,
            )
            if pngs:
                self.data.custom_image_path = pngs[0]
                pm = QPixmap(str(pngs[0]))
                if not pm.isNull():
                    self._img_plate.setPixmap(pm)
                    self._btn_browse_img.setStyleSheet(
                        f'background:{COLOR_GREEN}; color:white; font-weight:bold; border:none; border-radius:2px;'
                    )
                    return

        # Fallback: plate_1.png / thumbnail.png from temp work
        for name in ('plate_1.png', 'thumbnail.png'):
            candidate = self.data.temp_work / 'Metadata' / name
            if candidate.exists():
                pm = QPixmap(str(candidate))
                if not pm.isNull():
                    self._img_plate.setPixmap(pm)
                    return

    def _load_pick_image(self) -> None:
        # Load pick from existing *Full.gcode.3mf in folder
        if not self.data.folder_path.exists():
            return
        gcode_files = sorted(
            [p for p in self.data.folder_path.glob('*Full.gcode.3mf')],
            key=lambda p: p.stat().st_mtime, reverse=True,
        )
        if not gcode_files:
            return
        gcode_file = gcode_files[0]
        try:
            with zipfile.ZipFile(gcode_file, 'r') as zf:
                names = {e.replace('\\', '/').lower(): e for e in zf.namelist()}
                # Current plate image
                for cand in ('metadata/plate_1.png', 'metadata/thumbnail.png'):
                    if cand in names:
                        raw = zf.read(names[cand])
                        pm = bytes_to_qpixmap(raw)
                        if pm and not pm.isNull():
                            self._img_current.setPixmap(pm)
                            self._current_thumb.setVisible(True)
                        break
                # Pick image
                for norm, orig in names.items():
                    if norm.endswith('pick_1.png'):
                        raw = zf.read(orig)
                        colored = randomize_pick_colors_from_bytes(raw) or raw
                        pm = bytes_to_qpixmap(colored)
                        if pm and not pm.isNull():
                            self._img_pick.setPixmap(pm)
                            self._active_pick_path = str(
                                self.data.temp_work / 'pick_1_display.png'
                            )
                            # Save for merge-map lookup
                            try:
                                (self.data.temp_work / 'pick_1_display.png').write_bytes(colored)
                            except OSError:
                                pass
                        break
        except Exception:
            pass

    # ── File rows ─────────────────────────────────────────────────────────────

    def _populate_file_rows(self) -> None:
        if not self.data.folder_path.exists():
            return
        files = sorted(
            [p for p in self.data.folder_path.iterdir() if p.is_file()],
            key=lambda p: file_sort_key(p.name),
        )
        for f in files:
            row_w = FileRowWidget(f, pjob_widget=self)
            self._files_layout.addWidget(row_w)
            self._file_row_widgets.append(row_w)
            self.data.file_rows.append(
                FileRowData(
                    old_path=f,
                    suffix=row_w.suffix,
                    extension=row_w.extension,
                    base_color=row_w.base_color,
                )
            )

    def remove_file_row(self, row_widget: FileRowWidget) -> None:
        self._files_layout.removeWidget(row_widget)
        row_widget.deleteLater()
        if row_widget in self._file_row_widgets:
            self._file_row_widgets.remove(row_widget)
        self.update_preview()

    # ── Preview update ────────────────────────────────────────────────────────

    def update_preview(self) -> None:
        """Recompute all filename previews and update collision/folder labels."""
        ch = re.sub(r'[^a-zA-Z0-9]', '', self._tb_char.text())
        ad = re.sub(r'[^a-zA-Z0-9]', '', self._cb_adj.currentText())
        th = self._gp_widget.theme()
        pf = self._gp_widget.prefix()

        # Card title: CamelCase split, adj in parens, ALL CAPS
        ch_spaced = re.sub(r'([a-z])([A-Z])', r'\1 \2', ch)
        ad_spaced = re.sub(r'([a-z])([A-Z])', r'\1 \2', ad)
        display = ch_spaced + (f' ({ad_spaced})' if ad_spaced else '')
        self._lbl_char_card.setText(display.upper())

        # Build target names and detect collisions
        name_counts: dict[str, int] = {}
        target_names: dict[FileRowWidget, str] = {}

        for row_w in self._file_row_widgets:
            sfx = row_w.get_suffix()
            parts = [p for p in (pf, ch, ad, th, sfx) if p]
            target = '_'.join(parts) + row_w.extension
            target_names[row_w] = target
            name_counts[target] = name_counts.get(target, 0) + 1

        has_collision = False
        for row_w in self._file_row_widgets:
            target = target_names[row_w]
            sfx = row_w.get_suffix()
            sfx_part = (f'_{sfx}' if sfx else '') + row_w.extension
            base_len = len(target) - len(sfx_part)
            base = target[:base_len] if base_len > 0 else ''
            collision = name_counts[target] > 1
            if collision:
                has_collision = True
            row_w.set_new_name(base, sfx_part, collision)

        # Store target names on data (for use during processing)
        for row_w, dr in zip(self._file_row_widgets, self.data.file_rows):
            dr.target_name = target_names.get(row_w, '')

        # Folder preview
        folder_parts = [p for p in (pf, ch, ad, th) if p]
        folder_preview = '_'.join(folder_parts)
        if folder_preview:
            self._lbl_folder.setText(f'Folder: {folder_preview}')
        else:
            self._lbl_folder.setText(f'Folder: {self.data.folder_path.name}')

        self.data.has_collision = has_collision
        self._validate()
        self._main_window.update_process_all_button()

        # Also update GP preview label
        if hasattr(self._gp_widget, 'update_all_previews'):
            self._gp_widget._lbl_preview.setText(self._gp_widget._make_preview_text())

    def all_colors_matched(self) -> bool:
        return all(s.is_matched() for s in self._color_slot_widgets)

    def set_status_text(self, text: str) -> None:
        self._lbl_card_status.setText(f'[ {text} ]')
        self._lbl_pick_status.setText(f'[ {text} ]')
        self._btn_apply.setText(text)

    # ── Validation ────────────────────────────────────────────────────────────

    def _validate(self) -> None:
        if self.data.is_queued or self.data.is_done:
            return
        if self.data.has_collision:
            self._btn_apply.setText('Name Collision!')
            self._btn_apply.setStyleSheet(f'background:{COLOR_RED}; color:white; font-weight:bold; border:none; border-radius:4px;')
            self._btn_apply.setEnabled(False)
        elif not self.all_colors_matched():
            self._btn_apply.setText('Unmatched Colors')
            self._btn_apply.setStyleSheet(f'background:{COLOR_AMBER}; color:white; font-weight:bold; border:none; border-radius:4px;')
            self._btn_apply.setEnabled(False)
        else:
            self._btn_apply.setText('Add to Queue')
            self._btn_apply.setStyleSheet(f'background:{COLOR_GREEN}; color:white; font-weight:bold; border:none; border-radius:4px;')
            self._btn_apply.setEnabled(True)

    # ── Enqueue + Worker ──────────────────────────────────────────────────────

    def enqueue(self, gp_widget=None) -> None:
        """Validate and add this job to the main window's processing queue."""
        if self.data.is_queued or self.data.is_done or self.data.has_collision:
            return
        if not self.all_colors_matched():
            return
        gp_widget = gp_widget or self._gp_widget

        # First job in group: confirm GP rename if needed
        if not gp_widget.gp_rename_confirmed if hasattr(gp_widget, 'gp_rename_confirmed') else False:
            pass  # Handled per PS1 logic in launch_worker

        self.data.is_queued = True
        self._btn_apply.setText('Queued...')
        self._btn_apply.setStyleSheet(f'background:{COLOR_AMBER}; color:white; font-weight:bold; border:none; border-radius:4px;')
        self.setEnabled(False)
        self._proc_overlay.setVisible(True)
        self._pick_proc_overlay.setVisible(True)
        self._lbl_card_status.setText('[ PREPARING ]')
        self._lbl_pick_status.setText('[ PREPARING ]')

        self._main_window.enqueue_job({'pjob_widget': self, 'gp_widget': gp_widget})

    def launch_worker(self, gp_widget) -> subprocess.Popen:
        """
        Apply color substitutions, rename files, then launch the async Python worker.
        Returns the Popen handle.
        Mirrors PS1 Start-NextProcess.
        """
        import sys as _sys
        self._btn_apply.setText('Processing...')
        self.data.is_queued = False

        th = gp_widget.theme()
        pf = gp_widget.prefix()

        # Apply color substitutions to temp work files
        self._apply_color_substitutions()

        # Rename anchor file
        self._rename_files(pf, th)

        # Handle grandparent rename
        self._rename_grandparent(gp_widget)

        # Build worker script path
        anchor_path = self.data.processed_anchor_path or str(self.data.anchor_file)
        worker_py = self._build_worker_script(anchor_path, gp_widget)

        proc = subprocess.Popen(
            [_sys.executable, worker_py],
            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0,
        )
        return proc

    def _apply_color_substitutions(self) -> None:
        """Patch color hex values in the temp work files and re-zip into the source 3MF."""
        text_exts = {'.xml', '.model', '.config', '.json'}
        modified_files: list[Path] = []

        all_files = [f for f in self.data.temp_work.rglob('*') if f.is_file() and f.suffix.lower() in text_exts]
        for f in all_files:
            try:
                content = f.read_text(encoding='utf-8', errors='replace')
                changed = False
                for slot_w in self._color_slot_widgets:
                    sel_name = slot_w.current_name()
                    if self._color_lib.contains_name(sel_name):
                        new_hex = (self._color_lib.hex_for_name(sel_name) or slot_w.orig_hex).upper()
                        old_hex = slot_w.orig_hex.upper()
                        old9 = old_hex if len(old_hex) == 9 else old_hex + 'FF'
                        old7 = old9[:7]
                        new9 = new_hex if len(new_hex) == 9 else new_hex + 'FF'
                        new7 = new9[:7]
                        new_content = content.replace(old9, new9).replace(old7, new7)
                        if new_content != content:
                            content = new_content
                            changed = True
                if changed:
                    f.write_text(content, encoding='utf-8')
                    modified_files.append(f)
            except Exception:
                pass

        if not modified_files:
            return

        src = self.data.anchor_file
        if src.exists():
            try:
                with zipfile.ZipFile(src, 'a') as zf:
                    for mf in modified_files:
                        rel = mf.relative_to(self.data.temp_work).as_posix()
                        try:
                            info = zf.getinfo(rel)
                            # ZipFile.open for write isn't straightforward; delete + add
                            # We work around by writing the whole file fresh
                        except KeyError:
                            pass
                        zf.write(mf, rel)
            except Exception:
                pass

    def _rename_files(self, prefix: str, theme: str) -> None:
        """Rename all files per computed target names. Sets processed_anchor_path."""
        for row_w in self._file_row_widgets:
            target = row_w.target_name if hasattr(row_w, 'target_name') else ''
            if not target:
                continue
            old_path = row_w.old_path
            new_path = old_path.parent / target
            if old_path != new_path and old_path.exists():
                try:
                    if new_path.exists():
                        new_path.unlink()
                    old_path.rename(new_path)
                    row_w.old_path = new_path
                    row_w._file_path = new_path
                except OSError:
                    pass
            # Track the new anchor file path
            if row_w.old_path.name.lower().endswith('full.3mf') and \
               not row_w.old_path.name.lower().endswith('.gcode.3mf'):
                self.data.processed_anchor_path = str(new_path)

        # Fallback if no anchor row matched
        if not self.data.processed_anchor_path:
            self.data.processed_anchor_path = str(self.data.anchor_file)

        # Rename parent folder
        ch = re.sub(r'[^a-zA-Z0-9]', '', self._tb_char.text())
        ad = re.sub(r'[^a-zA-Z0-9]', '', self._cb_adj.currentText())
        parts = [p for p in (prefix, ch, ad, theme) if p]
        new_folder_name = '_'.join(parts)
        if new_folder_name and new_folder_name != self.data.folder_path.name:
            new_folder = self.data.folder_path.parent / new_folder_name
            try:
                self.data.folder_path.rename(new_folder)
                # Update anchor path
                if self.data.processed_anchor_path:
                    self.data.processed_anchor_path = self.data.processed_anchor_path.replace(
                        str(self.data.folder_path), str(new_folder)
                    )
                self.data.folder_path = new_folder
            except OSError:
                pass

    def _rename_grandparent(self, gp_widget) -> None:
        """Rename the grandparent folder if not skipping."""
        if gp_widget.skip_rename():
            return
        if gp_widget.gp_path.startswith('ROOT_'):
            return
        th = gp_widget.theme()
        pf = gp_widget.prefix()
        new_gp_name = f'{pf}_{th}' if pf and th else (pf or th)
        if not new_gp_name:
            return
        gp_dir = Path(gp_widget.gp_path)
        if not gp_dir.exists() or gp_dir.name == new_gp_name:
            return
        new_gp = gp_dir.parent / new_gp_name
        try:
            gp_dir.rename(new_gp)
            gp_widget.gp_path = str(new_gp)
            # Update our own folder path
            if str(self.data.folder_path).startswith(str(gp_dir)):
                self.data.folder_path = new_gp / self.data.folder_path.relative_to(gp_dir)
        except OSError:
            pass

    def _build_worker_script(self, anchor_path: str, gp_widget) -> str:
        """Build a temp Python worker script and return its path."""
        import uuid
        worker_path = Path(tempfile.gettempdir()) / f'AsyncWorker_{uuid.uuid4().hex[:8]}.py'

        anchor      = Path(anchor_path)
        base_name   = anchor.stem
        base_prefix = base_name[:-4] if base_name.lower().endswith('full') else base_name + '_'

        folder      = self.data.folder_path
        status_file = folder / 'AsyncWorker_Status.txt'
        nest_path   = folder / f'{base_prefix}Nest.3mf'
        final_path  = folder / f'{base_prefix}Final.3mf'
        temp_out    = folder / f'{base_name}_merged_temp.3mf'
        temp_iso    = Path(tempfile.gettempdir()) / f'iso_{uuid.uuid4().hex[:8]}'
        tsv_base    = re.sub(r'(?i)_Full$', '', base_name)
        sliced_file = folder / f'{base_name}.gcode.3mf'
        single_file = folder / f'{base_prefix}Final.gcode.3mf'
        tsv_file    = folder / f'{tsv_base}_Data.tsv'
        log_file    = folder / 'Worker_Py_Log.txt'

        py_workers  = self._script_dir / 'py_workers'
        work_dir    = self.data.temp_work
        bat_image   = self._script_dir.parent.parent / 'callers' / 'ReplaceImageNew.bat'

        do_merge   = self._chk_merge.isChecked()
        do_slice   = self._chk_slice.isChecked()
        do_extract = self._chk_extract.isChecked()
        do_image   = self._chk_image.isChecked()
        do_logs    = self._chk_logs.isChecked()

        def q(p): return repr(str(p))

        I = '    '
        lines: list[str] = [
            'import sys, subprocess, shutil, zipfile, time',
            'from pathlib import Path',
            '',
            f'PY  = Path({q(py_workers)})',
            f'STA = Path({q(status_file)})',
            f'WRK = Path({q(work_dir)})',
            f'ANC = Path({q(anchor_path)})',
            f'NST = Path({q(nest_path)})',
            f'FIN = Path({q(final_path)})',
            f'TMP = Path({q(temp_out)})',
            f'ISO = Path({q(temp_iso)})',
            f'SLC = Path({q(sliced_file)})',
            f'SNG = Path({q(single_file)})',
            f'TSV = Path({q(tsv_file)})',
            f'LOG = open({q(log_file)}, "w", encoding="utf-8", buffering=1)',
            '',
            'def ws(m):',
            '    STA.write_text(m, encoding="utf-8")',
            '    LOG.write("[STATUS] " + m + "\\n"); LOG.flush()',
            '',
            'def go(label, *a):',
            '    _cmd = " ".join(str(x) for x in a)',
            '    LOG.write("\\n[RUN] " + label + "\\n  " + _cmd + "\\n"); LOG.flush()',
            '    r = subprocess.run([sys.executable, *a], capture_output=True, text=True, errors="replace")',
            '    if r.stdout: LOG.write(r.stdout)',
            '    if r.stderr: LOG.write(f"[STDERR]\\n{r.stderr}")',
            '    LOG.write(f"[EXIT] {r.returncode}\\n"); LOG.flush()',
            '    return r.returncode == 0',
            '',
            'try:',
        ]

        if do_merge:
            lines += [
                f"{I}ws('MERGING...')",
                f"{I}ok = go('merge',",
                f"{I}        str(PY / 'merge_worker.py'),",
                f"{I}        '--work-dir', str(WRK),",
                f"{I}        '--input-path', str(ANC),",
                f"{I}        '--output-path', str(TMP),",
                f"{I}        '--do-colors', '0')",
                f"{I}if not (ok and TMP.exists()):",
                f"{I}    ws('[ERROR] Merge produced no output - check log')",
                f"{I}    raise SystemExit(1)",
                f"{I}NST.unlink(missing_ok=True)",
                f"{I}ANC.rename(NST)",
                f"{I}TMP.rename(ANC)",
                f"{I}ws('ISOLATING FINAL...')",
                f"{I}ISO.mkdir(parents=True, exist_ok=True)",
                f"{I}with zipfile.ZipFile(str(NST)) as _z: _z.extractall(str(ISO))",
                f"{I}go('isolate', str(PY / 'isolate_worker.py'), '--work-dir', str(ISO), '--output-path', str(FIN))",
                f"{I}shutil.rmtree(str(ISO), ignore_errors=True)",
            ]

        if do_slice:
            lines += [
                f"{I}ws('SLICING...')",
                f"{I}slice_ok = go('slice', str(PY / 'slice_worker.py'),",
                f"{I}              '--input-path', str(ANC), '--isolated-path', str(FIN))",
                f"{I}if not slice_ok:",
                f"{I}    ws('[ERROR] Slice failed - check log')",
                f"{I}    raise SystemExit(1)",
            ]
        elif do_extract or do_image:
            riso = repr(str(Path(tempfile.gettempdir()) / f'iso_rs_{uuid.uuid4().hex[:8]}'))
            lines += [
                f"{I}ws('RE-SLICING FINAL FOR DATA...')",
                f"{I}if not FIN.exists() and NST.exists():",
                f"{I}    _riso = Path({riso})",
                f"{I}    _riso.mkdir(parents=True, exist_ok=True)",
                f"{I}    with zipfile.ZipFile(str(NST)) as _z: _z.extractall(str(_riso))",
                f"{I}    go('isolate-reslice', str(PY / 'isolate_worker.py'), '--work-dir', str(_riso), '--output-path', str(FIN))",
                f"{I}    shutil.rmtree(str(_riso), ignore_errors=True)",
                f"{I}if FIN.exists():",
                f"{I}    slice_ok = go('slice-final', str(PY / 'slice_worker.py'), '--input-path', str(FIN))",
                f"{I}    if not slice_ok:",
                f"{I}        ws('[ERROR] Slice failed - check log')",
                f"{I}        raise SystemExit(1)",
            ]

        if do_extract or do_image:
            img_arg = ", '--generate-image'" if do_image else ''
            lines += [
                f"{I}ws('EXTRACTING DATA...')",
                f"{I}if SLC.exists():",
                f"{I}    TSV.unlink(missing_ok=True)",
                f"{I}    go('extract',",
                f"{I}       str(PY / 'extract_worker.py'),",
                f"{I}       '--input-file', str(SLC),",
                f"{I}       '--single-file', str(SNG),",
                f"{I}       '--individual-tsv-path', str(TSV){img_arg})",
                f"{I}    SNG.unlink(missing_ok=True)",
                f"{I}else:",
                f"{I}    ws('[ERROR] SLICE FAILED - MISSING GCODE')",
                f"{I}    raise SystemExit(1)",
            ]

        if do_image:
            lines += [
                f"{I}ws('IMAGE INJECTION...')",
                f"{I}_bat = Path({q(bat_image)})",
                f"{I}if _bat.exists():",
                f"{I}    subprocess.run(['cmd.exe', '/c', str(_bat), str(ANC.parent)], check=False)",
            ]

        lines += [
            'except SystemExit:',
            f"{I}pass",
            'except Exception as _e:',
            f"{I}LOG.write(f'[EXCEPTION] {{_e}}\\n'); LOG.flush()",
            f"{I}ws(f'[ERROR] Unexpected error - check log')",
            'finally:',
            f"{I}STA.unlink(missing_ok=True)",
            f"{I}shutil.rmtree(str(WRK), ignore_errors=True)",
        ]

        if not do_logs:
            lines += [
                f"{I}LOG.close()",
                f"{I}Path({q(log_file)}).unlink(missing_ok=True)",
                f"{I}for _p in (list(ANC.parent.glob('*ProcessLog*.txt'))",
                f"{I}           + list(ANC.parent.glob('*_Log.txt'))",
                f"{I}           + list(ANC.parent.glob('*_Errors.txt'))):",
                f"{I}    _p.unlink(missing_ok=True)",
            ]
        else:
            lines.append(f"{I}LOG.close()")

        script_text = '\n'.join(lines)
        worker_path.write_text(script_text, encoding='utf-8')
        # Also copy to job folder for easy inspection
        debug_copy = folder / 'LastWorkerScript.py'
        debug_copy.write_text(script_text, encoding='utf-8')
        return str(worker_path)

    def on_worker_finished(self, gp_widget) -> None:
        """Handle processing completion. Mirrors PS1 queueTimer finished branch."""
        self.data.is_done = True
        self.data.is_queued = False
        self.setEnabled(True)

        # Try to reload plate and pick images from the finished gcode file
        base_name = Path(self.data.processed_anchor_path or str(self.data.anchor_file)).stem
        gcode_file = self.data.folder_path / f'{base_name}.gcode.3mf'
        if gcode_file.exists():
            try:
                with zipfile.ZipFile(gcode_file, 'r') as zf:
                    names = {e.replace('\\', '/').lower(): e for e in zf.namelist()}
                    for cand in ('metadata/plate_1.png', 'metadata/thumbnail.png'):
                        if cand in names:
                            raw = zf.read(names[cand])
                            pm = bytes_to_qpixmap(raw)
                            if pm and not pm.isNull():
                                self._img_plate_finished.setPixmap(pm)
                                self._img_plate.setPixmap(pm)
                                self._finished_overlay.setVisible(True)
                            break
                    for norm, orig in names.items():
                        if norm.endswith('pick_1.png'):
                            raw = zf.read(orig)
                            colored = randomize_pick_colors_from_bytes(raw) or raw
                            pm = bytes_to_qpixmap(colored)
                            if pm and not pm.isNull():
                                self._img_pick.setPixmap(pm)
                            break
            except Exception:
                pass

        # Verify color slots
        for sw in self._color_slot_widgets:
            if self._color_lib.contains_name(sw.current_name()):
                sw.set_verified()

        # Reload file rows
        for w in list(self._file_row_widgets):
            self._files_layout.removeWidget(w)
            w.deleteLater()
        self._file_row_widgets.clear()
        self.data.file_rows.clear()
        self._populate_file_rows()
        self.update_preview()

        # Switch to KEEP / REVERT state
        self._proc_overlay.setVisible(False)
        self._pick_proc_overlay.setVisible(False)
        self._chk_merge.setEnabled(True)
        self._chk_slice.setEnabled(True)
        self._chk_extract.setEnabled(True)
        self._chk_image.setEnabled(True)

        self._btn_apply.setText('KEEP')
        self._btn_apply.setStyleSheet(f'background:{COLOR_GREEN}; color:white; font-weight:bold; border:none; border-radius:4px;')
        self._btn_apply.setEnabled(True)
        self._btn_apply.setFixedWidth(70)
        self._btn_revert_done.setVisible(True)

    # ── Button handlers ───────────────────────────────────────────────────────

    def _on_apply(self) -> None:
        if self._btn_apply.text() == 'KEEP':
            # Accept finished result
            self._finished_overlay.setVisible(False)
            self._btn_apply.setText('Finished')
            self._btn_apply.setStyleSheet(f'background:#333333; color:#666666; font-weight:bold; border:none; border-radius:4px;')
            self._btn_apply.setEnabled(False)
            self._btn_apply.setFixedWidth(150)
            self._btn_revert_done.setVisible(False)
        else:
            self.enqueue()

    def _on_revert_done(self) -> None:
        self._do_revert()

    def _on_revert_merge(self) -> None:
        self._do_revert()

    def _do_revert(self) -> None:
        """Undo a merge in pure Python: delete generated files, restore Nest → Full."""
        self._btn_apply.setText('Reverting...')
        self.setEnabled(False)
        self._proc_overlay.setVisible(True)
        self._lbl_card_status.setText('[REVERTING...]')

        folder = self.data.folder_path
        # Find any *Nest.3mf files to drive the revert
        nest_files = list(folder.glob('*Nest.3mf'))
        if not nest_files:
            # Fall back: derive from anchor name
            anchor = Path(self.data.processed_anchor_path or str(self.data.anchor_file))
            base   = anchor.stem  # e.g. P2S_Cyclops_Full
            prefix = base[:-4]    # e.g. P2S_Cyclops_  (removes 'Full')
            nest_files = [folder / f'{prefix}Nest.3mf']

        for nest_path in nest_files:
            stem = nest_path.stem                      # e.g. P2S_Cyclops_Nest
            sep  = stem[-5] if len(stem) >= 5 else '_' # char before 'Nest'
            base_prefix = stem[:-5]                    # e.g. P2S_Cyclops
            full_base   = f'{base_prefix}{sep}Full'
            final_base  = f'{base_prefix}{sep}Final'
            tsv_base    = base_prefix.rstrip('_. ')

            # Delete generated output files
            for fname in [
                f'{full_base}.gcode.3mf',
                f'{full_base}.3mf',
                f'{tsv_base}_Data.tsv',
                f'{final_base}.3mf',
                f'{final_base}.gcode.3mf',
            ]:
                p = folder / fname
                if p.exists():
                    try:
                        p.unlink()
                    except OSError:
                        pass

            # Restore Nest → Full
            full_path = folder / f'{full_base}.3mf'
            if nest_path.exists():
                try:
                    nest_path.rename(full_path)
                except OSError as e:
                    QMessageBox.warning(self, 'Revert Error', f'Could not restore file:\n{e}')

        QTimer.singleShot(50, self._on_refresh)

    def _on_refresh(self) -> None:
        """Rebuild this widget from disk. Mirrors PS1 Refresh-PJob."""
        # Clean up temp
        if self.data.temp_work.exists():
            shutil.rmtree(self.data.temp_work, ignore_errors=True)

        # Find new anchor
        if not self.data.folder_path.exists():
            return
        anchors = sorted(
            [p for p in self.data.folder_path.iterdir()
             if p.is_file() and p.name.lower().endswith('full.3mf')
             and not p.name.lower().endswith('.gcode.3mf')],
            key=lambda p: p.stat().st_mtime, reverse=True,
        )
        if not anchors:
            return

        # Ask the GP widget to rebuild us at the same position
        if hasattr(self._gp_widget, 'rebuild_parent_widget'):
            self._gp_widget.rebuild_parent_widget(self, anchors[0])

    def _on_remove_folder(self) -> None:
        self._gp_widget.remove_parent_widget(self)

    def _on_delete_logs(self) -> None:
        if not self.data.folder_path.exists():
            return
        logs = list(self.data.folder_path.rglob('*_Log.txt')) + \
               list(self.data.folder_path.rglob('*_Errors.txt')) + \
               list(self.data.folder_path.rglob('*ProcessLog*.txt'))
        if not logs:
            QMessageBox.information(self, 'No Logs', 'No log files found in this folder.')
            return
        res = QMessageBox.warning(
            self, 'Confirm Delete Logs',
            f'Found {len(logs)} log files. Delete them?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if res == QMessageBox.StandardButton.Yes:
            for log in logs:
                try:
                    log.unlink()
                except OSError:
                    pass
            QMessageBox.information(self, 'Logs Cleared', 'Logs deleted successfully.')

    def _on_browse_image(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, 'Select Custom Card Image', str(self.data.folder_path),
            'Image Files (*.png *.jpg *.jpeg)',
        )
        if not path:
            return
        src = Path(path)
        dest = self.data.folder_path / src.name
        if src != dest:
            shutil.copy2(src, dest)
        self.data.custom_image_path = dest
        pm = QPixmap(str(dest))
        if not pm.isNull():
            self._img_plate.setPixmap(pm)
        self._btn_browse_img.setStyleSheet(
            f'background:{COLOR_GREEN}; color:white; font-weight:bold; border:none; border-radius:2px;'
        )

    def _on_color_status_changed(self) -> None:
        self._validate()
        self._main_window.update_process_all_button()

    # ── Double-click handlers ─────────────────────────────────────────────────

    def _on_plate_dbl_click(self, event) -> None:
        pm = self._img_plate.pixmap()
        if pm and not pm.isNull():
            self._show_image_popup('Card Image', pm)

    def _on_finished_dbl_click(self, event) -> None:
        pm = self._img_plate_finished.pixmap()
        if pm and not pm.isNull():
            self._show_image_popup('Finished Plate', pm)

    def _on_pick_dbl_click(self, event) -> None:
        """Double-click on pick image: show merge map overlay."""
        if not self._active_pick_path:
            return
        post_path = Path(self._active_pick_path)
        if not post_path.exists():
            return
        # Find Nest.3mf
        base = re.sub(r'(?i)_?Full\.gcode\.3mf$|_?Full\.3mf$', '', self.data.anchor_file.name)
        nest = self.data.folder_path / f'{base}_Nest.3mf'
        if not nest.exists():
            QMessageBox.warning(self, 'Merge Map', f'Nest.3mf not found:\n{nest}')
            return
        try:
            with zipfile.ZipFile(nest, 'r') as zf:
                names = {e.replace('\\', '/').lower(): e for e in zf.namelist()}
                pre_bytes = None
                for norm, orig in names.items():
                    if norm.endswith('pick_1.png'):
                        pre_bytes = zf.read(orig)
                        break
                if not pre_bytes:
                    QMessageBox.warning(self, 'Merge Map', 'No pick_1.png in Nest.3mf')
                    return
            post_bytes = post_path.read_bytes()
            from image_utils import get_merge_map_image_bytes
            result_bytes = get_merge_map_image_bytes(pre_bytes, post_bytes)
            if result_bytes:
                pm = bytes_to_qpixmap(result_bytes)
                if pm and not pm.isNull():
                    self._show_image_popup('Merged RGB Verification Overlay', pm)
        except Exception as exc:
            QMessageBox.warning(self, 'Merge Map Error', str(exc))

    def _show_image_popup(self, title: str, pixmap: QPixmap) -> None:
        from PySide6.QtWidgets import QDialog
        dlg = QDialog(self)
        dlg.setWindowTitle(title)
        dlg.setStyleSheet('background:#0D0E10;')
        v = QVBoxLayout(dlg)
        lbl = QLabel()
        lbl.setPixmap(pixmap.scaled(900, 900, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        v.addWidget(lbl)
        dlg.exec()

    # ── Drag & drop on card panel ─────────────────────────────────────────────

    def _card_drag_enter(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and urls[0].toLocalFile().lower().endswith(('.png', '.jpg', '.jpeg')):
                event.setDropAction(Qt.DropAction.CopyAction)
                event.accept()
                return
        event.ignore()

    def _card_drop(self, event: QDropEvent) -> None:
        urls = event.mimeData().urls()
        if not urls:
            return
        src = Path(urls[0].toLocalFile())
        if not src.suffix.lower() in ('.png', '.jpg', '.jpeg'):
            return
        dest = self.data.folder_path / src.name
        if src != dest:
            shutil.copy2(src, dest)
        self.data.custom_image_path = dest
        pm = QPixmap(str(dest))
        if not pm.isNull():
            self._img_plate.setPixmap(pm)
        self._btn_browse_img.setStyleSheet(
            f'background:{COLOR_GREEN}; color:white; font-weight:bold; border:none; border-radius:2px;'
        )
        event.accept()

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _muted_label(text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setStyleSheet(f'color:{COLOR_TEXT_MUTED}; font-size:12px; background:transparent;')
        return lbl

    @staticmethod
    def _make_checkbox(label: str, checked: bool) -> QCheckBox:
        chk = QCheckBox(label)
        chk.setChecked(checked)
        chk.setStyleSheet(f'color:{COLOR_TEXT_WHITE}; background:transparent;')
        return chk
