# main_window.py - Main application window (replaces PS1 XAML window + top-level handlers)
from __future__ import annotations

import shutil
from pathlib import Path

from PySide6.QtCore import Qt, QTimer, QMimeData
from PySide6.QtGui import QColor, QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QScrollArea, QSizePolicy,
    QFrame, QApplication,
)

from models import (
    COLOR_BG_DARK, COLOR_BG_PANEL, COLOR_BG_HEADER,
    COLOR_TEXT_WHITE, COLOR_TEXT_MUTED,
    COLOR_GREEN, COLOR_BLUE,
)
from color_library import ColorLibrary
from file_utils import extract_3mf_to_temp, read_3mf_colors, read_3mf_images
from theme import DARK_QSS


class MainWindow(QMainWindow):
    """
    Top-level application window.

    Mirrors the PS1 WPF window structure:
      Row 0  - Header bar (title, Browse button, Process All button)
      Row 1  - Scrollable main area (MainStack of GpJob widgets)
    """

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle('Batch Pre-Flight Editor')
        self.resize(1550, 850)
        self.setMinimumSize(1100, 600)
        self.setAcceptDrops(True)
        self.setStyleSheet(DARK_QSS)

        # Resolve worker directory (same folder as this script)
        self._script_dir = Path(__file__).resolve().parent
        csv_path = self._script_dir.parent / 'colorNamesCSV.csv'
        self._color_lib = ColorLibrary(csv_path)

        # Active grandparent-job widgets (mirrors $script:jobs)
        self._gp_widgets: list = []   # list[GpWidget]

        # Processing queue
        self._process_queue: list = []   # list[dict]
        self._active_process = None      # subprocess.Popen | None
        self._active_job_ctx = None      # {pjob_widget, gp_widget}

        self._build_ui()
        self._start_queue_timer()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)
        self.setCentralWidget(root)

        # Header bar
        header = self._build_header()
        root_layout.addWidget(header)

        # Scrollable main area
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._scroll.setStyleSheet('background:#0D0E10; border:none;')

        scroll_content = QWidget()
        scroll_content.setStyleSheet('background:#0D0E10;')
        self._main_layout = QVBoxLayout(scroll_content)
        self._main_layout.setContentsMargins(15, 15, 15, 15)
        self._main_layout.setSpacing(0)
        self._main_layout.addStretch(1)   # push jobs to top

        self._scroll.setWidget(scroll_content)
        root_layout.addWidget(self._scroll, stretch=1)

    def _build_header(self) -> QWidget:
        header = QWidget()
        header.setObjectName('HeaderBar')
        header.setFixedHeight(60)
        header.setStyleSheet(f'background:{COLOR_BG_PANEL}; border-bottom:1px solid {COLOR_BG_HEADER};')

        layout = QHBoxLayout(header)
        layout.setContentsMargins(15, 0, 15, 0)

        # Title label (left)
        self._lbl_title = QLabel('Loading files into queue...')
        self._lbl_title.setStyleSheet(
            f'color:{COLOR_TEXT_WHITE}; font-size:18px; font-weight:bold; background:transparent;'
        )
        layout.addWidget(self._lbl_title)

        layout.addStretch(1)

        # Browse button + hint (center)
        center = QWidget()
        center_v = QVBoxLayout(center)
        center_v.setContentsMargins(0, 0, 0, 0)
        center_v.setSpacing(2)
        center_v.setAlignment(Qt.AlignmentFlag.AlignHCenter)

        self._btn_browse = QPushButton('Browse Files')
        self._btn_browse.setFixedSize(140, 30)
        self._btn_browse.setStyleSheet(
            f'background:{COLOR_BLUE}; color:{COLOR_TEXT_WHITE}; font-weight:bold; border:none; border-radius:4px;'
        )
        self._btn_browse.setCursor(Qt.CursorShape.PointingHandCursor)
        self._btn_browse.clicked.connect(self._on_browse)
        center_v.addWidget(self._btn_browse, alignment=Qt.AlignmentFlag.AlignHCenter)

        hint = QLabel('Browse or drop files to add')
        hint.setStyleSheet(f'color:{COLOR_TEXT_MUTED}; font-size:10px; background:transparent;')
        center_v.addWidget(hint, alignment=Qt.AlignmentFlag.AlignHCenter)

        layout.addWidget(center)
        layout.addStretch(1)

        # Process All button (right)
        self._btn_process_all = QPushButton('Process All Queued')
        self._btn_process_all.setFixedSize(150, 30)
        self._btn_process_all.setStyleSheet(
            f'background:{COLOR_GREEN}; color:{COLOR_TEXT_WHITE}; font-weight:bold; border:none; border-radius:4px;'
        )
        self._btn_process_all.setCursor(Qt.CursorShape.PointingHandCursor)
        self._btn_process_all.clicked.connect(self._on_process_all)
        layout.addWidget(self._btn_process_all)

        return header

    # ── Public API ────────────────────────────────────────────────────────────

    def load_paths(self, paths: list[Path]) -> None:
        """
        Scan a list of paths (files or folders) for *Full.3mf files and add
        them to the UI.  Mirrors PS1 startup + BtnBrowse logic.
        """
        self._lbl_title.setText('Scanning selected folders...')
        QApplication.processEvents()

        found: list[Path] = []
        for p in paths:
            if p.is_dir():
                found.extend(p.rglob('*Full.3mf'))
            elif p.is_file() and p.name.lower().endswith('full.3mf'):
                found.append(p)

        self._ingest_found_files(found)
        self._update_title()

    # ── Drag & drop ───────────────────────────────────────────────────────────

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.DropAction.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:
        urls = event.mimeData().urls()
        paths = [Path(u.toLocalFile()) for u in urls if u.isLocalFile()]
        if paths:
            self._lbl_title.setText('Scanning dropped folders...')
            QApplication.processEvents()
            self.load_paths(paths)

    # ── File ingestion ────────────────────────────────────────────────────────

    def _ingest_found_files(self, found: list[Path]) -> None:
        """
        Group found *Full.3mf files by grandparent -> parent structure and
        add them to existing or new GpWidgets.  Mirrors PS1 gpQueue logic.
        """
        from gp_widget import GpWidget

        # Build gp_path -> {parent_path -> anchor_file} mapping
        gp_queue: dict[str, dict[str, Path]] = {}
        for f in found:
            parent_path = str(f.parent)
            gp = f.parent.parent
            gp_path = str(gp) if gp else f'ROOT_{parent_path}'

            # Skip duplicates already in UI
            if self._find_parent_widget(parent_path):
                continue

            if gp_path not in gp_queue:
                gp_queue[gp_path] = {}
            if parent_path not in gp_queue[gp_path]:
                gp_queue[gp_path][parent_path] = f

        for gp_path, parent_dict in gp_queue.items():
            existing_gp = self._find_gp_widget(gp_path)
            if existing_gp:
                existing_gp.add_parents(parent_dict)
            else:
                gp_widget = GpWidget(
                    gp_path=gp_path,
                    parent_dict=parent_dict,
                    color_lib=self._color_lib,
                    script_dir=self._script_dir,
                    main_window=self,
                )
                self._add_gp_widget(gp_widget)

    def _add_gp_widget(self, widget) -> None:
        """Insert a GpWidget before the trailing stretch."""
        # Layout: [...gp_widgets..., stretch]
        insert_pos = self._main_layout.count() - 1
        self._main_layout.insertWidget(insert_pos, widget)
        self._gp_widgets.append(widget)
        self.update_process_all_button()

    def remove_gp_widget(self, widget) -> None:
        """Called by GpWidget when the user clicks Remove Group."""
        self._main_layout.removeWidget(widget)
        widget.deleteLater()
        if widget in self._gp_widgets:
            self._gp_widgets.remove(widget)
        self.update_process_all_button()

    # ── Button handlers ───────────────────────────────────────────────────────

    def _on_browse(self) -> None:
        from PySide6.QtWidgets import QFileDialog
        dialog = QFileDialog(self, 'Select folders containing Full.3mf files')
        dialog.setFileMode(QFileDialog.FileMode.Directory)
        dialog.setOption(QFileDialog.Option.DontUseNativeDialog, False)
        # Allow multi-select via the native dialog when possible
        dialog.setOption(QFileDialog.Option.ShowDirsOnly, True)
        if dialog.exec():
            paths = [Path(p) for p in dialog.selectedFiles()]
            if paths:
                self._lbl_title.setText('Scanning selected folders...')
                QApplication.processEvents()
                self.load_paths(paths)

    def _on_process_all(self) -> None:
        """Queue every unqueued, valid parent job across all groups."""
        for gp_w in self._gp_widgets:
            for p_w in gp_w.parent_widgets():
                if not p_w.data.is_queued and not p_w.data.is_done:
                    if not p_w.data.has_collision and p_w.all_colors_matched():
                        p_w.enqueue(gp_w)

    # ── Global status helpers ─────────────────────────────────────────────────

    def update_process_all_button(self) -> None:
        """
        Enable/disable the Process All button depending on whether any job has
        a collision or unmatched color.  Mirrors PS1 Update-GlobalProcessAllStatus.
        """
        has_issue = False
        for gp_w in self._gp_widgets:
            for p_w in gp_w.parent_widgets():
                if p_w.data.is_queued or p_w.data.is_done:
                    continue
                if p_w.data.has_collision or not p_w.all_colors_matched():
                    has_issue = True
                    break
            if has_issue:
                break

        if has_issue:
            self._btn_process_all.setEnabled(False)
            self._btn_process_all.setStyleSheet(
                'background:#555555; color:#888888; font-weight:bold; border:none; border-radius:4px;'
            )
        else:
            self._btn_process_all.setEnabled(True)
            self._btn_process_all.setStyleSheet(
                f'background:{COLOR_GREEN}; color:{COLOR_TEXT_WHITE}; font-weight:bold; border:none; border-radius:4px;'
            )

    def _update_title(self) -> None:
        count = len(self._gp_widgets)
        self._lbl_title.setText(
            f'Queue Dashboard ({count} Theme{"s" if count != 1 else ""} found)'
        )

    # ── Queue management (mirrors PS1 queueTimer) ─────────────────────────────

    def _start_queue_timer(self) -> None:
        self._queue_timer = QTimer(self)
        self._queue_timer.setInterval(500)
        self._queue_timer.timeout.connect(self._on_queue_tick)
        self._queue_timer.start()

    def _on_queue_tick(self) -> None:
        if self._active_process is not None:
            if self._active_process.poll() is None:
                # Still running - poll status file
                ctx = self._active_job_ctx
                if ctx:
                    status_file = Path(ctx['pjob_widget'].data.folder_path) / 'AsyncWorker_Status.txt'
                    if status_file.exists():
                        try:
                            txt = status_file.read_text(encoding='utf-8', errors='replace').strip()
                            if txt:
                                ctx['pjob_widget'].set_status_text(txt)
                        except OSError:
                            pass
            else:
                # Job finished
                self._on_job_finished()
        else:
            self._start_next_process()

    def enqueue_job(self, job_ctx: dict) -> None:
        """Called by PJobWidget.enqueue() to add a job to the processing queue."""
        self._process_queue.append(job_ctx)

    def _start_next_process(self) -> None:
        if not self._process_queue:
            return
        ctx = self._process_queue.pop(0)
        self._active_job_ctx = ctx
        pjob_widget = ctx['pjob_widget']
        self._active_process = pjob_widget.launch_worker(ctx['gp_widget'])

    def _on_job_finished(self) -> None:
        ctx = self._active_job_ctx
        if ctx:
            ctx['pjob_widget'].on_worker_finished(ctx['gp_widget'])
        self._active_process = None
        self._active_job_ctx = None
        if self._process_queue:
            self._start_next_process()

    # ── Lookup helpers ────────────────────────────────────────────────────────

    def _find_gp_widget(self, gp_path: str):
        for w in self._gp_widgets:
            if w.gp_path == gp_path:
                return w
        return None

    def _find_parent_widget(self, folder_path: str):
        for gp_w in self._gp_widgets:
            for p_w in gp_w.parent_widgets():
                if str(p_w.data.folder_path) == folder_path:
                    return p_w
        return None

    # ── Cleanup ───────────────────────────────────────────────────────────────

    def closeEvent(self, event) -> None:
        self._queue_timer.stop()
        if self._active_process and self._active_process.poll() is None:
            self._active_process.kill()
        # Clean up temp directories
        for gp_w in self._gp_widgets:
            for p_w in gp_w.parent_widgets():
                tw = p_w.data.temp_work
                if tw and tw.exists():
                    try:
                        shutil.rmtree(tw, ignore_errors=True)
                    except Exception:
                        pass
        super().closeEvent(event)
