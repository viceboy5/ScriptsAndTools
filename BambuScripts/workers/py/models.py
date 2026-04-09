# models.py - Shared constants and pure data models for CardQueueEditor
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

# ── Shared constants ──────────────────────────────────────────────────────────
PRINTER_PREFIXES = ['P2S', 'X1C', 'H2S']
GP_THEMES = [
    'Fantasy', 'Puppies', 'Original', 'Ocean', 'Farm', 'Foodz',
    'StarsAndStripes', 'Spring', 'Prehistoric', 'Halloween 2025', 'Christmas 2025',
]
ADJ_PRESETS = ['Common', 'RARE', 'EPIC', 'LEGENDARY', 'Default']

# ── Color constants (dark UI palette) ─────────────────────────────────────────
COLOR_BG_DARK    = '#16171B'
COLOR_BG_PANEL   = '#1C1D23'
COLOR_BG_INPUT   = '#1E2028'
COLOR_BG_HEADER  = '#2A2C35'
COLOR_TEXT_WHITE = '#FFFFFF'
COLOR_TEXT_MUTED = '#A0A0A0'
COLOR_TEXT_DIM   = '#6B6E7A'
COLOR_ACCENT     = '#FFD700'
COLOR_GREEN      = '#4CAF72'
COLOR_RED        = '#D95F5F'
COLOR_AMBER      = '#E8A135'
COLOR_BLUE       = '#5A78C4'
COLOR_PURPLE     = '#B57BFF'
COLOR_PINK       = '#FF69B4'
COLOR_BLUE_GREY  = '#90B8C8'
COLOR_STEEL_BLUE = '#6B9FD4'

# File badge colors (match PS1 logic)
FILE_COLOR_FINAL    = '#FFD700'   # Yellow
FILE_COLOR_NEST     = '#FF69B4'   # Pink
FILE_COLOR_FULL     = '#B57BFF'   # Purple
FILE_COLOR_GCODE    = '#4CAF72'   # Green
FILE_COLOR_DEFAULT  = '#90B8C8'   # Blue-grey


# ── Data models ───────────────────────────────────────────────────────────────

@dataclass
class ColorSlotData:
    """Data for a single filament color slot extracted from a 3MF file."""
    orig_hex: str   # e.g. "#FCCE4AFF"
    name: str       # matched library name, or "" if unmatched


@dataclass
class FileRowData:
    """Data for a single file entry within a parent job."""
    old_path: Path
    suffix: str        # e.g. "01", "Full"
    extension: str     # e.g. ".3mf", ".gcode.3mf"
    base_color: str    # display hex for the suffix badge
    target_name: str = ''


@dataclass
class PJobData:
    """Data model for a parent job (one card folder)."""
    folder_path: Path
    anchor_file: Path          # the *Full.3mf anchor file
    temp_work: Path            # extracted temp directory (caller cleans up)
    processed_anchor_path: str = ''
    custom_image_path: Optional[Path] = None
    color_slots: list = field(default_factory=list)   # list[ColorSlotData]
    file_rows: list = field(default_factory=list)     # list[FileRowData]
    is_done: bool = False
    is_queued: bool = False
    has_collision: bool = False


@dataclass
class GpJobData:
    """Data model for a grandparent job (one theme/folder group)."""
    gp_path: str        # "ROOT_<path>" prefix for files with no parent folder
    gp_name: str        # display name of the grandparent folder
    parents: list = field(default_factory=list)   # list[PJobData]
    gp_rename_confirmed: bool = False
