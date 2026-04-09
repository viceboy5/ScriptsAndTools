# file_utils.py - File parsing, 3MF reading, and SmartFill logic
import re
import zipfile
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional

from models import (
    PRINTER_PREFIXES, ColorSlotData,
    FILE_COLOR_FINAL, FILE_COLOR_NEST, FILE_COLOR_FULL,
    FILE_COLOR_GCODE, FILE_COLOR_DEFAULT,
)

# Compound extensions that must be checked before Path.suffix
_COMPOUND_EXTS = ['.gcode.3mf', '.gcode.stl', '.gcode.step', '.f3d.3mf']

# Sort key for file rows (matches PS1 sort order)
_FILE_SORT_ORDER = {
    'final.3mf': 0,
    'nest.3mf': 1,
    'full.3mf': 2,
    'full.gcode.3mf': 3,
}


# ── Filename parsing ──────────────────────────────────────────────────────────

def parse_file(filename: str) -> dict:
    """
    Split a filename into stem, suffix, extension, and parts tokens.

    Handles compound extensions (.gcode.3mf etc.) before falling back to
    pathlib suffix detection.

    Returns: {Suffix, Extension, Stem, Parts}
    """
    ext = None
    fl = filename.lower()
    for ce in _COMPOUND_EXTS:
        if fl.endswith(ce):
            ext = filename[len(filename) - len(ce):]
            break
    if ext is None:
        p = Path(filename)
        ext = p.suffix or ''

    stem = filename[:len(filename) - len(ext)]
    parts = [t for t in re.split(r'[\s._-]+', stem) if t]

    if ext.lower() == '.png':
        return {'Suffix': '', 'Extension': ext, 'Stem': stem, 'Parts': parts}

    suffix = parts[-1] if parts else stem
    return {'Suffix': suffix, 'Extension': ext, 'Stem': stem, 'Parts': parts}


def file_base_color(filename: str) -> str:
    """Return the display hex badge color for a given filename."""
    fl = filename.lower()
    if fl.endswith('nest.3mf'):
        return FILE_COLOR_NEST
    if fl.endswith('full.gcode.3mf'):
        return FILE_COLOR_GCODE
    if fl.endswith('full.3mf'):
        return FILE_COLOR_FULL
    if fl.endswith('final.3mf'):
        return FILE_COLOR_FINAL
    return FILE_COLOR_DEFAULT


def file_sort_key(filename: str) -> tuple:
    """Sort key matching the PS1 switch-Regex order."""
    fl = filename.lower()
    for pat, order in _FILE_SORT_ORDER.items():
        if fl.endswith(pat):
            return (order, fl)
    return (4, fl)


# ── SmartFill ─────────────────────────────────────────────────────────────────

def smart_fill(anchor_name: str, gp_name: str) -> dict:
    """
    Parse character name and adjective from an anchor filename + grandparent name.

    Mirrors the PS1 SmartFill function exactly.
    Returns: {'Char': str, 'Adj': str}
    """
    stem = re.sub(r'(?i)\.gcode\.3mf$|\.3mf$', '', anchor_name)
    stem = re.sub(r'(?i)_Full$', '', stem)

    # Strip printer prefix from grandparent name
    clean_gp = gp_name
    if gp_name:
        gp_parts = gp_name.split('_')
        if len(gp_parts) > 1 and gp_parts[0].upper() in [p.upper() for p in PRINTER_PREFIXES]:
            clean_gp = '_'.join(gp_parts[1:])

    # Try to strip the theme from the end of the stem
    theme_stripped = False
    prefix = stem
    if clean_gp:
        m = re.match(rf'^(.+)_{re.escape(clean_gp)}$', stem, re.IGNORECASE)
        if m:
            prefix = m.group(1)
            theme_stripped = True

    # Strip known printer prefix from prefix
    pref_parts = prefix.split('_')
    if len(pref_parts) > 1 and pref_parts[0].upper() in [p.upper() for p in PRINTER_PREFIXES]:
        prefix = '_'.join(pref_parts[1:])

    parts = [t for t in prefix.split('_') if t]

    # Positional fallback: if theme wasn't matched via gp_name and 3+ tokens remain,
    # treat the last token as the theme and strip it.
    if not theme_stripped and len(parts) >= 3:
        parts = parts[:-1]

    if len(parts) >= 2:
        return {'Char': parts[0], 'Adj': ''.join(parts[1:])}
    return {'Char': ''.join(parts), 'Adj': ''}


# ── 3MF extraction ────────────────────────────────────────────────────────────

def extract_3mf_to_temp(three_mf_path: Path) -> Path:
    """
    Extract a 3MF (ZIP) file to a new temp directory.
    The CALLER is responsible for cleanup (shutil.rmtree).
    """
    temp_dir = Path(tempfile.mkdtemp(prefix='LiveCard_'))
    with zipfile.ZipFile(three_mf_path, 'r') as zf:
        zf.extractall(temp_dir)
    return temp_dir


def read_3mf_colors(temp_work: Path, hex_to_name: dict) -> list:
    """
    Read color slot data from an extracted 3MF temp directory.

    Mirrors the PS1 Build-PJob color-parsing logic:
      1. Parse filament_colour list from project_settings.config
      2. Collect used extruder slots from model_settings.config
      3. Collect used materialid values from 3dmodel.model
      4. Return ColorSlotData for slots that are actually used (max 4)

    Args:
        temp_work:   Path to the extracted 3MF temp directory.
        hex_to_name: Dict mapping "#RRGGBBFF" / "#RRGGBB" -> color name.

    Returns list[ColorSlotData]
    """
    proj_path    = temp_work / 'Metadata' / 'project_settings.config'
    mod_set_path = temp_work / 'Metadata' / 'model_settings.config'

    slot_map: dict[str, str] = {}       # "#RRGGBB(FF)" -> slot_index (1-based str)
    used_slots: set[str] = {'1'}        # slot "1" is always considered used

    # 1. Parse filament colors from project_settings.config
    if proj_path.exists():
        try:
            content = proj_path.read_text(encoding='utf-8', errors='replace')
            m = re.search(r'(?is)"filament_colou?r"\s*:\s*\[(.*?)\]', content)
            if m:
                hexes = re.findall(r'#[0-9a-fA-F]{6,8}', m.group(1))
                for i, hx in enumerate(hexes, start=1):
                    hk = hx.upper()
                    if hk not in slot_map:
                        slot_map[hk] = str(i)
        except Exception:
            pass

    # 2. Parse used extruder slots from model_settings.config
    if mod_set_path.exists():
        try:
            tree = ET.parse(mod_set_path)
            for node in tree.iter():
                key = node.get('key') or ''
                if 'extruder' in key.lower():
                    val = (node.get('value') or '').strip()
                    if val:
                        used_slots.add(val)
        except Exception:
            pass

    # 3. Check 3dmodel.model for materialid assignments
    model_files = list(temp_work.rglob('3dmodel.model'))
    if model_files:
        try:
            content = model_files[0].read_text(encoding='utf-8', errors='replace')
            for m in re.finditer(r'(?i)materialid="(\d+)"', content):
                used_slots.add(m.group(1))
        except Exception:
            pass

    # 4. Build active slot list
    active: list[ColorSlotData] = []
    for hex_key, slot_id in slot_map.items():
        if slot_id in used_slots:
            check_hex = hex_key if len(hex_key) == 9 else hex_key + 'FF'
            matched = hex_to_name.get(check_hex, hex_to_name.get(check_hex[:7], ''))
            active.append(ColorSlotData(orig_hex=check_hex, name=matched))

    return active[:4]


def read_3mf_images(zip_path: Path) -> tuple[Optional[bytes], Optional[bytes]]:
    """
    Read plate_1.png and pick_1.png bytes from a 3MF zip.
    Falls back to thumbnail.png if plate_1.png is absent.
    Returns (plate_bytes, pick_bytes); either may be None.
    """
    plate_bytes: Optional[bytes] = None
    pick_bytes: Optional[bytes] = None
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            # Build a normalised name -> original-name map
            names = {e.replace('\\', '/').lower(): e for e in zf.namelist()}

            for candidate in ('metadata/plate_1.png', 'metadata/thumbnail.png'):
                if candidate in names:
                    plate_bytes = zf.read(names[candidate])
                    break

            for norm, orig in names.items():
                if norm.endswith('pick_1.png'):
                    pick_bytes = zf.read(orig)
                    break
    except Exception:
        pass
    return plate_bytes, pick_bytes


def detect_printer_prefix(name: str) -> str:
    """
    Return the printer prefix embedded in a folder/file name, or ''.
    E.g. 'P2S_Puppies' -> 'P2S', 'X1C_BabyDragon_Spicy_Ocean_01' -> 'X1C'
    """
    parts = name.split('_')
    if parts and parts[0].upper() in [p.upper() for p in PRINTER_PREFIXES]:
        return parts[0].upper()
    return ''
