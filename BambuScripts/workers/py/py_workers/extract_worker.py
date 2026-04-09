#!/usr/bin/env python3
"""
extract_worker.py - Parse gcode from a .gcode.3mf and write a TSV data row.

Usage:
    python extract_worker.py --input-file <Full.gcode.3mf>
                             [--single-file <Final.gcode.3mf>]
                             [--individual-tsv-path <output_Data.tsv>]
                             [--generate-image]
                             [--skip-extraction]

Mirrors PS1 DataExtract_worker.ps1.
"""
from __future__ import annotations

import argparse
import io
import re
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path


# ── Color CSV loader ──────────────────────────────────────────────────────────

def load_color_library(csv_path: Path) -> tuple[dict[str, str], dict[str, str]]:
    """
    Returns (library_names, name_to_hex):
      library_names: "RRGGBBFF" -> color name
      name_to_hex:   color name -> "#RRGGBBFF"
    """
    library_names: dict[str, str] = {}
    name_to_hex: dict[str, str] = {}
    if not csv_path.exists():
        return library_names, name_to_hex
    for line in csv_path.read_text(encoding='utf-8', errors='replace').splitlines():
        line = line.strip()
        if not line:
            continue
        parts = line.split(',')
        if len(parts) < 4:
            continue
        name = parts[0].replace('"', '').strip()
        if not name or name.lower() == 'name' or name in ('N/A', ''):
            continue
        try:
            r, g, b = int(parts[1].strip()), int(parts[2].strip()), int(parts[3].strip())
        except ValueError:
            continue
        raw_hex = f'{r:02X}{g:02X}{b:02X}FF'
        library_names[raw_hex] = name
        name_to_hex[name] = '#' + raw_hex
    return library_names, name_to_hex


# ── Gcode analyzer ────────────────────────────────────────────────────────────

class GcodeAnalyzer:
    """Parse gcode stream for print time, flush grams, tower grams, color changes."""

    def __init__(self) -> None:
        self.flush_grams:   float = 0.0
        self.tower_grams:   float = 0.0
        self.print_time:    str   = 'Not found'
        self.color_changes: int   = 0

    def analyze(self, gcode_bytes: bytes, grams_per_mm: float) -> None:
        flush_e: float = 0.0
        tower_e: float = 0.0
        in_flush = False
        in_tower = False
        is_relative = False
        current_e: float = 0.0
        max_e: float = 0.0

        for raw_line in io.BytesIO(gcode_bytes):
            try:
                line = raw_line.decode('utf-8', errors='replace').strip()
            except Exception:
                continue
            if not line:
                continue

            c = line[0]

            if c == 'G':
                if line.startswith('G1 ') or line.startswith('G0 '):
                    idx = line.find(' E')
                    if idx > -1:
                        start = idx + 2
                        end = line.find(' ', start)
                        if end == -1: end = line.find(';', start)
                        if end == -1: end = len(line)
                        try:
                            e_val = float(line[start:end])
                            if is_relative:
                                current_e += e_val
                            else:
                                current_e = e_val
                            if current_e > max_e:
                                delta = current_e - max_e
                                if in_flush: flush_e += delta
                                elif in_tower: tower_e += delta
                                max_e = current_e
                        except ValueError:
                            pass
                elif line.startswith('G92'):
                    idx = line.find(' E')
                    if idx > -1:
                        start = idx + 2
                        end = line.find(' ', start)
                        if end == -1: end = line.find(';', start)
                        if end == -1: end = len(line)
                        try:
                            e_val = float(line[start:end])
                            current_e = e_val
                            max_e = e_val
                        except ValueError:
                            pass
                    else:
                        current_e = 0.0
                        max_e = 0.0

            elif c == ';':
                if line.startswith('; FLUSH_START'):      in_flush = True
                elif line.startswith('; FLUSH_END'):       in_flush = False
                elif line.startswith('; WIPE_TOWER_START'): in_tower = True
                elif line.startswith('; WIPE_TOWER_END'):   in_tower = False
                elif line.startswith('; TYPE:'):
                    in_tower = ('Wipe tower' in line or 'Prime tower' in line)
                elif 'total estimated time:' in line:
                    idx = line.index('total estimated time:')
                    pt = line[idx + 21:].strip()
                    semi = pt.find(';')
                    if semi > -1: pt = pt[:semi].strip()
                    self.print_time = pt
                elif 'estimated printing time' in line and self.print_time == 'Not found':
                    idx = line.find('=')
                    if idx > -1:
                        pt = line[idx + 1:].strip()
                        semi = pt.find(';')
                        if semi > -1: pt = pt[:semi].strip()
                        self.print_time = pt

            elif c == 'M':
                if line.startswith('M83'):   is_relative = True
                elif line.startswith('M82'): is_relative = False
                elif line.startswith('M620 S'): self.color_changes += 1

        self.flush_grams = flush_e * grams_per_mm
        self.tower_grams = tower_e * grams_per_mm


# ── Time parser ───────────────────────────────────────────────────────────────

def parse_print_time(print_time: str) -> tuple[int, int]:
    """Parse a print time string like '1h 23m 45s' -> (total_hours, minutes)."""
    d = int(m.group(1)) if (m := re.search(r'(\d+)d', print_time)) else 0
    h = int(m.group(1)) if (m := re.search(r'(\d+)h', print_time)) else 0
    mi = int(m.group(1)) if (m := re.search(r'(\d+)m', print_time)) else 0
    s = int(m.group(1)) if (m := re.search(r'(\d+)s', print_time)) else 0
    h += d * 24
    mi += 1 if s >= 30 else 0
    if mi >= 60: mi -= 60; h += 1
    return h, mi


# ── Gcode.3mf reader ─────────────────────────────────────────────────────────

def analyze_gcode_3mf(
    zip_path: Path,
    library_names: dict[str, str],
) -> tuple[list[dict], float, float, int, int, int, GcodeAnalyzer]:
    """
    Open a .gcode.3mf file and extract filament data + gcode stats.

    Returns:
        fil_data      - list of 5 dicts {g, color, raw_hex} (index 0 unused)
        total_grams
        grams_per_mm
        h, mi         - print hours and minutes
        obj_count
        analyzer      - GcodeAnalyzer instance
    """
    fil_data = [{'g': 0.0, 'color': '', 'raw_hex': ''} for _ in range(5)]
    total_grams = 0.0
    total_meters = 0.0
    obj_count = 0

    with zipfile.ZipFile(zip_path, 'r') as zf:
        names = {e.replace('\\', '/').lower(): e for e in zf.namelist()}

        # slice_info.config
        config_key = next((k for k in names if k.endswith('metadata/slice_info.config')), None)
        if config_key:
            config_bytes = zf.read(names[config_key])
            config_text = config_bytes.decode('utf-8', errors='replace')

            try:
                xml_root = ET.fromstring(config_text)
                active_slot = 1
                for node in xml_root.iter('filament'):
                    weight_str = node.get('used_g', '0')
                    try:
                        weight = float(weight_str)
                    except ValueError:
                        weight = 0.0
                    if weight > 0 and active_slot <= 4:
                        fil_data[active_slot]['g'] = round(weight, 2)
                        raw_hex = node.get('color', '').replace('"', '').strip().upper().lstrip('#')
                        if len(raw_hex) == 6:
                            raw_hex += 'FF'
                        fil_data[active_slot]['raw_hex'] = '#' + raw_hex
                        if raw_hex in library_names:
                            fil_data[active_slot]['color'] = library_names[raw_hex]
                        else:
                            fil_data[active_slot]['color'] = '#' + raw_hex
                        active_slot += 1

                # Object count (strict pre-merge logic)
                for obj in xml_root.findall('.//plate/object'):
                    obj_name = obj.get('name', '')
                    if re.search(r'(?i)text|version', obj_name):
                        continue
                    m = re.search(r'MergedGroup_(\d+)$', obj_name)
                    obj_count += int(m.group(1)) if m else 1
            except ET.ParseError:
                pass

            # Total grams/meters from regex
            for m in re.finditer(r'used_g="([0-9.]+)"', config_text):
                total_grams += float(m.group(1))
            for m in re.finditer(r'used_m="([0-9.]+)"', config_text):
                total_meters += float(m.group(1))

        grams_per_mm = total_grams / (total_meters * 1000) if total_meters > 0 else 0.0

        # gcode stream
        gcode_key = next((k for k in names if k.endswith('.gcode')), None)
        analyzer = GcodeAnalyzer()
        if gcode_key:
            gcode_bytes = zf.read(names[gcode_key])
            analyzer.analyze(gcode_bytes, grams_per_mm)

    model_grams = total_grams - analyzer.flush_grams - analyzer.tower_grams
    h, mi = parse_print_time(analyzer.print_time)
    return fil_data, total_grams, grams_per_mm, h, mi, obj_count, analyzer


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(description='Extract gcode stats to TSV')
    parser.add_argument('--input-file',          required=True)
    parser.add_argument('--single-file',         default='')
    parser.add_argument('--individual-tsv-path', default='')
    parser.add_argument('--generate-image',      action='store_true')
    parser.add_argument('--skip-extraction',     action='store_true')
    args = parser.parse_args()

    input_file = Path(args.input_file)
    script_dir = Path(__file__).resolve().parent.parent   # workers/py/

    # Load color library
    csv_path = script_dir / 'colorNamesCSV.csv'
    library_names, name_to_hex = load_color_library(csv_path)

    # Derive project name
    project_name = re.sub(r'(?i)\.gcode\.3mf$', '', input_file.name)
    project_name = re.sub(r'(?i)_Full$', '', project_name)

    # Parse theme / rarity from filename
    name_parts = project_name.split('_')
    theme_slot = name_parts[-1] if len(name_parts) >= 2 else ''
    adj_slot   = name_parts[-2] if len(name_parts) >= 3 else ''
    theme_output = ('RARE ' + theme_slot) if re.search(r'(?i)Rare|Legendary|Epic', adj_slot) else theme_slot

    fil_data = [{'g': 0.0, 'color': '', 'raw_hex': ''} for _ in range(5)]
    h = mi = obj_count = actual_color_swaps = 0
    model_grams = 0.0
    time_add: float = 0.0
    single_print_time_str = 'N/A'
    analyzer = GcodeAnalyzer()

    if not args.skip_extraction:
        try:
            fil_data, total_grams, grams_per_mm, h, mi, obj_count, analyzer = \
                analyze_gcode_3mf(input_file, library_names)
            model_grams = total_grams - analyzer.flush_grams - analyzer.tower_grams
            total_minutes = h * 60 + mi
            actual_color_swaps = max(0, analyzer.color_changes - 1)

            # Single file time
            if args.single_file and Path(args.single_file).exists():
                try:
                    _, _, _, sh, sm, _, s_analyzer = analyze_gcode_3mf(Path(args.single_file), library_names)
                    single_print_time_str = s_analyzer.print_time
                    single_total = sh * 60 + sm
                    if obj_count > 1:
                        time_add = round((total_minutes - single_total) / (obj_count - 1), 2)
                except Exception:
                    pass
        except Exception as exc:
            print(f'[extract] ERROR: {exc}', file=sys.stderr)

    else:
        # Fast TSV-only mode
        tsv_path = Path(args.individual_tsv_path)
        if tsv_path.exists():
            try:
                last = tsv_path.read_text(encoding='utf-8', errors='replace').splitlines()[-1]
                cols = last.split('\t')
                if len(cols) >= 19:
                    for i in range(1, 5):
                        g_idx = 5 + (i - 1) * 2
                        c_idx = g_idx + 1
                        try:
                            fil_data[i]['g'] = float(cols[g_idx])
                        except ValueError:
                            pass
                        fil_data[i]['color'] = cols[c_idx]
                        if fil_data[i]['color'].startswith('#'):
                            fil_data[i]['raw_hex'] = fil_data[i]['color']
                        elif fil_data[i]['color'] in name_to_hex:
                            fil_data[i]['raw_hex'] = name_to_hex[fil_data[i]['color']]
                    try:
                        time_add = float(cols[18])
                    except ValueError:
                        pass
            except Exception:
                pass

    # Build TSV row
    today = date.today().strftime('%-m/%-d/%Y') if sys.platform != 'win32' else date.today().strftime('%#m/%#d/%Y')
    row_values = [
        project_name, theme_output, today,
        h, mi,
        fil_data[1]['g'] if fil_data[1]['g'] > 0 else 0, fil_data[1]['color'],
        fil_data[2]['g'] if fil_data[2]['g'] > 0 else 0, fil_data[2]['color'],
        fil_data[3]['g'] if fil_data[3]['g'] > 0 else 0, fil_data[3]['color'],
        fil_data[4]['g'] if fil_data[4]['g'] > 0 else 0, fil_data[4]['color'],
        actual_color_swaps, obj_count, round(model_grams, 2),
        '=SUM(INDIRECT("G"&ROW()&":N"&ROW()))',
        time_add,
    ]
    tsv_row = '\t'.join(str(v) for v in row_values)

    if args.individual_tsv_path:
        out_path = Path(args.individual_tsv_path)
        out_path.write_text(tsv_row, encoding='utf-8')
        print(f'[extract] TSV written -> {out_path.name}')

    # Image generation
    if args.generate_image:
        _generate_image(input_file, script_dir, project_name, fil_data, time_add)

    return 0


def _generate_image(
    input_file: Path,
    script_dir: Path,
    project_name: str,
    fil_data: list[dict],
    time_add: float,
) -> None:
    py_script = script_dir / 'generate_image_worker.py'
    if not py_script.exists():
        print('[extract] generate_image_worker.py not found - skipping image', file=sys.stderr)
        return

    input_folder = input_file.parent
    py_base = re.sub(r'(?i)[ ._-]Full$', '', project_name)
    expected_png = input_folder / f'{py_base}_slicePreview.png'

    # Find source image
    pngs = [p for p in input_folder.glob('*.png') if not p.name.endswith('_slicePreview.png')]
    source_img: Path | None = pngs[0] if pngs else None
    is_temp = False

    if source_img is None and input_file.exists():
        try:
            with zipfile.ZipFile(input_file, 'r') as zf:
                names = {e.replace('\\', '/').lower(): e for e in zf.namelist()}
                plate_key = next((k for k in names if k.endswith('metadata/plate_1.png')), None)
                if plate_key:
                    tmp = tempfile.NamedTemporaryFile(suffix='_plate_1.png', delete=False)
                    tmp.write(zf.read(names[plate_key]))
                    tmp.close()
                    source_img = Path(tmp.name)
                    is_temp = True
        except Exception:
            pass

    if source_img is None or not source_img.exists():
        print('[extract] No source image found - skipping image generation')
        return

    args = [
        sys.executable, str(py_script),
        '--name', project_name,
        '--time', str(time_add),
        '--img',  str(source_img),
        '--out',  str(expected_png),
        '--colors',
    ]
    for i in range(1, 5):
        if fil_data[i]['g'] > 0:
            args.append(f"{fil_data[i]['color']}|{fil_data[i]['raw_hex']}|{fil_data[i]['g']}")

    print('  -> Generating Composite Card... ', end='', flush=True)
    try:
        result = subprocess.run(args, capture_output=True, timeout=60)
        if expected_png.exists():
            print('[DONE]')
        else:
            print('[FAILED]')
            if result.stderr:
                print(result.stderr.decode(errors='replace'), file=sys.stderr)
    except Exception as exc:
        print(f'[CRASHED]: {exc}')
    finally:
        if is_temp and source_img and source_img.exists():
            source_img.unlink(missing_ok=True)


if __name__ == '__main__':
    sys.exit(main())
