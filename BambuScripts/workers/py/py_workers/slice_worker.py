#!/usr/bin/env python3
"""
slice_worker.py - Invoke Bambu Studio CLI to slice one or more .3mf files.

Usage:
    python slice_worker.py --input-path <Full.3mf> [--isolated-path <Final.3mf>]
                           [--bambu-path "C:\\...\\bambu-studio.exe"]

Mirrors PS1 Slice_worker.ps1.
"""
import argparse
import subprocess
import sys
import tempfile
from pathlib import Path

BAMBU_DEFAULT = r'C:\Program Files\Bambu Studio\bambu-studio.exe'


def slice_file(file_path: Path, bambu_exe: Path, label: str = '') -> bool:
    """
    Invoke Bambu Studio CLI to slice a single .3mf file.
    Returns True on success (output gcode .3mf was created).
    """
    work_dir  = file_path.parent
    base_name = file_path.stem
    sliced_out = work_dir / f'{base_name}.gcode.3mf'

    print(f'  -> Slicing {label}: {base_name} ', end='', flush=True)

    args = [
        str(bambu_exe),
        '--debug', '3',
        '--no-check',
        '--uptodate',
        '--allow-newer-file',
        '--slice', '1',
        '--min-save',
        '--export-3mf', str(sliced_out),
        str(file_path),
    ]

    log_out = Path(tempfile.gettempdir()) / f'bbs_out_{base_name}.txt'
    log_err = Path(tempfile.gettempdir()) / f'bbs_err_{base_name}.txt'

    try:
        import time
        with open(log_out, 'w') as fout, open(log_err, 'w') as ferr:
            proc = subprocess.Popen(args, stdout=fout, stderr=ferr)
        while proc.poll() is None:
            print('.', end='', flush=True)
            time.sleep(3)
        print(f' [DONE] exit={proc.returncode}')
    except Exception as exc:
        print(f'\n  [!] ERROR: Failed to launch Bambu Studio: {exc}')
        return False

    if not sliced_out.exists():
        print(f'\n  [!] ERROR: Bambu Studio did not generate {sliced_out}')
        for lf, tag in ((log_out, 'STDOUT'), (log_err, 'STDERR')):
            if lf.exists():
                txt = lf.read_text(encoding='utf-8', errors='replace').strip()
                if txt:
                    print(f'  [Bambu {tag} (last 3000 chars)]\n{txt[-3000:]}')
                lf.unlink(missing_ok=True)
        return False

    for lf in (log_out, log_err):
        lf.unlink(missing_ok=True)
    return True


def main() -> int:
    parser = argparse.ArgumentParser(description='Slice 3MF files with Bambu Studio')
    parser.add_argument('--input-path',    default='', help='Primary 3MF to slice')
    parser.add_argument('--input-paths',   nargs='*', default=[], help='Multiple 3MF files')
    parser.add_argument('--isolated-path', default='', help='Optional isolated Final.3mf to also slice')
    parser.add_argument('--bambu-path',    default=BAMBU_DEFAULT, help='Path to bambu-studio.exe')
    args = parser.parse_args()

    bambu_exe = Path(args.bambu_path)
    if not bambu_exe.exists():
        print(f'  [!] ERROR: Bambu Studio not found at: {bambu_exe}')
        return 1

    # Collect all inputs
    all_inputs: list[Path] = []
    if args.input_path:
        all_inputs.append(Path(args.input_path))
    for p in (args.input_paths or []):
        if p:
            all_inputs.append(Path(p))

    if not all_inputs:
        print('  [!] ERROR: No input file specified.')
        return 1

    total = len(all_inputs)
    for i, f in enumerate(all_inputs, start=1):
        label = f'[{i}/{total}]' if total > 1 else 'Full Plate'
        slice_file(f, bambu_exe, label)

    if args.isolated_path:
        iso = Path(args.isolated_path)
        if iso.exists():
            slice_file(iso, bambu_exe, 'Isolated Object')

    print('  -> Slicing Automation Complete!\n')
    return 0


if __name__ == '__main__':
    sys.exit(main())
