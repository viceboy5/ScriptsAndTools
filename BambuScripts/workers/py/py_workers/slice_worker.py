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

    try:
        proc = subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        while proc.poll() is None:
            print('.', end='', flush=True)
            import time; time.sleep(3)
        print(' [DONE]')
    except Exception as exc:
        print(f'\n  [!] ERROR: Failed to launch Bambu Studio: {exc}')
        return False

    if not sliced_out.exists():
        print(f'\n  [!] ERROR: Bambu Studio did not generate {sliced_out}')
        return False

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
