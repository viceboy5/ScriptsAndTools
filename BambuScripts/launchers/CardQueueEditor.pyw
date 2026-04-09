# CardQueueEditor.pyw - Python launcher (replaces CardQueueEditorWPF.vbs)
#
# Associate .pyw files with pythonw.exe so double-clicking runs without a
# console window.  Drag-and-drop folders/files onto this file to pre-load them.
#
# Setup (one time):
#   pip install PySide6 Pillow
#
import sys
import os
from pathlib import Path

# Resolve the py workers package relative to this launcher
_SCRIPT_DIR = Path(__file__).resolve().parent
_PY_DIR = _SCRIPT_DIR.parent / 'workers' / 'py'

if str(_PY_DIR) not in sys.path:
    sys.path.insert(0, str(_PY_DIR))

# Also ensure the script's own directory is on PATH so relative imports work
os.chdir(str(_PY_DIR))

from app import main

if __name__ == '__main__':
    main(sys.argv[1:])
