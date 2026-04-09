# app.py - Entry point for CardQueueEditor Python application
"""
CardQueueEditor - Python/PySide6 conversion of CardQueueEditorWPF.ps1

Usage:
    python app.py [folder_or_file ...]
    pythonw CardQueueEditor.pyw [folder_or_file ...]   # no console window
"""
import sys
from pathlib import Path


def main(drop_args: list[str] | None = None) -> None:
    from PySide6.QtWidgets import QApplication
    from PySide6.QtCore import Qt

    # High-DPI support
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    app = QApplication(sys.argv)
    app.setApplicationName('CardQueueEditor')
    app.setOrganizationName('BambuScripts')

    # Import here so Qt is initialized first
    from main_window import MainWindow

    window = MainWindow()
    window.show()

    # Load drag-and-dropped / command-line paths
    if drop_args:
        paths = [Path(a) for a in drop_args if a]
        window.load_paths(paths)

    sys.exit(app.exec())


if __name__ == '__main__':
    main(sys.argv[1:])
