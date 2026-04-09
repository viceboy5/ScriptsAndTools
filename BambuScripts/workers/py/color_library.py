# color_library.py - Color CSV loading and hex/name lookup
from pathlib import Path
from typing import Optional


class ColorLibrary:
    """
    Loads colorNamesCSV.csv and provides hex <-> name lookups.

    CSV format (no header required, header row is skipped):
        Name, R, G, B [, optional alt hex columns ...]
    """

    def __init__(self, csv_path: Path):
        self.library_colors: dict[str, str] = {}   # name  -> "#RRGGBBFF"
        self.hex_to_name: dict[str, str] = {}       # "#RRGGBBFF" or "#RRGGBB" -> name
        if csv_path.exists():
            self._load(csv_path)

    # ── Loading ───────────────────────────────────────────────────────────────

    def _load(self, csv_path: Path) -> None:
        try:
            text = csv_path.read_text(encoding='utf-8', errors='replace')
        except OSError:
            return

        for line in text.splitlines():
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
                r = int(parts[1].replace('"', '').strip())
                g = int(parts[2].replace('"', '').strip())
                b = int(parts[3].replace('"', '').strip())
            except (ValueError, IndexError):
                continue

            hex8 = f'#{r:02X}{g:02X}{b:02X}FF'
            hex6 = f'#{r:02X}{g:02X}{b:02X}'

            self.library_colors[name] = hex8
            self.hex_to_name[hex8] = name
            self.hex_to_name[hex6] = name

    # ── Lookups ───────────────────────────────────────────────────────────────

    def name_for_hex(self, hex_color: str) -> str:
        """Return the library name for a hex string, or '' if not found."""
        h = hex_color.upper()
        return self.hex_to_name.get(h, self.hex_to_name.get(h[:7], ''))

    def hex_for_name(self, name: str) -> Optional[str]:
        """Return the '#RRGGBBFF' hex for a name, or None if not in library."""
        return self.library_colors.get(name)

    def all_names(self) -> list[str]:
        """Return all color names in load order."""
        return list(self.library_colors.keys())

    def contains_name(self, name: str) -> bool:
        return name in self.library_colors

    def contains_hex(self, hex_color: str) -> bool:
        h = hex_color.upper()
        return h in self.hex_to_name or h[:7] in self.hex_to_name

    def __len__(self) -> int:
        return len(self.library_colors)
