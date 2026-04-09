# image_utils.py - Image manipulation utilities (pick-color randomizer, merge map)
import random
import tempfile
from pathlib import Path
from typing import Optional

try:
    from PIL import Image, ImageDraw
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


def randomize_pick_colors(src_path: Path, dst_path: Path) -> bool:
    """
    Load a pick-layer PNG, randomize each unique opaque color to a new random
    RGB value (preserving alpha), and save to dst_path.

    Mirrors PS1 Invoke-RandomizePickColors.
    Returns True on success, False if PIL is unavailable or an error occurs.
    """
    if not PIL_AVAILABLE:
        return False
    try:
        img = Image.open(src_path).convert('RGBA')
        pixels = img.load()
        color_map: dict[tuple, tuple] = {}
        rng = random.Random()
        w, h = img.size
        for y in range(h):
            for x in range(w):
                r, g, b, a = pixels[x, y]
                if a < 10:
                    continue
                key = (r, g, b)
                if key not in color_map:
                    color_map[key] = (
                        rng.randint(20, 255),
                        rng.randint(20, 255),
                        rng.randint(20, 255),
                    )
                nr, ng, nb = color_map[key]
                pixels[x, y] = (nr, ng, nb, a)
        img.save(dst_path)
        return True
    except Exception:
        return False


def randomize_pick_colors_from_bytes(raw_bytes: bytes) -> Optional[bytes]:
    """
    Randomize pick colors from in-memory PNG bytes, return result as PNG bytes.
    Returns None on failure.
    """
    if not PIL_AVAILABLE:
        return None
    try:
        import io
        buf_in = io.BytesIO(raw_bytes)
        img = Image.open(buf_in).convert('RGBA')
        pixels = img.load()
        color_map: dict[tuple, tuple] = {}
        rng = random.Random()
        w, h = img.size
        for y in range(h):
            for x in range(w):
                r, g, b, a = pixels[x, y]
                if a < 10:
                    continue
                key = (r, g, b)
                if key not in color_map:
                    color_map[key] = (
                        rng.randint(20, 255),
                        rng.randint(20, 255),
                        rng.randint(20, 255),
                    )
                nr, ng, nb = color_map[key]
                pixels[x, y] = (nr, ng, nb, a)
        buf_out = io.BytesIO()
        img.save(buf_out, format='PNG')
        return buf_out.getvalue()
    except Exception:
        return None


def get_merge_map_image_bytes(pre_bytes: bytes, post_bytes: bytes) -> Optional[bytes]:
    """
    Generate a merge-map visualization from two pick-layer PNG byte strings.

    Mirrors the PS1 Get-ImageBasedMergeMap / FastMergeMap logic:
    - Scans both images to build color -> first-occurrence-point anchor maps.
    - For each pre-anchor point, checks the same pixel in post.
    - If that post color is also a known post-anchor, draws a connecting line
      (black outline + color fill) with crosshairs at both ends.

    Returns the rendered PNG as bytes, or None on failure.
    """
    if not PIL_AVAILABLE:
        return None
    try:
        import io
        pre  = Image.open(io.BytesIO(pre_bytes)).convert('RGBA')
        post = Image.open(io.BytesIO(post_bytes)).convert('RGBA')

        pre_w,  pre_h  = pre.size
        post_w, post_h = post.size
        pre_px  = pre.load()
        post_px = post.load()

        # Build anchor maps: ARGB-int -> (x, y) of first occurrence
        pre_anchors: dict[tuple, tuple]  = {}
        post_anchors: dict[tuple, tuple] = {}

        for y in range(pre_h):
            for x in range(pre_w):
                r, g, b, a = pre_px[x, y]
                if a == 255 and (r > 10 or g > 10 or b > 10):
                    key = (r, g, b, a)
                    if key not in pre_anchors:
                        pre_anchors[key] = (x, y)
                if x < post_w and y < post_h:
                    r2, g2, b2, a2 = post_px[x, y]
                    if a2 == 255 and (r2 > 10 or g2 > 10 or b2 > 10):
                        key2 = (r2, g2, b2, a2)
                        if key2 not in post_anchors:
                            post_anchors[key2] = (x, y)

        result = post.copy()
        draw = ImageDraw.Draw(result)

        for old_pt in pre_anchors.values():
            ox, oy = old_pt
            if ox >= post_w or oy >= post_h:
                continue
            current_color = post_px[ox, oy]
            if current_color not in post_anchors:
                continue
            new_pt = post_anchors[current_color]
            r, g, b, a = current_color
            rgb = (r, g, b)
            black = (0, 0, 0)

            draw.line([old_pt, new_pt], fill=black, width=5)
            draw.line([old_pt, new_pt], fill=rgb, width=2)
            for pt in (old_pt, new_pt):
                px_x, px_y = pt
                draw.line([(px_x - 3, px_y), (px_x + 3, px_y)], fill=black, width=3)
                draw.line([(px_x, px_y - 3), (px_x, px_y + 3)], fill=black, width=3)
                draw.line([(px_x - 3, px_y), (px_x + 3, px_y)], fill=rgb, width=1)
                draw.line([(px_x, px_y - 3), (px_x, px_y + 3)], fill=rgb, width=1)

        buf = io.BytesIO()
        result.save(buf, format='PNG')
        return buf.getvalue()
    except Exception:
        return None


def bytes_to_qpixmap(png_bytes: bytes):
    """Convert raw PNG bytes to a QPixmap (requires PySide6)."""
    try:
        from PySide6.QtGui import QPixmap
        from PySide6.QtCore import QByteArray
        ba = QByteArray(png_bytes)
        pm = QPixmap()
        pm.loadFromData(ba, 'PNG')
        return pm
    except Exception:
        return None
