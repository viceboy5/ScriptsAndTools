import argparse
import os
import re
import math
import csv
from PIL import Image, ImageDraw, ImageFont, ImageChops, ImageFilter

# --- DYNAMIC CONFIGURATION ---
CANVAS_SIZE = 512

# Layout Ratios
MARGIN_RATIO = 0.01
COLOR_BOX_SIZE_RATIO = 0.155
FONT_TITLE_RATIO = 0.09
FONT_NUM_RATIO = 0.06
FONT_TEXT_RATIO = 0.043
FONT_MASS_RATIO = 0.039
FONT_TIME_RATIO = 0.075

MARGIN = int(CANVAS_SIZE * MARGIN_RATIO)
COLOR_BOX_SIZE = int(CANVAS_SIZE * COLOR_BOX_SIZE_RATIO)


def load_font(size):
    for font_name in ["comicbd.ttf", "ariblk.ttf", "arialbd.ttf", "arial.ttf"]:
        try:
            return ImageFont.truetype(font_name, size)
        except:
            continue
    return ImageFont.load_default()


def parse_hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    if len(hex_str) >= 6:
        return tuple(int(hex_str[i:i + 2], 16) for i in (0, 2, 4))
    return (255, 255, 255)


def is_color_light(rgb):
    r, g, b = rgb
    return (0.299 * r + 0.587 * g + 0.114 * b) > 186


def draw_text_with_outline(draw, pos, text, font, fill, outline_color=(0, 0, 0), outline_ratio=0.004):
    x, y = pos
    outline_width = max(1, int(CANVAS_SIZE * outline_ratio))

    for dx in range(-outline_width, outline_width + 1):
        for dy in range(-outline_width, outline_width + 1):
            if dx != 0 or dy != 0:
                draw.text((x + dx, y + dy), text, font=font, fill=outline_color)
    draw.text((x, y), text, font=font, fill=fill)


def load_gradient_library(csv_filename="colorNamesCSV.csv"):
    """Reads the CSV and maps the base RGB tuple to a list of gradient hexes."""
    gradient_map = {}
    script_dir = os.path.dirname(os.path.abspath(__file__))
    csv_path = os.path.join(script_dir, csv_filename)

    if os.path.exists(csv_path):
        with open(csv_path, mode='r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            try:
                next(reader)  # Skip the header row
            except StopIteration:
                return gradient_map

            for row in reader:
                # Check if there are columns beyond R, G, B
                if len(row) > 4:
                    try:
                        # Grab the base RGB tuple to use as the lookup key
                        base_rgb = (int(row[1].strip()), int(row[2].strip()), int(row[3].strip()))

                        # Dynamically capture all remaining columns that contain a hex code
                        gradients = [cell.strip() for cell in row[4:] if cell.strip().startswith('#')]

                        if len(gradients) >= 2:
                            gradient_map[base_rgb] = gradients
                    except ValueError:
                        continue  # Skip rows with missing or malformed RGB data
    return gradient_map


def create_gradient_swatch(width, height, hex_colors):
    """Draws a pixel-perfect linear gradient from an infinite list of hexes."""
    base = Image.new('RGB', (width, height))
    draw = ImageDraw.Draw(base)

    rgb_colors = []
    for h in hex_colors:
        h = h.strip().lstrip('#')
        rgb_colors.append(tuple(int(h[i:i + 2], 16) for i in (0, 2, 4)))

    segment_width = width / (len(rgb_colors) - 1)

    for i in range(len(rgb_colors) - 1):
        color1, color2 = rgb_colors[i], rgb_colors[i + 1]
        start_x = int(i * segment_width)
        end_x = int((i + 1) * segment_width) if i < len(rgb_colors) - 2 else width

        for x in range(start_x, end_x):
            ratio = (x - start_x) / max(1, (end_x - start_x))
            r = int(color1[0] * (1 - ratio) + color2[0] * ratio)
            g = int(color1[1] * (1 - ratio) + color2[1] * ratio)
            b = int(color1[2] * (1 - ratio) + color2[2] * ratio)
            draw.line([(x, 0), (x, height)], fill=(r, g, b))

    return base


def draw_gradient_text(layer, pos, text, font, gradient_colors, outline_color=(0, 0, 0), outline_ratio=0.004):
    """Draws text filled with a horizontal gradient (used for RARE highlight)."""
    x, y = pos
    outline_width = max(1, int(CANVAS_SIZE * outline_ratio))
    draw = ImageDraw.Draw(layer)
    for dx in range(-outline_width, outline_width + 1):
        for dy in range(-outline_width, outline_width + 1):
            if dx != 0 or dy != 0:
                draw.text((x + dx, y + dy), text, font=font, fill=outline_color)
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2], bbox[3]
    if tw <= 0 or th <= 0:
        return
    mask_img = Image.new("L", (tw, th), 0)
    ImageDraw.Draw(mask_img).text((-bbox[0], -bbox[1]), text, font=font, fill=255)
    grad_rgba = create_gradient_swatch(tw, th, gradient_colors).convert("RGBA")
    grad_rgba.putalpha(mask_img)
    layer.paste(grad_rgba, (x, y + bbox[1]), grad_rgba)


def draw_metallic_border(draw, box_x, y, size, base_rgb, border_width=4):
    """Multi-layer metallic border derived from the filament color."""
    r, g, b = base_rgb
    highlight = (min(255,int(r*0.5+200)), min(255,int(g*0.5+200)), min(255,int(b*0.5+200)))
    midtone   = (min(255,int(r*0.7+80)),  min(255,int(g*0.7+80)),  min(255,int(b*0.7+80)))
    shadow    = (max(0,int(r*0.3)),        max(0,int(g*0.3)),        max(0,int(b*0.3)))
    draw.rectangle([box_x, y, box_x+size, y+size], outline=midtone, width=border_width)
    o = border_width // 2
    draw.line([box_x+o, y+o, box_x+size-o, y+o],          fill=highlight, width=max(1,border_width//2))
    draw.line([box_x+o, y+o, box_x+o, y+size-o],          fill=highlight, width=max(1,border_width//2))
    draw.line([box_x+o, y+size-o, box_x+size-o, y+size-o],fill=shadow,    width=max(1,border_width//2))
    draw.line([box_x+size-o, y+o, box_x+size-o, y+size-o],fill=shadow,    width=max(1,border_width//2))


def create_silk_swatch(width, height, base_rgb):
    """Diagonal sheen concentrated at corners, base color dominant in center.
    Top-left corner brightens to highlight, bottom-right darkens to shadow.
    Squared falloff keeps the center close to base so numbers stay readable."""
    r, g, b = base_rgb
    highlight = (min(255,int(r*0.55+200)), min(255,int(g*0.55+200)), min(255,int(b*0.55+200)))
    shadow    = (max(0,int(r*0.35)),       max(0,int(g*0.35)),       max(0,int(b*0.35)))
    img = Image.new("RGB", (width, height))
    pixels = img.load()
    for px in range(width):
        for py in range(height):
            diag = (px / width + py / height) / 2.0
            t = (diag - 0.5) * 2.0
            t = t * abs(t)  # squared falloff: strong at edges, weak at centre
            if t < 0:
                s = -t
                pr = int(r*(1-s) + highlight[0]*s)
                pg = int(g*(1-s) + highlight[1]*s)
                pb = int(b*(1-s) + highlight[2]*s)
            else:
                s = t
                pr = int(r*(1-s) + shadow[0]*s)
                pg = int(g*(1-s) + shadow[1]*s)
                pb = int(b*(1-s) + shadow[2]*s)
            pixels[px, py] = (pr, pg, pb)
    return img


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--name", required=True)
    parser.add_argument("--time", required=True)
    parser.add_argument("--img", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--colors", nargs='*', default=[])
    args = parser.parse_args()

    gradient_library = load_gradient_library()

    # --- 1. CLEAN NAME LOGIC ---
    original_name = args.name

    # Strip _Full suffix (may already be stripped by caller, but safe to repeat)
    name_stripped = re.sub(r'(?i)[ ._-]Full$', '', original_name)

    # Parse naming structure: Character[_Adj]_Theme
    #   enforced naming means no underscores inside Character or Adj (CamelCase only)
    #   so: 3 parts = Character, Adj, Theme | 2 parts = Character, Theme | 1 part = Character
    parts = name_stripped.split('_')
    if len(parts) >= 3:
        character_raw = parts[0]
        adj_raw       = '_'.join(parts[1:-1])   # handles edge case of extra underscores
        theme_raw     = parts[-1]
    elif len(parts) == 2:
        character_raw = parts[0]
        adj_raw       = ''
        theme_raw     = parts[1]
    else:
        character_raw = parts[0]
        adj_raw       = ''
        theme_raw     = ''

    # CamelCase → space-separated for Character  (e.g. DragonHatchling → Dragon Hatchling)
    char_display = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', character_raw).upper()
    adj_display  = adj_raw.upper()

    # Build the base filename (used for output path) - keep original stripped name
    base_filename  = name_stripped
    output_filename = f"{base_filename}_slicePreview.png"

    # RESTORED: This is the line that was missing!
    output_path = os.path.join(os.path.dirname(args.out), output_filename)

    # --- 2. INITIALIZE CANVAS ---
    background = Image.new("RGBA", (CANVAS_SIZE, CANVAS_SIZE), (0, 0, 0, 255))
    ui_layer   = Image.new("RGBA", (CANVAS_SIZE, CANVAS_SIZE), (0, 0, 0, 0))
    ui_draw    = ImageDraw.Draw(ui_layer)

    # --- 3. DRAW ALL UI ELEMENTS TO THE TRANSPARENT LAYER ---

    # A. TITLE TEXT
    # Build the display string for sizing: "CHAR PART (ADJ)" or just "CHAR PART"
    display_for_sizing = char_display + (f" ({adj_display})" if adj_display else "")

    current_font_size = int(CANVAS_SIZE * FONT_TITLE_RATIO)
    font_title = load_font(current_font_size)
    max_title_width = CANVAS_SIZE - (MARGIN * 2)

    while ui_draw.textbbox((0, 0), display_for_sizing, font=font_title)[2] > max_title_width and current_font_size > int(
            CANVAS_SIZE * 0.03):
        current_font_size -= 2
        font_title = load_font(current_font_size)

    bbox_title = ui_draw.textbbox((0, 0), display_for_sizing, font=font_title)

    # Lock to top-right margin
    x_name = CANVAS_SIZE - MARGIN - bbox_title[2]
    y_name = MARGIN - bbox_title[1]

    RARE_GRADIENT       = ["#60B4FF", "#1A6FD4", "#0A3FA8", "#1A6FD4", "#60B4FF"]
    EPIC_GRADIENT       = ["#FFD700", "#C0C0C0", "#FFD700", "#C0C0C0", "#FFD700"]
    LEGENDARY_GRADIENT  = ["#ff66c4", "#5170ff", "#4b9941", "#ffb717", "#4b9941", "#5170ff", "#ff66c4"]

    SPECIAL_WORDS = {
        "RARE":      RARE_GRADIENT,
        "EPIC":      EPIC_GRADIENT,
        "LEGENDARY": LEGENDARY_GRADIENT,
    }

    space_w = ui_draw.textbbox((0, 0), " ", font=font_title)[2]
    cursor_x = x_name

    # Draw each word of the Character portion (plain white)
    for word in char_display.split(" "):
        word_w = ui_draw.textbbox((0, 0), word, font=font_title)[2]
        draw_text_with_outline(ui_draw, (cursor_x, y_name), word, font_title, (255, 255, 255))
        cursor_x += word_w + space_w

    # Draw the Adj in parentheses, gradient-coloured if it's a special word
    if adj_display:
        adj_token = f"({adj_display})"
        if adj_display in SPECIAL_WORDS:
            draw_gradient_text(ui_layer, (cursor_x, y_name), adj_token, font_title, SPECIAL_WORDS[adj_display])
        else:
            draw_text_with_outline(ui_draw, (cursor_x, y_name), adj_token, font_title, (255, 255, 255))

    lowest_title_y = y_name + bbox_title[3]

    # B. SKIP TIME TEXT
    time_text = f"Skip Time: {round(float(args.time))} min"
    font_time = load_font(int(CANVAS_SIZE * FONT_TIME_RATIO))
    bbox_time = ui_draw.textbbox((0, 0), time_text, font_time)

    # Strip invisible padding to lock strictly to Bottom-Left margin
    x_time = MARGIN - bbox_time[0]
    y_time = CANVAS_SIZE - int(CANVAS_SIZE * 0.08) - bbox_time[3]

    draw_text_with_outline(ui_draw, (x_time, y_time), time_text, font_time, (255, 255, 255))
    highest_time_y = y_time + bbox_time[1]

    # C. FILAMENT SECTION
    font_num = load_font(int(CANVAS_SIZE * FONT_NUM_RATIO))
    font_text = load_font(int(CANVAS_SIZE * FONT_TEXT_RATIO))
    font_mass = load_font(int(CANVAS_SIZE * FONT_MASS_RATIO))

    active_colors = [c.split('|') for c in args.colors if len(c.split('|')) == 3 and float(c.split('|')[2]) > 0]

    if active_colors:
        start_y = lowest_title_y + int(CANVAS_SIZE * 0.03)
        available_height = highest_time_y - start_y - int(CANVAS_SIZE * 0.02)
        row_spacing = min(COLOR_BOX_SIZE + int(CANVAS_SIZE * 0.02), available_height // len(active_colors))

        for idx, (cname, chex, cmass) in enumerate(active_colors):
            y = start_y + (idx * row_spacing)
            box_x = CANVAS_SIZE - MARGIN - COLOR_BOX_SIZE
            rgb = parse_hex_to_rgb(chex)
            rounded_mass = int(math.ceil(float(cmass) / 10.0)) * 10

            # Check if this RGB tuple is mapped to a gradient in the CSV
            is_silk = "silk" in cname.lower()

            if rgb in gradient_library:
                grad_img = create_gradient_swatch(COLOR_BOX_SIZE, COLOR_BOX_SIZE, gradient_library[rgb])
                background.paste(grad_img, (box_x, y))
                if is_silk:
                    draw_metallic_border(ui_draw, box_x, y, COLOR_BOX_SIZE, rgb)
                else:
                    ui_draw.rectangle([box_x, y, box_x + COLOR_BOX_SIZE, y + COLOR_BOX_SIZE], outline="gray", width=2)
            elif is_silk:
                silk_img = create_silk_swatch(COLOR_BOX_SIZE, COLOR_BOX_SIZE, rgb)
                background.paste(silk_img, (box_x, y))
                draw_metallic_border(ui_draw, box_x, y, COLOR_BOX_SIZE, rgb)
            else:
                ui_draw.rectangle([box_x, y, box_x + COLOR_BOX_SIZE, y + COLOR_BOX_SIZE], fill=rgb)

            num_txt = str(idx + 1)
            num_w = ui_draw.textbbox((0, 0), num_txt, font=font_num)[2]
            num_h = ui_draw.textbbox((0, 0), num_txt, font=font_num)[3]
            num_color = (0, 0, 0) if is_color_light(rgb) else (255, 255, 255)
            ui_draw.text(
                (box_x + (COLOR_BOX_SIZE - num_w) // 2, y + (COLOR_BOX_SIZE - num_h) // 2 - int(CANVAS_SIZE * 0.008)),
                num_txt, font=font_num, fill=num_color)

            mass_txt = f"{rounded_mass} g"

            # Split color name into brand line and rest line.
            # Checked longest-match first so "Sunlu Silk" beats "Sunlu".
            BRAND_PREFIXES = ["Sunlu Silk", "Voxel", "Esun", "Sunlu", "Eryone"]
            brand_line = cname
            rest_line  = ""
            for prefix in BRAND_PREFIXES:
                if cname.lower().startswith(prefix.lower()):
                    brand_line = cname[:len(prefix)]
                    rest_line  = cname[len(prefix):].strip()
                    break

            # Distribute lines evenly across the full swatch height
            num_lines = 3 if rest_line else 2
            slot_h    = COLOR_BOX_SIZE // num_lines

            # Shrink font until brand line fits its slot with a 2px cushion
            fit_font = font_text
            fit_size = int(CANVAS_SIZE * FONT_TEXT_RATIO)
            while ui_draw.textbbox((0, 0), brand_line, font=fit_font)[3] > (slot_h - 2) and fit_size > 8:
                fit_size -= 1
                fit_font = load_font(fit_size)

            # Recalculate all widths and x positions with the final font
            right_edge = box_x - int(CANVAS_SIZE * 0.02)
            brand_w = ui_draw.textbbox((0, 0), brand_line, font=fit_font)[2]
            rest_w  = ui_draw.textbbox((0, 0), rest_line,  font=fit_font)[2] if rest_line else 0
            mass_w  = ui_draw.textbbox((0, 0), mass_txt,   font=fit_font)[2]
            line_h  = ui_draw.textbbox((0, 0), brand_line, font=fit_font)[3]
            mass_h  = ui_draw.textbbox((0, 0), mass_txt,   font=fit_font)[3]

            brand_x = right_edge - brand_w
            rest_x  = right_edge - rest_w
            mass_x  = right_edge - mass_w

            brand_y = y + (slot_h - line_h) // 2
            draw_text_with_outline(ui_draw, (brand_x, brand_y), brand_line, font=fit_font, fill=(255, 255, 255))

            if rest_line:
                rest_y = y + slot_h + (slot_h - line_h) // 2
                draw_text_with_outline(ui_draw, (rest_x, rest_y), rest_line, font=fit_font, fill=(255, 255, 255))

            mass_y = y + slot_h * (num_lines - 1) + (slot_h - mass_h) // 2
            draw_text_with_outline(ui_draw, (mass_x, mass_y), mass_txt, font=fit_font, fill=(180, 180, 180))

    # --- 4. DYNAMIC SPATIAL COLLISION SCALING ---
    if os.path.exists(args.img):
        char_img = Image.open(args.img).convert("RGBA")
        bbox = char_img.getbbox()
        if bbox:
            char_img = char_img.crop(bbox)

        # 1.5% canvas buffer so the image dodges text beautifully
        buffer_pixels = int(CANVAS_SIZE * 0.015)
        filter_size = buffer_pixels * 2 + 1

        ui_mask = ui_layer.split()[3]
        ui_bumper = ui_mask.filter(ImageFilter.MaxFilter(filter_size)).convert("1")
        char_mask = char_img.split()[3]

        def check_scale(test_scale):
            test_w = int(char_img.width * test_scale)
            test_h = int(char_img.height * test_scale)

            # Instantly reject if the image exceeds the canvas itself
            if test_w > (CANVAS_SIZE - 2 * MARGIN) or test_h > (CANVAS_SIZE - 2 * MARGIN):
                return False, None

            y_max = CANVAS_SIZE - MARGIN - test_h
            x_max = CANVAS_SIZE - MARGIN - test_w

            if MARGIN > y_max or MARGIN > x_max:
                return False, None

            test_char_mask = char_mask.resize((test_w, test_h), Image.Resampling.NEAREST).convert("1")

            # Scan step size (larger = faster, smaller = tighter fit)
            step = max(5, int(CANVAS_SIZE * 0.01))

            y_range = list(range(MARGIN, y_max + 1, step))
            # Sort Y to prefer vertically centered placements when possible
            y_center_ideal = (CANVAS_SIZE - test_h) // 2
            y_range.sort(key=lambda y: abs(y - y_center_ideal))

            # Scan X from left to right
            x_range = list(range(MARGIN, x_max + 1, step))

            # SLIDING WINDOW: Try pasting the image all over the safe area
            for paste_x in x_range:
                for paste_y in y_range:
                    # Crop the UI forcefield to only the size of the test image
                    ui_crop = ui_bumper.crop((paste_x, paste_y, paste_x + test_w, paste_y + test_h))

                    # If any pixels overlap, collision is true. If none overlap, we found a safe spot!
                    collision = ImageChops.logical_and(ui_crop, test_char_mask)
                    if not collision.getbbox():
                        return True, (paste_x, paste_y)

            return False, None

        # Binary Search Variables
        low_scale = 0.1
        high_scale = min(CANVAS_SIZE / char_img.width, CANVAS_SIZE / char_img.height) * 2.0
        best_scale = low_scale
        best_pos = (MARGIN, MARGIN)

        # 15 iterations ensures mathematically perfect scaling
        for _ in range(15):
            test_scale = (low_scale + high_scale) / 2
            fits, pos = check_scale(test_scale)

            if fits:
                best_scale = test_scale
                best_pos = pos
                low_scale = test_scale  # It fits, let's try going bigger
            else:
                high_scale = test_scale  # It crashed, shrink it down

        # Apply the absolute maximum scale discovered
        final_w = int(char_img.width * best_scale)
        final_h = int(char_img.height * best_scale)
        char_img = char_img.resize((final_w, final_h), Image.Resampling.LANCZOS)

        background.paste(char_img, best_pos, char_img)

    background.alpha_composite(ui_layer)
    # Save using our new specific naming convention
    background.save(output_path)
    print(f"Generated: {output_path}")


if __name__ == "__main__":
    main()