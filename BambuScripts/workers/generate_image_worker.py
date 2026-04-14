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
    gradient_map = {}
    script_dir = os.path.dirname(os.path.abspath(__file__))
    csv_path = os.path.join(script_dir, csv_filename)

    if os.path.exists(csv_path):
        with open(csv_path, mode='r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            try:
                next(reader)
            except StopIteration:
                return gradient_map

            for row in reader:
                if len(row) > 4:
                    try:
                        # Safely unpack variables to avoid index brackets
                        c0, c1, c2, c3 = row[:4]
                        base_rgb = (int(c1.strip()), int(c2.strip()), int(c3.strip()))

                        gradients = []
                        for cell in row[4:]:
                            if cell.strip().startswith('#'):
                                gradients.append(cell.strip())

                        if len(gradients) >= 2:
                            gradient_map[base_rgb] = gradients
                    except ValueError:
                        continue
    return gradient_map


def create_gradient_swatch(width, height, hex_colors):
    base = Image.new('RGB', (width, height))
    draw = ImageDraw.Draw(base)
    rgb_colors = []
    for h in hex_colors:
        h = h.strip().lstrip('#')
        rgb_colors.append(tuple(int(h[i:i + 2], 16) for i in (0, 2, 4)))

    segment_width = width / (len(rgb_colors) - 1)

    for i in range(len(rgb_colors) - 1):
        c1_r, c1_g, c1_b = rgb_colors[i]
        c2_r, c2_g, c2_b = rgb_colors[i + 1]
        start_x = int(i * segment_width)
        end_x = int((i + 1) * segment_width) if i < len(rgb_colors) - 2 else width

        for x in range(start_x, end_x):
            ratio = (x - start_x) / max(1, (end_x - start_x))
            r = int(c1_r * (1 - ratio) + c2_r * ratio)
            g = int(c1_g * (1 - ratio) + c2_g * ratio)
            b = int(c1_b * (1 - ratio) + c2_b * ratio)
            draw.line([(x, 0), (x, height)], fill=(r, g, b))
    return base


def draw_gradient_text(layer, pos, text, font, gradient_colors, outline_color=(0, 0, 0), outline_ratio=0.004):
    x, y = pos
    outline_width = max(1, int(CANVAS_SIZE * outline_ratio))
    draw = ImageDraw.Draw(layer)
    for dx in range(-outline_width, outline_width + 1):
        for dy in range(-outline_width, outline_width + 1):
            if dx != 0 or dy != 0:
                draw.text((x + dx, y + dy), text, font=font, fill=outline_color)

    bbox = draw.textbbox((0, 0), text, font=font)
    b_left, b_top, b_right, b_bottom = bbox
    tw, th = b_right, b_bottom

    if tw <= 0 or th <= 0:
        return

    mask_img = Image.new("L", (tw, th), 0)
    ImageDraw.Draw(mask_img).text((-b_left, -b_top), text, font=font, fill=255)
    grad_rgba = create_gradient_swatch(tw, th, gradient_colors).convert("RGBA")
    grad_rgba.putalpha(mask_img)
    layer.paste(grad_rgba, (x, y + b_top), grad_rgba)


def draw_metallic_border(draw, box_x, y, size, base_rgb, border_width=4):
    r, g, b = base_rgb
    highlight = (min(255, int(r * 0.5 + 200)), min(255, int(g * 0.5 + 200)), min(255, int(b * 0.5 + 200)))
    midtone = (min(255, int(r * 0.7 + 80)), min(255, int(g * 0.7 + 80)), min(255, int(b * 0.7 + 80)))
    shadow = (max(0, int(r * 0.3)), max(0, int(g * 0.3)), max(0, int(b * 0.3)))
    draw.rectangle([box_x, y, box_x + size, y + size], outline=midtone, width=border_width)
    o = border_width // 2
    draw.line([box_x + o, y + o, box_x + size - o, y + o], fill=highlight, width=max(1, border_width // 2))
    draw.line([box_x + o, y + o, box_x + o, y + size - o], fill=highlight, width=max(1, border_width // 2))
    draw.line([box_x + o, y + size - o, box_x + size - o, y + size - o], fill=shadow, width=max(1, border_width // 2))
    draw.line([box_x + size - o, y + o, box_x + size - o, y + size - o], fill=shadow, width=max(1, border_width // 2))


def create_silk_swatch(width, height, base_rgb):
    r, g, b = base_rgb
    hr, hg, hb = (min(255, int(r * 0.55 + 200)), min(255, int(g * 0.55 + 200)), min(255, int(b * 0.55 + 200)))
    sr, sg, sb = (max(0, int(r * 0.35)), max(0, int(g * 0.35)), max(0, int(b * 0.35)))

    img = Image.new("RGB", (width, height))
    pixels = img.load()
    for px in range(width):
        for py in range(height):
            diag = (px / width + py / height) / 2.0
            t = (diag - 0.5) * 2.0
            t = t * abs(t)
            if t < 0:
                s = -t
                pr = int(r * (1 - s) + hr * s)
                pg = int(g * (1 - s) + hg * s)
                pb = int(b * (1 - s) + hb * s)
            else:
                s = t
                pr = int(r * (1 - s) + sr * s)
                pg = int(g * (1 - s) + sg * s)
                pb = int(b * (1 - s) + sb * s)
            pixels[px, py] = (pr, pg, pb)
    return img


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--name", required=True)
    parser.add_argument("--time", required=True)
    parser.add_argument("--img", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--tag", default="")
    parser.add_argument("--colors", nargs='*', default=[])
    args = parser.parse_args()

    gradient_library = load_gradient_library()

    original_name = args.name
    name_stripped = re.sub(r'(?i)[ ._-]Full$', '', original_name)
    base_filename = name_stripped
    output_filename = f"{base_filename}_slicePreview.png"
    output_path = os.path.join(os.path.dirname(args.out), output_filename)

    display_name = name_stripped
    known_prefixes = ["P2S", "X1C", "H2S"]
    for prefix in known_prefixes:
        if display_name.upper().startswith(prefix + "_") or display_name.upper().startswith(prefix + "-"):
            display_name = display_name[len(prefix) + 1:]
            break

    parts = display_name.split('_')
    if len(parts) >= 3:
        character_raw = parts.pop(0)
        theme_raw = parts.pop(-1)
        adj_raw = '_'.join(parts)
    elif len(parts) == 2:
        character_raw = parts[0]
        adj_raw = ''
        theme_raw = parts[1]
    else:
        character_raw = parts[0]
        adj_raw = ''
        theme_raw = ''

    char_display = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', character_raw).upper()
    adj_display = adj_raw.upper()

    # Prepend tag to title when provided: "KC - HUNTER" instead of "KCHUNTER"
    # Strip the tag prefix from the character name first to avoid doubling (e.g. "KC" + "KCHunter" -> "KC - HUNTER")
    tag_display = args.tag.strip().upper()
    if tag_display:
        if char_display.upper().startswith(tag_display):
            char_display = char_display[len(tag_display):].lstrip(' -_')
        char_display = f"{tag_display} - {char_display}"

    background = Image.new("RGBA", (CANVAS_SIZE, CANVAS_SIZE), (0, 0, 0, 255))
    ui_layer = Image.new("RGBA", (CANVAS_SIZE, CANVAS_SIZE), (0, 0, 0, 0))
    ui_draw = ImageDraw.Draw(ui_layer)

    # --- PARSE COLORS ---
    active_colors = []
    for c in args.colors:
        pcs = c.split('|')
        if len(pcs) == 3 and float(pcs[2]) > 0:
            active_colors.append(pcs)

    # --- PRELIMINARY SWATCH GEOMETRY ---
    # Use a conservative title-height estimate so effective_box is stable across both passes.
    # FONT_TITLE_RATIO * 1.3 + padding gives ~13% of canvas (~67px), safely above actual title height.
    font_num  = load_font(int(CANVAS_SIZE * FONT_NUM_RATIO))
    font_text = load_font(int(CANVAS_SIZE * FONT_TEXT_RATIO))

    n = len(active_colors)
    swatch_top_prelim = int(CANVAS_SIZE * 0.13)   # conservative top offset for effective_box sizing
    if n > 0:
        row_h_prelim  = (CANVAS_SIZE - MARGIN - swatch_top_prelim) / n
        effective_box = min(int(CANVAS_SIZE * 0.40),
                            max(int(CANVAS_SIZE * 0.07), int(row_h_prelim) - 4))
    else:
        effective_box = COLOR_BOX_SIZE
    box_x      = CANVAS_SIZE - MARGIN - effective_box
    right_edge = box_x - int(CANVAS_SIZE * 0.02)

    # --- FIRST PASS: measure filament labels → establish left-column boundary ---
    BRAND_PREFIXES = ["Sunlu Silk", "Voxel", "Esun", "Sunlu", "Eryone"]
    slot_renders       = []
    max_fil_text_width = 0

    for cname, chex, cmass in active_colors:
        rounded_mass = int(math.ceil(float(cmass) / 10.0)) * 10
        mass_txt   = f"{rounded_mass} g"
        brand_line = cname
        rest_line  = ""
        for prefix in BRAND_PREFIXES:
            if cname.lower().startswith(prefix.lower()):
                brand_line = cname[:len(prefix)]
                rest_line  = cname[len(prefix):].strip()
                break
        num_lines = 3 if rest_line else 2
        slot_h    = effective_box // num_lines

        fit_size = int(CANVAS_SIZE * FONT_TEXT_RATIO)
        fit_font = font_text
        while True:
            _, _, _, test_h = ui_draw.textbbox((0, 0), brand_line, font=fit_font)
            if test_h > (slot_h - 2) and fit_size > 8:
                fit_size -= 1
                fit_font  = load_font(fit_size)
            else:
                break

        _, _, brand_w, line_h = ui_draw.textbbox((0, 0), brand_line, font=fit_font)
        rest_w = 0
        if rest_line:
            _, _, rest_w, _ = ui_draw.textbbox((0, 0), rest_line, font=fit_font)
        _, _, mass_w, mass_h = ui_draw.textbbox((0, 0), mass_txt, font=fit_font)

        max_fil_text_width = max(max_fil_text_width, brand_w, rest_w, mass_w)
        slot_renders.append((cname, chex, cmass, brand_line, rest_line, mass_txt,
                             fit_font, slot_h, line_h, brand_w, rest_w, mass_w, mass_h))

    fil_gap        = int(CANVAS_SIZE * 0.01)
    left_col_right = right_edge - max_fil_text_width - fil_gap
    max_left_width = max(MARGIN * 4, left_col_right - MARGIN)

    # --- TITLE: centered on the full canvas, shrinks if needed ---
    # The title sits above the swatch zone so it can use the full canvas width.
    max_title_width    = CANVAS_SIZE - 2 * MARGIN
    display_for_sizing = char_display + (f" ({adj_display})" if adj_display else "")
    current_font_size  = int(CANVAS_SIZE * FONT_TITLE_RATIO)
    font_title = load_font(current_font_size)

    while True:
        _, _, w, _ = ui_draw.textbbox((0, 0), display_for_sizing, font=font_title)
        if w > max_title_width and current_font_size > int(CANVAS_SIZE * 0.03):
            current_font_size -= 2
            font_title = load_font(current_font_size)
        else:
            break

    bbox_title = ui_draw.textbbox((0, 0), display_for_sizing, font=font_title)
    title_l, title_top, title_r, title_bottom = bbox_title
    title_w = title_r - title_l
    x_name  = (CANVAS_SIZE - title_w) // 2   # centered on full canvas width
    y_name  = MARGIN - title_top

    RARE_GRADIENT      = ["#60B4FF", "#1A6FD4", "#0A3FA8", "#1A6FD4", "#60B4FF"]
    EPIC_GRADIENT      = ["#FFD700", "#C0C0C0", "#FFD700", "#C0C0C0", "#FFD700"]
    LEGENDARY_GRADIENT = ["#ff66c4", "#5170ff", "#4b9941", "#ffb717", "#4b9941", "#5170ff", "#ff66c4"]
    SPECIAL_WORDS      = {"RARE": RARE_GRADIENT, "EPIC": EPIC_GRADIENT, "LEGENDARY": LEGENDARY_GRADIENT}

    _, _, space_w, _ = ui_draw.textbbox((0, 0), " ", font=font_title)
    cursor_x = x_name
    for word in char_display.split(" "):
        _, _, word_w, _ = ui_draw.textbbox((0, 0), word, font=font_title)
        draw_text_with_outline(ui_draw, (cursor_x, y_name), word, font_title, (255, 255, 255))
        cursor_x += word_w + space_w
    if adj_display:
        adj_token = f"({adj_display})"
        if adj_display in SPECIAL_WORDS:
            draw_gradient_text(ui_layer, (cursor_x, y_name), adj_token, font_title, SPECIAL_WORDS[adj_display])
        else:
            draw_text_with_outline(ui_draw, (cursor_x, y_name), adj_token, font_title, (255, 255, 255))

    # --- FINAL SWATCH GEOMETRY: first row starts just below the title ---
    title_bottom_y = y_name + title_bottom
    swatch_top_y   = title_bottom_y + int(CANVAS_SIZE * 0.015)
    if n > 0:
        row_h = (CANVAS_SIZE - MARGIN - swatch_top_y) / n
        # Clamp effective_box to actual row height (in case title was taller than estimate)
        effective_box = min(effective_box, max(int(CANVAS_SIZE * 0.07), int(row_h) - 4))
        box_x      = CANVAS_SIZE - MARGIN - effective_box
        right_edge = box_x - int(CANVAS_SIZE * 0.02)
    else:
        row_h = 0

    # --- SKIP TIME: bottom-left, unchanged position, red border box for visual separation ---
    time_text      = f"Skip Time: {round(float(args.time))} min"
    time_font_size = int(CANVAS_SIZE * FONT_TIME_RATIO)
    font_time      = load_font(time_font_size)

    while True:
        _, _, tw, _ = ui_draw.textbbox((0, 0), time_text, font=font_time)
        if tw > max_left_width and time_font_size > int(CANVAS_SIZE * 0.03):
            time_font_size -= 2
            font_time = load_font(time_font_size)
        else:
            break

    bbox_time = ui_draw.textbbox((0, 0), time_text, font_time)
    time_l, time_top, time_r, time_bottom = bbox_time
    x_time = MARGIN - time_l
    y_time = CANVAS_SIZE - int(CANVAS_SIZE * 0.08) - time_bottom
    draw_text_with_outline(ui_draw, (x_time, y_time), time_text, font_time, (255, 255, 255))

    # Red border box around Skip Time
    box_pad = max(4, int(CANVAS_SIZE * 0.008))
    ui_draw.rectangle(
        [MARGIN             - box_pad,
         y_time + time_top  - box_pad,
         x_time + time_r    + box_pad,
         y_time + time_bottom + box_pad],
        outline=(210, 40, 40), width=2)

    # --- SECOND PASS: draw swatches ---
    for idx, (cname, chex, cmass, brand_line, rest_line, mass_txt,
              fit_font, slot_h, line_h, brand_w, rest_w, mass_w, mass_h) in enumerate(slot_renders):
        y   = int(swatch_top_y + idx * row_h + (row_h - effective_box) / 2)
        rgb = parse_hex_to_rgb(chex)
        is_silk = "silk" in cname.lower()

        if rgb in gradient_library:
            grad_img = create_gradient_swatch(effective_box, effective_box, gradient_library[rgb])
            background.paste(grad_img, (box_x, y))
            if is_silk:
                draw_metallic_border(ui_draw, box_x, y, effective_box, rgb)
            else:
                ui_draw.rectangle([box_x, y, box_x + effective_box, y + effective_box], outline="gray", width=2)
        elif is_silk:
            silk_img = create_silk_swatch(effective_box, effective_box, rgb)
            background.paste(silk_img, (box_x, y))
            draw_metallic_border(ui_draw, box_x, y, effective_box, rgb)
        else:
            ui_draw.rectangle([box_x, y, box_x + effective_box, y + effective_box], fill=rgb)

        num_txt = str(idx + 1)
        _, _, num_w, num_h = ui_draw.textbbox((0, 0), num_txt, font=font_num)
        num_color = (0, 0, 0) if is_color_light(rgb) else (255, 255, 255)
        ui_draw.text(
            (box_x + (effective_box - num_w) // 2, y + (effective_box - num_h) // 2 - int(CANVAS_SIZE * 0.008)),
            num_txt, font=font_num, fill=num_color)

        brand_y = y + (slot_h - line_h) // 2
        draw_text_with_outline(ui_draw, (right_edge - brand_w, brand_y), brand_line, font=fit_font, fill=(255, 255, 255))
        if rest_line:
            rest_y = y + slot_h + (slot_h - line_h) // 2
            draw_text_with_outline(ui_draw, (right_edge - rest_w, rest_y), rest_line, font=fit_font, fill=(255, 255, 255))
        mass_y = y + slot_h * (2 if rest_line else 1) + (slot_h - mass_h) // 2
        draw_text_with_outline(ui_draw, (right_edge - mass_w, mass_y), mass_txt, font=fit_font, fill=(180, 180, 180))

    if os.path.exists(args.img):
        char_img = Image.open(args.img).convert("RGBA")
        bbox = char_img.getbbox()
        if bbox:
            char_img = char_img.crop(bbox)

        buffer_pixels = int(CANVAS_SIZE * 0.015)
        filter_size = buffer_pixels * 2 + 1

        _, _, _, ui_mask = ui_layer.split()
        ui_bumper = ui_mask.filter(ImageFilter.MaxFilter(filter_size)).convert("1")
        _, _, _, char_mask = char_img.split()

        def check_scale(test_scale):
            test_w = int(char_img.width * test_scale)
            test_h = int(char_img.height * test_scale)

            if test_w > (CANVAS_SIZE - 2 * MARGIN) or test_h > (CANVAS_SIZE - 2 * MARGIN):
                return False, None

            y_max = CANVAS_SIZE - MARGIN - test_h
            x_max = CANVAS_SIZE - MARGIN - test_w

            if MARGIN > y_max or MARGIN > x_max:
                return False, None

            test_char_mask = char_mask.resize((test_w, test_h), Image.Resampling.NEAREST).convert("1")
            step = max(5, int(CANVAS_SIZE * 0.01))
            y_range = list(range(MARGIN, y_max + 1, step))
            y_center_ideal = (CANVAS_SIZE - test_h) // 2
            y_range.sort(key=lambda yi: abs(yi - y_center_ideal))
            x_range = list(range(MARGIN, x_max + 1, step))

            for paste_x in x_range:
                for paste_y in y_range:
                    ui_crop = ui_bumper.crop((paste_x, paste_y, paste_x + test_w, paste_y + test_h))
                    collision = ImageChops.logical_and(ui_crop, test_char_mask)
                    if not collision.getbbox():
                        return True, (paste_x, paste_y)
            return False, None

        low_scale = 0.1
        high_scale = min(CANVAS_SIZE / char_img.width, CANVAS_SIZE / char_img.height) * 2.0
        best_scale = low_scale
        best_pos = (MARGIN, MARGIN)

        for _ in range(15):
            test_scale = (low_scale + high_scale) / 2
            fits, pos = check_scale(test_scale)
            if fits:
                best_scale = test_scale
                best_pos = pos
                low_scale = test_scale
            else:
                high_scale = test_scale

        final_w = int(char_img.width * best_scale)
        final_h = int(char_img.height * best_scale)
        char_img = char_img.resize((final_w, final_h), Image.Resampling.LANCZOS)
        background.paste(char_img, best_pos, char_img)

    background.alpha_composite(ui_layer)
    background.save(output_path)
    print(f"Generated: {output_path}")


if __name__ == "__main__":
    main()