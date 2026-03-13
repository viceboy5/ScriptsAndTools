import argparse
import os
import re
import math
from PIL import Image, ImageDraw, ImageFont, ImageChops, ImageFilter

# --- DYNAMIC CONFIGURATION ---
CANVAS_SIZE = 1080

# Layout Ratios
MARGIN_RATIO = 0.01
COLOR_BOX_SIZE_RATIO = 0.12
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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--name", required=True)
    parser.add_argument("--time", required=True)
    parser.add_argument("--img", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--colors", nargs='*', default=[])
    args = parser.parse_args()

    # --- 1. CLEAN NAME LOGIC ---
    original_name = args.name
    clean_name = re.sub(r'(?i)[._-]Full$', '', original_name).replace('.', ' ').replace('_', ' ').strip().upper()

    # Create the specific output filename for the Batch script to find
    # Example: "MyModel_Full.gcode.3mf" -> "MyModel_slicePreview.png"
    base_filename = re.sub(r'(?i)[._-]Full(\.gcode(\.3mf)?)?$', '', original_name)
    output_filename = f"{base_filename}_slicePreview.png"
    output_path = os.path.join(os.path.dirname(args.out), output_filename)

    # --- 2. INITIALIZE CANVAS ---
    background = Image.new("RGBA", (CANVAS_SIZE, CANVAS_SIZE), (0, 0, 0, 255))
    ui_layer   = Image.new("RGBA", (CANVAS_SIZE, CANVAS_SIZE), (0, 0, 0, 0))
    ui_draw    = ImageDraw.Draw(ui_layer)

    # --- 3. DRAW ALL UI ELEMENTS TO THE TRANSPARENT LAYER ---

    # A. TITLE TEXT
    current_font_size = int(CANVAS_SIZE * FONT_TITLE_RATIO)
    font_title = load_font(current_font_size)
    max_title_width = CANVAS_SIZE - (MARGIN * 2)

    while ui_draw.textbbox((0, 0), clean_name, font=font_title)[2] > max_title_width and current_font_size > int(
            CANVAS_SIZE * 0.03):
        current_font_size -= 2
        font_title = load_font(current_font_size)

    bbox_title = ui_draw.textbbox((0, 0), clean_name, font=font_title)

    # Strip invisible padding to lock strictly to Top-Right margin
    x_name = CANVAS_SIZE - MARGIN - bbox_title[2]
    y_name = MARGIN - bbox_title[1]

    draw_text_with_outline(ui_draw, (x_name, y_name), clean_name, font_title, (255, 255, 255))
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
        start_y = lowest_title_y + int(CANVAS_SIZE * 0.05)
        available_height = highest_time_y - start_y - int(CANVAS_SIZE * 0.05)
        row_spacing = min(COLOR_BOX_SIZE + int(CANVAS_SIZE * 0.04), available_height // len(active_colors))

        for idx, (cname, chex, cmass) in enumerate(active_colors):
            y = start_y + (idx * row_spacing)
            box_x = CANVAS_SIZE - MARGIN - COLOR_BOX_SIZE
            rgb = parse_hex_to_rgb(chex)
            rounded_mass = int(math.ceil(float(cmass) / 10.0)) * 10

            ui_draw.rectangle([box_x, y, box_x + COLOR_BOX_SIZE, y + COLOR_BOX_SIZE], fill=rgb)

            num_txt = str(idx + 1)
            num_w = ui_draw.textbbox((0, 0), num_txt, font=font_num)[2]
            num_h = ui_draw.textbbox((0, 0), num_txt, font=font_num)[3]
            num_color = (0, 0, 0) if is_color_light(rgb) else (255, 255, 255)
            ui_draw.text(
                (box_x + (COLOR_BOX_SIZE - num_w) // 2, y + (COLOR_BOX_SIZE - num_h) // 2 - int(CANVAS_SIZE * 0.008)),
                num_txt, font=font_num, fill=num_color)

            mass_txt = f"{rounded_mass} g"
            name_w = ui_draw.textbbox((0, 0), cname, font=font_text)[2]
            mass_w = ui_draw.textbbox((0, 0), mass_txt, font=font_mass)[2]

            text_x_pos = box_x - int(CANVAS_SIZE * 0.02) - name_w
            mass_x_pos = box_x - int(CANVAS_SIZE * 0.02) - mass_w

            draw_text_with_outline(ui_draw, (text_x_pos, y + int(CANVAS_SIZE * 0.008)), cname, font=font_text,
                                   fill=(255, 255, 255))
            draw_text_with_outline(ui_draw, (mass_x_pos, y + int(CANVAS_SIZE * 0.058)), mass_txt, font=font_mass,
                                   fill=(200, 200, 200))

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