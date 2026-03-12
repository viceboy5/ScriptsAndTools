import argparse
import os
from PIL import Image, ImageDraw, ImageFont

# Canvas Configuration
CANVAS_SIZE = 512
MARGIN = 24
COLOR_BOX_SIZE = 60
IMAGE_LEFT_RATIO = 0.55


def load_font(size):
    # Tries to find a fun, chunky font natively installed on Windows, falls back to Arial
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


def draw_text_with_outline(draw, pos, text, font, fill, outline_color=(0, 0, 0), outline_width=2):
    x, y = pos
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
    parser.add_argument("--colors", nargs='*', default=[])  # Format expected: "Name|Hex|Mass"
    args = parser.parse_args()

    # Create Canvas (Dark Grey matching your Canva example)
    canvas = Image.new("RGBA", (CANVAS_SIZE, CANVAS_SIZE), (75, 75, 80, 255))
    draw = ImageDraw.Draw(canvas)

    # 1. Draw Character Image
    if os.path.exists(args.img):
        char_img = Image.open(args.img).convert("RGBA")
        max_width = int(CANVAS_SIZE * IMAGE_LEFT_RATIO)
        max_height = CANVAS_SIZE - MARGIN * 2

        char_img.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)
        x = MARGIN
        y = (CANVAS_SIZE - char_img.height) // 2
        canvas.paste(char_img, (x, y), char_img)

    # 2. Draw Name (Top Right)
    font_title = load_font(46)
    bbox = draw.textbbox((0, 0), args.name.upper(), font=font_title)
    x_name = CANVAS_SIZE - MARGIN - (bbox[2] - bbox[0])
    draw_text_with_outline(draw, (x_name, MARGIN), args.name.upper(), font_title, (255, 255, 255))

    # 3. Draw Colors
    font_num = load_font(32)
    font_text = load_font(22)
    font_mass = load_font(20)

    # Filter out empty masses
    active_colors = [c.split('|') for c in args.colors if len(c.split('|')) == 3 and float(c.split('|')[2]) > 0]

    if active_colors:
        start_y = MARGIN + 60
        available_height = CANVAS_SIZE - start_y - 80
        row_spacing = min(COLOR_BOX_SIZE + 20, available_height // len(active_colors))

        for idx, (cname, chex, cmass) in enumerate(active_colors):
            y = start_y + (idx * row_spacing)
            box_x = CANVAS_SIZE - MARGIN - COLOR_BOX_SIZE
            rgb = parse_hex_to_rgb(chex)

            # Color Box
            draw.rectangle([box_x, y, box_x + COLOR_BOX_SIZE, y + COLOR_BOX_SIZE], fill=rgb)

            # Number inside box
            num_txt = str(idx + 1)
            num_w = draw.textbbox((0, 0), num_txt, font=font_num)[2]
            num_h = draw.textbbox((0, 0), num_txt, font=font_num)[3]
            num_color = (0, 0, 0) if is_color_light(rgb) else (255, 255, 255)
            draw.text((box_x + (COLOR_BOX_SIZE - num_w) // 2, y + (COLOR_BOX_SIZE - num_h) // 2 - 4), num_txt,
                      font=font_num, fill=num_color)

            # Name and Mass (Right-aligned to the box)
            mass_txt = f"{float(cmass):.0f} g"
            name_w = draw.textbbox((0, 0), cname, font=font_text)[2]
            mass_w = draw.textbbox((0, 0), mass_txt, font=font_mass)[2]

            draw_text_with_outline(draw, (box_x - 10 - name_w, y + 4), cname, font=font_text, fill=(255, 255, 255))
            draw_text_with_outline(draw, (box_x - 10 - mass_w, y + 30), mass_txt, font=font_mass, fill=(200, 200, 200))

    # 4. Draw Skip Time (Bottom Left)
    time_text = f"Skip Time: {args.time} min"
    draw_text_with_outline(draw, (MARGIN, CANVAS_SIZE - MARGIN - 40), time_text, load_font(38), (255, 255, 255))

    # Save Composite
    canvas.save(args.out)


if __name__ == "__main__":
    main()