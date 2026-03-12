from PIL import Image, ImageDraw, ImageFont, ImageTk
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
import os
import re

# ===============================
# Configuration
# ===============================

CANVAS_SIZE = 512
MARGIN = 24
COLOR_BOX_SIZE = 60
IMAGE_LEFT_RATIO = 0.7  # proportion of canvas width for the character image

COLOR_MAP = {
    "Cold White": {"rgb": (240, 240, 240)},
    "Black": {"rgb": (0, 0, 0)},
    "Blue": {"rgb": (0, 102, 204)},
    "Pink": {"rgb": (255, 105, 180)},
    "Gold": {"rgb": (255, 215, 0)},
}

RARITY_COLORS = {
    "Common": (130, 130, 130),
    "Rare": (0, 102, 255),
    "Epic": (163, 53, 238),
    "Legendary": (255, 215, 0),
}

# ===============================
# Utility Functions
# ===============================

def load_font(size, bold=False):
    try:
        if bold:
            return ImageFont.truetype("arialbd.ttf", size)
        return ImageFont.truetype("arial.ttf", size)
    except:
        return ImageFont.load_default()

def is_color_light(rgb):
    r,g,b = rgb
    luminance = 0.299*r + 0.587*g + 0.114*b
    return luminance > 186

def sanitize_filename(s):
    return re.sub(r'[^a-zA-Z0-9_\- ]', '', s).strip()

def validate_inputs(character_name, color_names, rarity, image_path):
    if not character_name:
        raise ValueError("Character name cannot be empty")
    for cname in color_names:
        if cname not in COLOR_MAP:
            raise ValueError(f"Invalid color name: {cname}. Must be one of {list(COLOR_MAP.keys())}")
    if rarity not in RARITY_COLORS:
        raise ValueError(f"Invalid rarity: {rarity}. Must be one of {list(RARITY_COLORS.keys())}")
    if not os.path.isfile(image_path):
        raise FileNotFoundError("Character image not found")

# ===============================
# Drawing Functions
# ===============================

def create_canvas():
    return Image.new("RGBA", (CANVAS_SIZE, CANVAS_SIZE), (25,25,30,255))

def draw_character(canvas, path, left_ratio=IMAGE_LEFT_RATIO):
    img = Image.open(path).convert("RGBA")
    max_width = int(CANVAS_SIZE * left_ratio)
    max_height = CANVAS_SIZE - MARGIN*2
    img.thumbnail((max_width, max_height), Image.LANCZOS)
    x = 0  # stick to left edge
    y = (CANVAS_SIZE - img.height)//2
    canvas.paste(img, (x, y), img)

def draw_character_name(canvas, name):
    draw = ImageDraw.Draw(canvas)
    font = load_font(42, bold=True)
    bbox = draw.textbbox((0,0), name, font=font)
    x = (CANVAS_SIZE - (bbox[2]-bbox[0]))//2
    y = MARGIN//2
    draw.text((x,y), name, fill=(255,255,255), font=font)

def draw_rarity(canvas, rarity):
    draw = ImageDraw.Draw(canvas)
    font = load_font(26, bold=True)
    bbox = draw.textbbox((0,0), rarity, font=font)
    x = CANVAS_SIZE - MARGIN - (bbox[2]-bbox[0])
    y = CANVAS_SIZE - MARGIN - 30
    draw.text((x,y), rarity, fill=RARITY_COLORS[rarity], font=font)

def draw_text_with_outline(draw, pos, text, font, fill, outline_color=(0,0,0)):
    x, y = pos
    # Draw outline
    for dx in [-1,0,1]:
        for dy in [-1,0,1]:
            if dx !=0 or dy !=0:
                draw.text((x+dx, y+dy), text, font=font, fill=outline_color)
    # Draw main text
    draw.text((x,y), text, font=font, fill=fill)

def draw_color_panel(canvas, colors_data):
    draw = ImageDraw.Draw(canvas)
    font_num = load_font(32, bold=True)
    font_text = load_font(24, bold=True)
    font_mass = load_font(24, bold=True)

    total_rows = len(colors_data)
    start_y = 60
    row_spacing = (CANVAS_SIZE - 2*start_y) // total_rows

    for idx, (cname, mass) in enumerate(colors_data):
        y = start_y + idx*row_spacing

        # Box on far right
        box_x = CANVAS_SIZE - MARGIN - COLOR_BOX_SIZE
        box_y = y
        draw.rectangle([box_x, box_y, box_x+COLOR_BOX_SIZE, box_y+COLOR_BOX_SIZE], fill=COLOR_MAP[cname]["rgb"])

        # Number inside box
        number_text = str(idx+1)
        bbox_num = draw.textbbox((0,0), number_text, font=font_num)
        num_w = bbox_num[2]-bbox_num[0]
        num_h = bbox_num[3]-bbox_num[1]
        num_x = box_x + (COLOR_BOX_SIZE - num_w)//2
        num_y = box_y + (COLOR_BOX_SIZE - num_h)//2
        num_color = (0,0,0) if is_color_light(COLOR_MAP[cname]["rgb"]) else (255,255,255)
        draw_text_with_outline(draw, (num_x,num_y), number_text, font_num, num_color)

        # Color name and mass, right-aligned to box, vertically centered
        mass_text = f"{int(mass)} g"
        # Measure widths
        name_w = draw.textlength(cname, font=font_text)
        mass_w = draw.textlength(mass_text, font=font_mass)
        max_width = max(name_w, mass_w)
        text_x = box_x - 10 - max_width
        text_y = box_y + (COLOR_BOX_SIZE - font_text.size)//2
        draw_text_with_outline(draw, (text_x, text_y), cname, font_text, (230,230,230))
        draw_text_with_outline(draw, (text_x, text_y + 30), mass_text, font_mass, (180,180,180))

# ===============================
# GUI
# ===============================

def generate():
    try:
        character_name = name_entry.get().strip()
        rarity = rarity_combo.get().strip()
        image_path = image_path_var.get().strip()
        color_names = [color1_name.get().strip(),
                       color2_name.get().strip(),
                       color3_name.get().strip(),
                       color4_name.get().strip()]
        masses = [float(color1_mass.get()),
                  float(color2_mass.get()),
                  float(color3_mass.get()),
                  float(color4_mass.get())]

        validate_inputs(character_name, color_names, rarity, image_path)
        colors_data = list(zip(color_names, masses))

        canvas = create_canvas()
        draw_character(canvas, image_path)
        draw_character_name(canvas, character_name)
        draw_color_panel(canvas, colors_data)
        draw_rarity(canvas, rarity)

        safe_name = sanitize_filename(character_name) or "character"
        filename = f"{safe_name} Gcode Image.png"
        canvas.save(filename)
        messagebox.showinfo("Success", f"Image saved as {filename}")

        # Preview
        img = Image.open(filename)
        img.thumbnail((300,300), Image.LANCZOS)
        img_tk = ImageTk.PhotoImage(img)
        preview_label.config(image=img_tk)
        preview_label.image = img_tk
        preview_label.pack(pady=10)

    except Exception as e:
        messagebox.showerror("Error", str(e))

root = TkinterDnD.Tk()
root.title("Character Card Generator")
root.geometry("640x860")

image_path_var = tk.StringVar()

# Character name
tk.Label(root, text="Character Name", font=("Arial",12,"bold")).pack(pady=(10,0))
name_entry = tk.Entry(root, font=("Arial",14))
name_entry.pack(pady=(0,10), fill="x", padx=20)

# Rarity
tk.Label(root, text="Rarity", font=("Arial",12,"bold")).pack()
rarity_combo = ttk.Combobox(root, values=list(RARITY_COLORS.keys()), font=("Arial",12))
rarity_combo.pack(pady=(0,10), fill="x", padx=20)

# Image drag & drop / browse
tk.Label(root, text="Drag & Drop Character Image Below", font=("Arial",12,"bold")).pack()
image_label = tk.Label(root, text="Drop PNG Here", bg="#333", fg="white", height=4, font=("Arial",10))
image_label.pack(fill="x", padx=20, pady=5)

def drop(event):
    path = event.data.strip("{}")
    image_path_var.set(path)
    image_label.config(text=os.path.basename(path))

image_label.drop_target_register(DND_FILES)
image_label.dnd_bind("<<Drop>>", drop)

def browse():
    path = filedialog.askopenfilename(filetypes=[("PNG files","*.png")])
    if path:
        image_path_var.set(path)
        image_label.config(text=os.path.basename(path))

tk.Button(root, text="Browse Image", command=browse, font=("Arial",12)).pack(pady=5)

# Color inputs
def color_row(label_text):
    frame = tk.Frame(root)
    tk.Label(frame, text=label_text, font=("Arial",11)).pack(side="left", padx=5)
    color_combo = ttk.Combobox(frame, values=list(COLOR_MAP.keys()), font=("Arial",11), width=15)
    color_combo.pack(side="left", padx=5)
    mass_entry = tk.Entry(frame, width=10, font=("Arial",11))
    mass_entry.pack(side="left", padx=5)
    frame.pack(pady=8, fill="x", padx=20)
    return color_combo, mass_entry

color1_name, color1_mass = color_row("Color 1:")
color2_name, color2_mass = color_row("Color 2:")
color3_name, color3_mass = color_row("Color 3:")
color4_name, color4_mass = color_row("Color 4:")

# Preview label
preview_label = tk.Label(root)
preview_label.pack_forget()

tk.Button(root, text="Generate Card", command=generate, bg="#4CAF50", fg="white", font=("Arial",14)).pack(pady=20, fill="x", padx=20)

root.mainloop()
