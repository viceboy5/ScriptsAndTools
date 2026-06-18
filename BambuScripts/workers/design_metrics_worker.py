#!/usr/bin/env python3
"""design_metrics_worker.py

Per-design metric extractor for the efficiency analyses.

  PART A  - read straight from the design's *_Data.tsv (our data; trusted).
            print time, objects (pre-merge), color changes, model + total
            material, filament/unit, waste/unit, throughput, per-colour grams.

  PART B  - parsed from the sliced *Full.gcode.3mf (things NOT in our data):
            * plate utilization %                 (pick.png + plate_1.json + config)
            * print height, layer count           (gcode header)
            * variable layer height               (layer_heights_profile.txt)
            * feature mix (wall/infill/tower/...)  (gcode body)
            * travel distance + moves, retractions, z-hops   (gcode body)
            * outer-wall loops / fragmentation     (gcode body)
            * prime-tower + support filament       (gcode body)
            * per-layer print time                 (gcode body, M73 R)
            * per-object filament                  (gcode body, object labels)

Usage:
  python design_metrics_worker.py "C:\\ZB_Designs\\...\\X1C_Avocado_Foodz"   # folder
  python design_metrics_worker.py "...\\X1C_Avocado_Foodz_Full.gcode.3mf"
  python design_metrics_worker.py "..." --json
  python design_metrics_worker.py "..." --no-body     # skip the heavy gcode body pass
"""
import argparse
import glob
import io
import json
import math
import os
import re
import sys
import zipfile

try:
    from PIL import Image
except ImportError:
    sys.stderr.write("Pillow is required:  pip install Pillow\n")
    sys.exit(2)

# Front purge / flow-calibration line - fixed machine constant, keyed by bed size.
# Verified front purge / flow-calibration line, keyed by PRINTER (the TSV/folder
# prefix), so each machine gets its own even when beds match. (slice_info's
# printer_model_id is empty in current Bambu versions, so we key by name.)
CALIBRATION_LINES = {
    "X1C": (18.0, 0.0, 246.0, 14.0),   # verified
    # add "P2S" / "H2S" here once tuned; until then they use the estimate.
}
ALPHA_THRESHOLD = 10
DATE_RE = re.compile(r"^\d{1,2}/\d{1,2}/\d{4}$")


def _feature_bucket(name):
    n = name.lower()
    if "outer wall" in n:           return "outer_wall"
    if "inner wall" in n:           return "inner_wall"
    if "overhang" in n:             return "overhang"
    if "bridge" in n:               return "bridge"
    if "prime tower" in n or "wipe tower" in n: return "prime_tower"
    if "support" in n:              return "support"
    if "infill" in n or "surface" in n or "shell" in n: return "infill"
    return "other"


# =============================================================================
#  file discovery
# =============================================================================
def find_design_files(path):
    folder = path if os.path.isdir(path) else os.path.dirname(path)
    tsv = sorted(glob.glob(os.path.join(folder, "*_Data.tsv")))
    g3 = sorted(glob.glob(os.path.join(folder, "*Full.gcode.3mf")))
    if not g3:
        g3 = [f for f in glob.glob(os.path.join(folder, "*.gcode.3mf"))
              if not re.search(r"(?i)bod", os.path.basename(f))]
    return folder, (tsv[0] if tsv else None), (g3[0] if g3 else None)


def _is_design_folder(folder):
    try:
        return any(f.lower().endswith("full.gcode.3mf") for f in os.listdir(folder))
    except OSError:
        return False


def find_design_folders(paths):
    """For each dropped path, recursively collect every design folder beneath it
    (a folder that directly contains a *Full.gcode.3mf). A dropped .3mf or a
    single design folder resolves to just that folder. Folders WITHOUT a
    Full.gcode.3mf are skipped (per the harvest rule). De-duplicated, in order."""
    found, seen = [], set()
    def add(d):
        d = os.path.normpath(d)
        if d not in seen:
            seen.add(d); found.append(d)
    for p in paths:
        if os.path.isfile(p) and p.lower().endswith(".3mf"):
            add(os.path.dirname(p))
        elif os.path.isdir(p):
            for root, dirs, files in os.walk(p):
                if any(f.lower().endswith("full.gcode.3mf") for f in files):
                    add(root)
    return found


# =============================================================================
#  PART A - the _Data.tsv (our data)
# =============================================================================
def parse_data_tsv(tsv_path):
    """Read the design's _Data.tsv row. Returns a dict or None if absent/incomplete.

    TSV layout (DataExtract_worker):
      Printer, FileType, FileName, SKU, Theme, Date, H, M,
      [8x (grams, color)], ColorSwaps, ObjCount, ModelGrams, TotalGrams, TimeAdd
    The 5 summary columns are always last (indexed from the end); H/M follow the
    Date column (found by its m/d/yyyy pattern) - robust across TSV versions.
    """
    if not tsv_path or not os.path.isfile(tsv_path):
        return None
    try:
        with open(tsv_path, encoding="utf-8-sig", errors="replace") as fh:   # -sig strips the BOM
            last = [ln for ln in fh.read().splitlines() if ln.strip()][-1]
    except Exception:
        return None
    c = last.split("\t")
    if len(c) < 13:
        return None
    date_idx = next((i for i in (4, 5) if i < len(c) and DATE_RE.match(c[i].strip())), -1)
    if date_idx < 0 or (date_idx + 2) >= len(c) - 5:
        return None

    def fnum(s):
        try: return float(s.strip())
        except Exception: return None

    n = len(c)
    color_swaps = fnum(c[n - 5]); obj = fnum(c[n - 4])
    model_g = fnum(c[n - 3]); total_g = fnum(c[n - 2]); time_add = fnum(c[n - 1])
    H = fnum(c[date_idx + 1]) or 0.0
    M = fnum(c[date_idx + 2]) or 0.0
    time_h = H + M / 60.0
    if not obj or obj <= 0 or time_h <= 0:
        return None

    # per-colour slots sit between M and the 5 trailing summary columns
    colours = []
    i = date_idx + 3
    while i + 1 <= n - 6:
        g = fnum(c[i]); name = c[i + 1].strip() if i + 1 < n else ""
        if g and g > 0:
            colours.append({"grams": round(g, 2), "color": name})
        i += 2

    out = {
        "printer": c[0].strip(), "file_type": c[1].strip(), "file_name": c[2].strip(),
        "sku": c[3].strip() if len(c) > 3 else "", "theme": c[4].strip() if len(c) > 4 else "",
        "print_time_h": round(time_h, 2),
        "objects_pre_merge": int(obj),
        "color_changes": int(color_swaps) if color_swaps is not None else None,
        "model_material_g": round(model_g, 2) if model_g is not None else None,
        "total_material_g": round(total_g, 2) if total_g is not None else None,
        "colors_used": len(colours),
        "per_color_g": colours,
        "time_add_per_wig_min": round(time_add, 2) if time_add is not None else None,
    }
    if model_g and obj:
        out["filament_per_unit_g"] = round(model_g / obj, 2)
    if model_g and model_g > 0:
        out["time_per_gram_min"] = round(time_h * 60.0 / model_g, 2)   # complexity / fiddliness proxy
    if model_g is not None and total_g and total_g > 0:
        out["waste_per_unit_g"] = round((total_g - model_g) / obj, 2)
    out["throughput_wig_day"] = round(obj / time_h * 24.0, 1)
    return out


# =============================================================================
#  PART B - the sliced 3mf
# =============================================================================
def _read(zf, name):
    t = name.lower()
    for n in zf.namelist():
        if n.replace("\\", "/").lower() == t:
            return zf.read(n)
    return None

def _open(zf, name):
    t = name.lower()
    for n in zf.namelist():
        if n.replace("\\", "/").lower() == t:
            return zf.open(n)
    return None

def _poly(cfg, key):
    m = re.search(r'"%s"\s*:\s*\[(.*?)\]' % re.escape(key), cfg, re.S)
    if not m: return None
    pts = re.findall(r'"([\d.]+)x([\d.]+)"', m.group(1))
    return [(float(a), float(b)) for a, b in pts] or None

def _bbox(poly):
    xs = [p[0] for p in poly]; ys = [p[1] for p in poly]
    return (min(xs), min(ys), max(xs), max(ys))

def _area(r):
    return max(0.0, r[2] - r[0]) * max(0.0, r[3] - r[1]) if r else 0.0


def plate_utilization(zf, printer=None):
    """Plate utilization % from pick.png coverage over the available bed."""
    cfg = _read(zf, "Metadata/project_settings.config")
    pjb = _read(zf, "Metadata/plate_1.json")
    pkb = _read(zf, "Metadata/pick_1.png")
    if not (cfg and pjb and pkb):
        return {"utilization_error": "missing pick_1.png / plate_1.json / config"}
    cfg = cfg.decode("utf-8", "replace"); pj = json.loads(pjb)
    pa = _poly(cfg, "printable_area") or [(0, 0), (256, 0), (256, 256), (0, 256)]
    bx0, by0, bx1, by1 = _bbox(pa)
    bed_w, bed_h = bx1 - bx0, by1 - by0
    bed_area = bed_w * bed_h
    excl = _bbox(_poly(cfg, "bed_exclude_area")) if _poly(cfg, "bed_exclude_area") else None
    # calibration line is a per-printer machine constant; key by printer (X1C/P2S/H2S)
    calib = CALIBRATION_LINES.get((printer or "").upper())
    calib_estimated = False
    if calib is None:
        # Best-guess for non-verified beds (P2S / H2S / unknown): a front purge
        # strip from just past the exclusion zone to near the right edge, 14mm
        # tall. This formula reproduces the verified X1C value exactly. Tune per
        # printer later from a real design of that machine.
        excl_right = excl[2] if excl else 18.0
        calib = (excl_right, 0.0, max(excl_right + 10.0, bed_w - 10.0), 14.0)
        calib_estimated = True
    tower = None
    for o in pj.get("bbox_objects", []):
        if re.search("wipe|prime", str(o.get("name", "")), re.I):
            tower = tuple(float(v) for v in o["bbox"]); break

    pick = Image.open(io.BytesIO(pkb)).convert("RGBA")
    W, Hh = pick.size; mmppx, mmppy = bed_w / W, bed_h / Hh; px = pick.load()
    def in_tower(xp, yp):
        if not tower: return False
        mx = bx0 + xp * mmppx; my = by1 - yp * mmppy
        return tower[0] <= mx <= tower[2] and tower[1] <= my <= tower[3]
    obj_px = sum(1 for y in range(Hh) for x in range(W)
                 if px[x, y][3] > ALPHA_THRESHOLD and not in_tower(x, y))
    obj_area = obj_px * mmppx * mmppy
    available = bed_area - _area(excl) - _area(calib) - _area(tower)
    return {
        "plate_utilization_pct": round(obj_area / available * 100, 1) if available > 0 else 0.0,
        "object_area_mm2": round(obj_area, 1),
        "available_area_mm2": round(available, 1),
        "bed_area_mm2": round(bed_area, 1),
        "exclusion_area_mm2": round(_area(excl), 1),
        "printer": printer,
        "calibration_area_mm2": round(_area(calib), 1),
        "calibration_estimated": calib_estimated,
        "prime_tower_bbox": [round(v, 2) for v in tower] if tower else None,
        "prime_tower_area_mm2": round(_area(tower), 1),
        "objects_on_plate": sum(1 for o in pj.get("bbox_objects", [])
                                if not re.search("wipe|prime", str(o.get("name", "")), re.I)),
    }


def gcode_header(zf):
    """Print height + layer count (top ~60 lines of the gcode - instant)."""
    out = {}
    gh = _open(zf, "Metadata/plate_1.gcode")
    if not gh: return out
    sr = io.TextIOWrapper(gh, encoding="utf-8", errors="replace")
    for i, line in enumerate(sr):
        if "HEADER_BLOCK_END" in line or i > 60: break
        m = re.search(r"max_z_height:\s*([0-9.]+)", line)
        if m: out["print_height_mm"] = float(m.group(1))
        m = re.search(r"total layer number:\s*([0-9]+)", line)
        if m: out["total_layers"] = int(m.group(1))
    sr.detach()
    if out.get("print_height_mm") and out.get("total_layers"):
        out["effective_layer_height_mm"] = round(out["print_height_mm"] / out["total_layers"], 3)
    return out


def variable_layer_height(zf):
    out = {}
    lhp = _read(zf, "Metadata/layer_heights_profile.txt")
    if not lhp: return out
    nums = [float(x) for x in re.findall(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", lhp.decode("utf-8", "replace"))]
    heights = [v for v in nums if 0.04 < v < 0.6]
    if heights:
        out["variable_layer_height"] = (max(heights) - min(heights) > 0.01)
        # harvested but not shown in the readout (sampling-biased - see notes)
        out["layer_height_min_mm"] = round(min(heights), 3)
        out["layer_height_max_mm"] = round(max(heights), 3)
        out["layer_height_mean_mm"] = round(sum(heights) / len(heights), 3)
    return out


def gcode_body(zf):
    """Single streaming pass over plate_1.gcode for the motion/feature metrics."""
    gh = _open(zf, "Metadata/plate_1.gcode")
    if not gh: return {"gcode_body_error": "no plate_1.gcode"}
    sr = io.TextIOWrapper(gh, encoding="utf-8", errors="replace")
    feature = "other"; feat_fil = {}
    extrude_dist = travel_dist = 0.0
    travel_moves = retractions = zhops = outer_loops = toolchanges = layers = 0
    e_relative = True; lastE = 0.0
    x = y = z = None
    obj_fil = {}; cur_obj = None
    layer_R = []; last_R = None; in_cfg = False

    for s in sr:
        c = s[0] if s else ""
        if c == ";":
            if s.startswith("; FEATURE:"):
                feature = _feature_bucket(s[10:].strip())
                if feature == "outer_wall": outer_loops += 1
            elif s.startswith("; CHANGE_LAYER"):
                layers += 1
                if last_R is not None: layer_R.append(last_R)
            elif s.startswith("; start printing object"):
                m = re.search(r"id:\s*(\S+)", s); cur_obj = m.group(1) if m else None
            elif s.startswith("; stop printing object"):
                cur_obj = None
            elif s.startswith("; CONFIG_BLOCK_START"): in_cfg = True
            elif s.startswith("; CONFIG_BLOCK_END"): in_cfg = False
            continue
        if c == "M":
            if s.startswith("M83"): e_relative = True
            elif s.startswith("M82"): e_relative = False
            elif s.startswith("M73 P"):
                m = re.search(r"\bR([0-9.]+)", s)
                if m: last_R = float(m.group(1))
            continue
        if c == "T" and len(s) > 1 and s[1].isdigit() and not in_cfg:
            toolchanges += 1
            continue
        if c == "G" and (s.startswith("G1 ") or s.startswith("G0 ")):
            nx = ny = nz = e = None
            for tok in s.split():
                t0 = tok[0]
                try:
                    if t0 == "X": nx = float(tok[1:])
                    elif t0 == "Y": ny = float(tok[1:])
                    elif t0 == "Z": nz = float(tok[1:])
                    elif t0 == "E": e = float(tok[1:])
                except ValueError:
                    pass
            de = 0.0
            if e is not None:
                if e_relative: de = e
                else: de = e - lastE; lastE = e
            dxy = 0.0
            if nx is not None or ny is not None:
                px_, py_ = (nx if nx is not None else x), (ny if ny is not None else y)
                if x is not None and y is not None:
                    dxy = math.hypot(px_ - x, py_ - y)
                x, y = px_, py_
            zup = (nz is not None and z is not None and nz > z + 1e-6)
            if nz is not None: z = nz
            if de > 1e-9:
                extrude_dist += dxy
                feat_fil[feature] = feat_fil.get(feature, 0.0) + de
                if cur_obj is not None:
                    obj_fil[cur_obj] = obj_fil.get(cur_obj, 0.0) + de
            else:
                if dxy > 1e-9:
                    travel_dist += dxy; travel_moves += 1
                    if zup: zhops += 1
                if de < -1e-9:
                    retractions += 1
    sr.detach()

    total = sum(feat_fil.values()) or 1.0
    res = {
        "travel_distance_mm": round(travel_dist, 1),
        "travel_moves": travel_moves,
        "travel_ratio_pct": round(travel_dist / (extrude_dist + travel_dist) * 100, 1) if (extrude_dist + travel_dist) else 0.0,
        "retractions": retractions,
        "z_hops": zhops,
        "layers_gcode": layers,
        # --- harvested but not shown in the readout (for later use) ---
        "total_extruded_filament_mm": round(sum(feat_fil.values()), 1),
        "feature_mix_pct": {k: round(v / total * 100, 1) for k, v in sorted(feat_fil.items(), key=lambda kv: -kv[1])},
        "feature_filament_mm": {k: round(v, 1) for k, v in feat_fil.items()},
        "prime_tower_filament_mm": round(feat_fil.get("prime_tower", 0.0), 1),
        "support_filament_mm": round(feat_fil.get("support", 0.0), 1),
        "outer_wall_loops": outer_loops,
        "tool_changes_gcode": toolchanges,
    }
    if layers:
        res["outer_loops_per_layer"] = round(outer_loops / layers, 1)
    if len(layer_R) >= 3:
        deltas = [layer_R[i] - layer_R[i + 1] for i in range(len(layer_R) - 1) if layer_R[i] >= layer_R[i + 1]]
        if deltas:
            res["layer_time_min_mean"] = round(sum(deltas) / len(deltas), 2)
            res["layer_time_min_max"] = round(max(deltas), 2)
    if obj_fil:
        v = list(obj_fil.values())
        res["object_count_gcode"] = len(v)
        res["object_filament_mm_mean"] = round(sum(v) / len(v), 1)
        res["object_filament_mm_min"] = round(min(v), 1)
        res["object_filament_mm_max"] = round(max(v), 1)
    return res


# =============================================================================
def extract_design(folder, want_body=True):
    """Run PART A + PART B for one design folder and return the result dict."""
    _, tsv_path, g3_path = find_design_files(folder)
    result = {"design": os.path.basename(folder.rstrip("\\/")), "folder": folder}
    part_a = parse_data_tsv(tsv_path)
    result["part_a_source"] = os.path.basename(tsv_path) if tsv_path else None
    result["part_a"] = part_a if part_a else {"error": "no usable _Data.tsv row found"}
    printer = (part_a or {}).get("printer") or os.path.basename(folder.rstrip("\\/")).split("_")[0]
    result["part_b_source"] = os.path.basename(g3_path) if g3_path else None
    part_b = {}
    if g3_path:
        zf = zipfile.ZipFile(g3_path)
        part_b.update(plate_utilization(zf, printer))
        part_b.update(gcode_header(zf))
        part_b.update(variable_layer_height(zf))
        if want_body:
            part_b.update(gcode_body(zf))
    else:
        part_b = {"error": "no *Full.gcode.3mf found"}
    if part_a and part_a.get("color_changes") is not None and part_b.get("total_layers"):
        part_b["color_changes_per_layer"] = round(part_a["color_changes"] / part_b["total_layers"], 2)
    result["part_b"] = part_b
    return result


def print_readout(result, no_body):
    a, b = result["part_a"], result["part_b"]
    def gv(d, k, dflt="-"): return d.get(k, dflt)
    print("=" * 60)
    print("DESIGN: %s" % result["design"])
    print("=" * 60)
    print("[ PART A  -  from %s ]" % (result["part_a_source"] or "(no data file)"))
    if "error" in a:
        print("  %s" % a["error"])
    else:
        print("  Print time: %s h    Objects (pre-merge): %s    Throughput: %s wig/day"
              % (gv(a, "print_time_h"), gv(a, "objects_pre_merge"), gv(a, "throughput_wig_day")))
        print("  Total material: %s g    Model material: %s g    Filament/unit: %s g    Time/gram: %s min"
              % (gv(a, "total_material_g"), gv(a, "model_material_g"), gv(a, "filament_per_unit_g"), gv(a, "time_per_gram_min")))
        print("  Color changes: %s    Waste/unit: %s g"
              % (gv(a, "color_changes"), gv(a, "waste_per_unit_g")))
    print("[ PART B  -  from %s ]" % (result["part_b_source"] or "(no 3mf)"))
    if "error" in b:
        print("  %s" % b["error"])
    else:
        print("  Plate utilization: %s%%   (object %s / available %s mm2)"
              % (gv(b, "plate_utilization_pct"), gv(b, "object_area_mm2"), gv(b, "available_area_mm2")))
        print("  Print height: %s mm    Layers: %s    Effective layer height: %s mm"
              % (gv(b, "print_height_mm"), gv(b, "total_layers"), gv(b, "effective_layer_height_mm")))
        print("  Color changes per layer: %s   (%s changes / %s layers)"
              % (gv(b, "color_changes_per_layer"), gv(a, "color_changes"), gv(b, "total_layers")))
        print("  Variable layer height: %s" % gv(b, "variable_layer_height"))
        if not no_body:
            print("  Travel: %s mm / %s moves (%s%% of motion)   Retractions: %s   Z-hops: %s"
                  % (gv(b, "travel_distance_mm"), gv(b, "travel_moves"), gv(b, "travel_ratio_pct"),
                     gv(b, "retractions"), gv(b, "z_hops")))
            print("  Per-layer time: mean %s min (max %s)   Per-object filament: mean %s mm (%s-%s)"
                  % (gv(b, "layer_time_min_mean"), gv(b, "layer_time_min_max"),
                     gv(b, "object_filament_mm_mean"), gv(b, "object_filament_mm_min"), gv(b, "object_filament_mm_max")))
        if b.get("calibration_estimated"):
            sys.stderr.write("\nNOTE: calibration line is an ESTIMATE for this printer (not verified yet); "
                             "utilization may shift slightly once tuned.\n")


def print_compact(results):
    print("Found %d design(s):\n" % len(results))
    print("  %-34s %-4s %-10s %7s %9s %7s %9s" % ("design", "prn", "type", "util%", "wig/day", "t/g", "chg/lyr"))
    print("  " + "-" * 86)
    for r in results:
        a = r.get("part_a", {}); b = r.get("part_b", {})
        pr = b.get("printer") or a.get("printer") or "?"
        print("  %-34s %-4s %-10s %7s %9s %7s %9s"
              % (r["design"][:34], pr, str(a.get("file_type", "?"))[:10],
                 b.get("plate_utilization_pct", "-"), a.get("throughput_wig_day", "-"),
                 a.get("time_per_gram_min", "-"), b.get("color_changes_per_layer", "-")))


def flatten_result(r):
    row = {"design": r.get("design"), "folder": os.path.normpath(r.get("folder", "")),
           "part_a_source": r.get("part_a_source"), "part_b_source": r.get("part_b_source")}
    for sec in ("part_a", "part_b"):
        d = r.get(sec, {})
        if isinstance(d, dict):
            for k, v in d.items():
                row[k] = json.dumps(v) if isinstance(v, (dict, list)) else v
    return row


def load_existing_csv(path):
    if not os.path.isfile(path):
        return []
    import csv
    with open(path, newline="", encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


def write_csv(flat_rows, path):
    import csv
    keys = []
    for row in flat_rows:
        for k in row:
            if k not in keys: keys.append(k)
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=keys, extrasaction="ignore")
        w.writeheader()
        for row in flat_rows:
            w.writerow(row)


def prompt_select(label, items):
    """items: list of (value, count). Returns the chosen set of values."""
    print("\n%s found:" % label)
    for i, (v, c) in enumerate(items, 1):
        print("  %2d) %-14s (%d design%s)" % (i, v, c, "" if c == 1 else "s"))
    raw = input("Select %s  [numbers comma-separated, or A for all]: " % label.lower()).strip()
    if raw.upper() in ("", "A", "ALL"):
        return {v for v, _ in items}
    chosen = set()
    for tok in raw.replace(",", " ").split():
        if tok.isdigit() and 1 <= int(tok) <= len(items):
            chosen.add(items[int(tok) - 1][0])
    return chosen or {v for v, _ in items}


def interactive_filter(folders):
    """Classify each design (printer + type, cheaply from the TSV / name), then
    let the user multi-select which printers and types to harvest."""
    from collections import Counter
    classified = []
    for f in folders:
        _, tsv_path, _ = find_design_files(f)
        a = parse_data_tsv(tsv_path) if tsv_path else None
        printer = (a or {}).get("printer") or os.path.basename(f.rstrip("\\/")).split("_")[0]
        ftype = (a or {}).get("file_type") or "Unknown"
        classified.append((f, printer, ftype))
    pr_items = sorted(Counter(p for _, p, _ in classified).items())
    sel_pr = prompt_select("Printers", pr_items)
    tp_items = sorted(Counter(t for _, p, t in classified if p in sel_pr).items())
    sel_tp = prompt_select("Types", tp_items)
    return [f for f, p, t in classified if p in sel_pr and t in sel_tp]


def main():
    # The Windows cmd console is cp1252; force utf-8 so a non-cp1252 character in
    # any design name / color / value can never crash a print or progress write.
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    ap = argparse.ArgumentParser(description="Per-design metrics: PART A (data file) + PART B (3mf files). "
                                             "Accepts design folders, parent folders (searched recursively), "
                                             "or *.3mf files - one or more.")
    ap.add_argument("paths", nargs="+", help="Design/parent folders (searched recursively) or *Full.gcode.3mf files.")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--no-body", action="store_true", help="Skip the heavy gcode-body pass.")
    ap.add_argument("--full", action="store_true",
                    help="Harvest mode: force the full gcode-body parse and emit the complete superset (implies --json).")
    ap.add_argument("--csv", metavar="NAME.csv",
                    help="Accumulate all designs' full metrics into BambuScripts/data/NAME.csv "
                         "(appends new designs, skips ones already harvested).")
    ap.add_argument("--overwrite", action="store_true",
                    help="With --csv: start the file fresh instead of appending.")
    ap.add_argument("--select", action="store_true",
                    help="Interactively pick which printers / types to harvest before parsing.")
    args = ap.parse_args()
    if args.full:
        args.json = True
        args.no_body = False

    folders = find_design_folders(args.paths)
    if not folders:
        sys.stderr.write("No design folders found (need a *Full.gcode.3mf). Nothing to do.\n")
        sys.exit(1)

    if args.select:
        print("Scanning %d design(s)..." % len(folders))
        folders = interactive_filter(folders)
        print("\n-> %d design(s) selected.\n" % len(folders))
        if not folders:
            sys.stderr.write("Nothing selected. Done.\n")
            sys.exit(0)

    # --- CSV harvest mode: accumulate, skip already-harvested (parse once) ---
    if args.csv:
        data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
        os.makedirs(data_dir, exist_ok=True)
        out_path = os.path.join(data_dir, os.path.basename(args.csv))
        existing = [] if args.overwrite else load_existing_csv(out_path)
        done = {os.path.normpath(r.get("folder", "")) for r in existing}
        todo = [f for f in folders if os.path.normpath(f) not in done]
        if len(todo) < len(folders):
            sys.stderr.write("Skipping %d already-harvested; %d new to parse.\n" % (len(folders) - len(todo), len(todo)))
        new_rows = []
        for i, folder in enumerate(todo, 1):
            sys.stderr.write("[%d/%d] %s\n" % (i, len(todo), os.path.basename(folder)))
            try:
                new_rows.append(flatten_result(extract_design(folder, want_body=not args.no_body)))
            except Exception as e:
                sys.stderr.write("  ERROR on %s: %s\n" % (folder, e))
        write_csv(existing + new_rows, out_path)
        sys.stderr.write("\nHarvested %d new; %d total -> %s\n" % (len(new_rows), len(existing) + len(new_rows), out_path))
        return

    # --- readout / json mode ---
    results = []
    for i, folder in enumerate(folders, 1):
        if len(folders) > 1:
            sys.stderr.write("[%d/%d] %s\n" % (i, len(folders), os.path.basename(folder)))
        try:
            results.append(extract_design(folder, want_body=not args.no_body))
        except Exception as e:
            sys.stderr.write("  ERROR on %s: %s\n" % (folder, e))

    if args.json:
        print(json.dumps(results[0] if len(results) == 1 else results, indent=2))
    elif len(results) == 1:
        print_readout(results[0], args.no_body)
    else:
        print_compact(results)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        import traceback
        log = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "last_error.log")
        try:
            os.makedirs(os.path.dirname(log), exist_ok=True)
            with open(log, "w", encoding="utf-8") as fh:
                fh.write(traceback.format_exc())
            sys.stderr.write("\nCRASHED - full traceback written to:\n  %s\n" % log)
        except Exception:
            traceback.print_exc()
        sys.exit(1)
