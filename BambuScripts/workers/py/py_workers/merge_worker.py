#!/usr/bin/env python3
"""
merge_worker.py - N-way 3MF merge worker.

Usage:
    python merge_worker.py --work-dir <extracted_3mf_dir>
                           --input-path <input.3mf>
                           --output-path <merged_temp.3mf>
                           [--do-colors 0|1]

Mirrors PS1 merge_3mf_worker.ps1.
All transform math is a faithful port of the PS1 column-major matrix routines.
"""
from __future__ import annotations

import argparse
import copy
import math
import re
import sys
import uuid
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS_CORE = 'http://schemas.microsoft.com/3dmanufacturing/core/2015/02'
NS_PROD = 'http://schemas.microsoft.com/3dmanufacturing/production/2015/06'

ET.register_namespace('',  NS_CORE)
ET.register_namespace('p', NS_PROD)


def _register_ns(path: Path) -> None:
    """Register every xmlns prefix found in an XML file so ET preserves them on write."""
    try:
        text = path.read_text(encoding='utf-8', errors='replace')[:8000]
        m = re.search(r'\bxmlns=["\']([^"\']+)["\']', text)
        if m:
            ET.register_namespace('', m.group(1))
        for prefix, uri in re.findall(r'\bxmlns:(\w[\w.-]*)=["\']([^"\']+)["\']', text):
            ET.register_namespace(prefix, uri)
    except Exception:
        pass


# ── Transform math (column-major, matching PS1 exactly) ───────────────────────

def parse_tx(s: str) -> list[float]:
    """Parse a space-separated transform string to 12 floats (identity if blank)."""
    if not s or not s.strip():
        return [1,0,0, 0,1,0, 0,0,1, 0,0,0]
    return [float(v) for v in s.strip().split()]


def fmt_tx(v: list[float]) -> str:
    return ' '.join(f'{x:.15g}' for x in v)


def mul_tx(A: list[float], B: list[float]) -> list[float]:
    """Multiply two 3x3+translation column-major matrices."""
    a00,a10,a20, a01,a11,a21, a02,a12,a22, atx,aty,atz = A
    b00,b10,b20, b01,b11,b21, b02,b12,b22, btx,bty,btz = B
    return [
        a00*b00 + a01*b10 + a02*b20,  a10*b00 + a11*b10 + a12*b20,  a20*b00 + a21*b10 + a22*b20,
        a00*b01 + a01*b11 + a02*b21,  a10*b01 + a11*b11 + a12*b21,  a20*b01 + a21*b11 + a22*b21,
        a00*b02 + a01*b12 + a02*b22,  a10*b02 + a11*b12 + a12*b22,  a20*b02 + a21*b12 + a22*b22,
        a00*btx + a01*bty + a02*btz + atx,
        a10*btx + a11*bty + a12*btz + aty,
        a20*btx + a21*bty + a22*btz + atz,
    ]


def inv_tx(A: list[float]) -> list[float]:
    """Invert a rigid-body (rotation + translation) transform."""
    ir00,ir10,ir20, ir01,ir11,ir21, ir02,ir12,ir22, tx,ty,tz = A
    return [
        ir00, ir01, ir02,
        ir10, ir11, ir12,
        ir20, ir21, ir22,
        -(ir00*tx + ir01*ty + ir02*tz),
        -(ir10*tx + ir11*ty + ir12*tz),
        -(ir20*tx + ir21*ty + ir22*tz),
    ]


# ── XML helpers ───────────────────────────────────────────────────────────────

def _find_file(base: Path, rel: str) -> Path | None:
    normalized = rel.replace('\\', '/').lstrip('/')
    p = base / normalized
    if p.exists():
        return p
    p2 = base / rel
    if p2.exists():
        return p2
    return None


def _find_all_with_tag(root: ET.Element, local_name: str) -> list[ET.Element]:
    """Find all elements with the given local name regardless of namespace."""
    results = root.findall(f'.//{{{NS_CORE}}}{local_name}')
    results += root.findall(f'.//{local_name}')
    return results


# ── Merge plan ────────────────────────────────────────────────────────────────

def get_merge_plan(total: int, lone: int, max_slots: int, ignored_count: int) -> list[int]:
    """
    Compute N-way merge grouping plan.  Mirrors PS1 Get-MergePlan.
    Returns a list of group sizes.
    """
    pool = total - lone
    effective_max = max_slots - lone - ignored_count

    if pool <= 0:
        return []

    for g in range(2, pool + 1):
        best_a, best_b = -1, -1
        period = g + 1
        for a0 in range(0, g + 1):
            rem = pool - a0 * g
            if rem < 0:
                break
            if rem % (g + 1) != 0:
                continue
            a = a0
            while a <= pool // g:
                b = (pool - a * g) // (g + 1)
                if (a + b) <= effective_max:
                    if a > best_a:
                        best_a, best_b = a, b
                a += period
            break
        if best_a >= 0:
            return [g] * best_a + [(g + 1)] * best_b

    return [pool]


# ── Main merge ────────────────────────────────────────────────────────────────

def merge(work_dir: Path, input_path: Path, output_path: Path, do_colors: bool = False) -> int:
    """
    Perform N-way merge on extracted 3MF work_dir contents and write output_path.
    Returns 0 on success, 1 on failure.
    """
    # ── Persistent debug log (survives log-file deletion by caller) ───────────
    _dbg_path = output_path.parent / 'merge_debug.txt'
    _dbg = open(_dbg_path, 'w', encoding='utf-8', buffering=1)
    def _log(msg: str) -> None:
        print(msg)
        _dbg.write(msg + '\n')
        _dbg.flush()

    _log(f'[merge] input_path  = {input_path}')
    _log(f'[merge] output_path = {output_path}')
    _log(f'[merge] work_dir    = {work_dir}')

    # Locate files
    model_files = list(work_dir.rglob('3dmodel.model'))
    if not model_files:
        _log('[merge] ERROR: 3dmodel.model not found')
        _dbg.close()
        return 1
    model_file = model_files[0]
    objects_dir   = work_dir / '3D' / 'Objects'
    rels_path     = _find_file(work_dir, '3D/_rels/3dmodel.model.rels')
    settings_path = _find_file(work_dir, 'Metadata/model_settings.config')
    cut_info_path = _find_file(work_dir, 'Metadata/cut_information.xml')
    slice_info_path = _find_file(work_dir, 'Metadata/slice_info.config')
    vlh_path      = work_dir / 'Metadata' / 'layer_heights_profile.txt'

    # Parse main model XML (register namespaces first so ET preserves them on write)
    _register_ns(model_file)
    tree = ET.parse(model_file)
    root = tree.getroot()

    build_el = root.find(f'{{{NS_CORE}}}build')
    resources_el = root.find(f'{{{NS_CORE}}}resources')
    if build_el is None or resources_el is None:
        _log('[merge] ERROR: malformed 3dmodel.model (missing build or resources)')
        _dbg.close()
        return 1

    build_items = list(build_el)
    obj_by_id: dict[str, ET.Element] = {
        obj.get('id', ''): obj
        for obj in resources_el
        if obj.get('id')
    }

    # Parse model_settings.config
    has_settings = bool(settings_path and settings_path.exists())
    sett_obj_by_id: dict[str, ET.Element] = {}
    sett_root: ET.Element | None = None
    sett_tree: ET.ElementTree | None = None
    if has_settings:
        try:
            _register_ns(settings_path)
            sett_tree = ET.parse(settings_path)
            sett_root = sett_tree.getroot()
            # Only top-level <object> children of config root
            for node in sett_root:
                if node.tag in ('object',) or node.tag.endswith('}object'):
                    nid = node.get('id')
                    if nid:
                        sett_obj_by_id[nid] = node
        except Exception:
            has_settings = False

    # VLH: dominant layer heights profile string
    vlh_data_string: str | None = None
    if vlh_path.exists():
        try:
            lines = vlh_path.read_text(encoding='utf-8', errors='replace').splitlines()
            counts: dict[str, int] = {}
            for line in lines:
                idx = line.find('|')
                if idx >= 0:
                    val = line[idx + 1:]
                    counts[val] = counts.get(val, 0) + 1
            if counts:
                vlh_data_string = max(counts, key=lambda k: counts[k])
        except Exception:
            pass

    # ── Purge off-plate objects ───────────────────────────────────────────────
    plate_assigned: set[str] = set()
    if has_settings and sett_root is not None:
        for inst in sett_root.findall('.//plate/model_instance'):
            meta = inst.find('metadata[@key="object_id"]')
            if meta is None:
                meta = inst.find(f'{{{NS_CORE}}}metadata[@key="object_id"]')
            if meta is not None:
                plate_assigned.add(meta.get('value', ''))

    killed_ids: set[str] = set()
    valid_items: list[ET.Element] = []
    for item in build_items:
        oid = item.get('objectid', '')
        tx = parse_tx(item.get('transform', ''))
        x, y = tx[9], tx[10]
        off_coord = x < -50 or x > 300 or y < -50 or y > 300
        off_settings = bool(plate_assigned) and oid not in plate_assigned
        if off_coord or off_settings:
            killed_ids.add(oid)
            build_el.remove(item)
        else:
            valid_items.append(item)
    build_items = valid_items

    # Closure: expand protected IDs to include referenced components
    protected: set[str] = {item.get('objectid', '') for item in build_items}
    added = True
    while added:
        added = False
        for oid in list(protected):
            obj = obj_by_id.get(oid)
            if obj is not None:
                for comp in obj.findall(f'.//{{{NS_CORE}}}component'):
                    cid = comp.get('objectid', '')
                    if cid and cid not in protected:
                        protected.add(cid)
                        added = True

    for oid in list(obj_by_id.keys()):
        if oid not in protected:
            killed_ids.add(oid)
            try:
                resources_el.remove(obj_by_id[oid])
            except ValueError:
                pass
            del obj_by_id[oid]
            if has_settings and sett_root is not None:
                if oid in sett_obj_by_id:
                    try:
                        sett_root.remove(sett_obj_by_id[oid])
                    except ValueError:
                        pass
                # Also remove any node with object_id attribute matching this id
                for node in list(sett_root.iter()):
                    if node.get('object_id') == oid:
                        for parent in sett_root.iter():
                            if node in list(parent):
                                try:
                                    parent.remove(node)
                                except ValueError:
                                    pass
                                break

    # ── Outlier detection: separate merge targets from text/version objects ───
    fc_map: dict[str, int] = {}
    for item in build_items:
        oid = item.get('objectid', '')
        fc = 'unknown'
        if has_settings and oid in sett_obj_by_id:
            fc_node = sett_obj_by_id[oid].find('metadata[@face_count]')
            if fc_node is not None:
                fc = fc_node.get('face_count', 'unknown')
        fc_map[fc] = fc_map.get(fc, 0) + 1

    majority_fc = max(fc_map, key=lambda k: fc_map[k]) if fc_map else 'unknown'

    merge_items: list[ET.Element] = []
    ignored_items: list[ET.Element] = []
    for item in build_items:
        oid = item.get('objectid', '')
        is_target = True
        if has_settings and oid in sett_obj_by_id:
            fc_node = sett_obj_by_id[oid].find('metadata[@face_count]')
            fc = fc_node.get('face_count', 'unknown') if fc_node is not None else 'unknown'
            if fc != majority_fc:
                is_target = False
            name_node = sett_obj_by_id[oid].find('metadata[@key="name"]')
            if name_node is not None:
                v = name_node.get('value', '')
                if 'text' in v.lower() or 'version' in v.lower():
                    is_target = False
        (merge_items if is_target else ignored_items).append(item)

    # ── Compute merge plan ────────────────────────────────────────────────────
    total_items = len(merge_items)
    lone = 2 if total_items % 2 == 0 else 1
    merge_plan = get_merge_plan(total_items, lone, 64, len(ignored_items))

    _log(f'[merge] sett_obj_by_id count: {len(sett_obj_by_id)}, sample ids: {list(sett_obj_by_id.keys())[:5]}')
    _log(f'[merge] plate_assigned count: {len(plate_assigned)}, sample: {list(plate_assigned)[:5]}')
    _log(f'[merge] build_items after purge: {len(build_items)}')
    _log(f'[merge] fc_map: {fc_map}')
    _log(f'[merge] majority_fc: {majority_fc!r}')
    _log(f'[merge] Total on plate: {len(build_items)}, Ignored: {len(ignored_items)}, '
         f'Merge targets: {total_items}, Lone: {lone}, Plan: {merge_plan}')
    _log(f'[merge] effective_max (64-lone-ignored): {64 - lone - len(ignored_items)}')

    # ── Mesh file tracker ─────────────────────────────────────────────────────
    source_to_master: dict[str, str] = {}
    used_model_paths: set[str] = set()

    # ── Dynamic merge loop ────────────────────────────────────────────────────
    cursor = 0
    survivor_faces: dict[str, int] = {}
    survivor_names: dict[str, str] = {}
    identify_id_counter = 0
    survivor_identify_ids: dict[str, int] = {}
    obj_uuid_counter = 1

    for group_size in merge_plan:
        group_items  = merge_items[cursor:cursor + group_size]
        group_objs   = [obj_by_id[item.get('objectid', '')] for item in group_items]
        group_txs    = [parse_tx(item.get('transform', '')) for item in group_items]
        cursor += group_size

        id_survivor = group_items[0].get('objectid', '')

        # Centroid
        sum_x = sum(tx[9] for tx in group_txs) / group_size
        sum_y = sum(tx[10] for tx in group_txs) / group_size
        sum_z = sum(tx[11] for tx in group_txs) / group_size
        tx_new = [1,0,0, 0,1,0, 0,0,1, sum_x, sum_y, sum_z]
        group_items[0].set('transform', fmt_tx(tx_new))
        # Re-parse to get exact same floats as stored (mirrors PS1 Parse-Tx after Fmt-Tx)
        tx_new = parse_tx(fmt_tx(tx_new))
        inv_tx_new = inv_tx(tx_new)

        merged_comps: list[ET.Element] = []
        merged_parts: list[ET.Element] = []
        comp_base_suffix = '-' + str(uuid.uuid4())[9:]
        comp_index = obj_uuid_counter * 65536

        for k in range(group_size):
            obj = group_objs[k]
            orig_tx = group_txs[k]
            member_id = group_items[k].get('objectid', '')

            comp_list = obj.findall(f'{{{NS_CORE}}}components/{{{NS_CORE}}}component')
            part_list: list[ET.Element] = []
            if has_settings and sett_root is not None and member_id in sett_obj_by_id:
                part_list = list(sett_obj_by_id[member_id].findall('part'))

            for i, comp in enumerate(comp_list):
                # Resolve geometry path
                c_path = (
                    comp.get(f'{{{NS_PROD}}}path') or
                    comp.get('p:path') or
                    comp.get('path') or ''
                )
                if c_path:
                    if not c_path.startswith('/'):
                        c_path = '/' + c_path
                    if c_path not in source_to_master:
                        source_to_master[c_path] = c_path
                        used_model_paths.add(c_path)
                    resolved = source_to_master[c_path]
                else:
                    resolved = None

                comp_tx  = parse_tx(comp.get('transform', ''))
                baked_tx = mul_tx(inv_tx_new, mul_tx(orig_tx, comp_tx))

                new_comp = ET.Element(f'{{{NS_CORE}}}component')
                if resolved:
                    new_comp.set(f'{{{NS_PROD}}}path', resolved)
                new_comp.set('objectid', comp.get('objectid', ''))
                comp_uuid = format(comp_index, '08x') + comp_base_suffix
                new_comp.set(f'{{{NS_PROD}}}UUID', comp_uuid)
                comp_index += 1
                new_comp.set('transform', fmt_tx(baked_tx))

                new_part: ET.Element | None = None
                if has_settings and sett_root is not None and i < len(part_list):
                    new_part = copy.deepcopy(part_list[i])
                    mat_node = new_part.find('metadata[@key="matrix"]')
                    if mat_node is None:
                        mat_node = ET.SubElement(new_part, 'metadata')
                        mat_node.set('key', 'matrix')
                    mat_node.set('value', fmt_tx(baked_tx))

                merged_comps.append(new_comp)
                if new_part is not None:
                    merged_parts.append(new_part)

        # Tally faces
        total_faces = 0
        for p in merged_parts:
            mesh_stat = p.find('mesh_stat')
            if mesh_stat is not None:
                try:
                    total_faces += int(mesh_stat.get('face_count', '0'))
                except ValueError:
                    pass
        if total_faces == 0:
            for k in range(group_size):
                mid = group_items[k].get('objectid', '')
                if has_settings and mid in sett_obj_by_id:
                    fc_node = sett_obj_by_id[mid].find('metadata[@face_count]')
                    if fc_node is not None:
                        try:
                            total_faces += int(fc_node.get('face_count', '0'))
                        except ValueError:
                            pass

        part_count = len(merged_parts) or 36 * group_size
        id_gap = round(442 * (part_count / 36))
        identify_id_counter += id_gap
        survivor_identify_ids[id_survivor] = identify_id_counter
        survivor_faces[id_survivor] = total_faces
        survivor_names[id_survivor] = f'MergedGroup_{group_size}'

        # Replace components on survivor object
        comps_el = ET.Element(f'{{{NS_CORE}}}components')
        for c in merged_comps:
            comps_el.append(c)
        old_comps = group_objs[0].find(f'{{{NS_CORE}}}components')
        if old_comps is not None:
            group_objs[0].remove(old_comps)
        group_objs[0].append(comps_el)

        obj_uuid_str = format(obj_uuid_counter, '08x') + '-71cb-4c03-9d28-80fed5dfa1dc'
        group_objs[0].set(f'{{{NS_PROD}}}UUID', obj_uuid_str)
        obj_uuid_counter += 1

        # Update model_settings survivor
        if has_settings and sett_root is not None and id_survivor in sett_obj_by_id:
            s_surv = sett_obj_by_id[id_survivor]
            name_node = s_surv.find('metadata[@key="name"]')
            if name_node is not None:
                name_node.set('value', survivor_names[id_survivor])
            fc_node = s_surv.find('metadata[@face_count]')
            if fc_node is not None:
                fc_node.set('face_count', str(total_faces))
            # Replace parts
            for part in list(s_surv.findall('part')):
                s_surv.remove(part)
            for pi, p in enumerate(merged_parts):
                comp_obj_id = merged_comps[pi].get('objectid', '')
                p.set('id', comp_obj_id)
                pn = p.find('metadata[@key="name"]')
                if pn is not None:
                    pn.set('value', f'MergedPart_{comp_obj_id}')
                s_surv.append(p)
            # Kill other members from settings
            for k in range(1, group_size):
                mid = group_items[k].get('objectid', '')
                killed_ids.add(mid)
                if mid in sett_obj_by_id:
                    try:
                        sett_root.remove(sett_obj_by_id[mid])
                    except ValueError:
                        pass
            # Assemble
            assemble = sett_root.find('.//assemble')
            if assemble is not None:
                asm_surv = assemble.find(f'assemble_item[@object_id="{id_survivor}"]')
                if asm_surv is not None:
                    sR = group_txs[0]
                    asm_surv.set('transform', (
                        f'{sR[0]} {sR[1]} {sR[2]} {sR[3]} {sR[4]} {sR[5]} '
                        f'{sR[6]} {sR[7]} {sR[8]} {tx_new[9]} {tx_new[10]} {tx_new[11]}'
                    ))
                for k in range(1, group_size):
                    mid = group_items[k].get('objectid', '')
                    for ai in list(assemble):
                        if ai.get('object_id') == mid:
                            assemble.remove(ai)

        # Remove non-survivor items from build + resources
        for k in range(1, group_size):
            try:
                build_el.remove(group_items[k])
            except ValueError:
                pass
            if group_objs[k] is not None:
                try:
                    resources_el.remove(group_objs[k])
                except ValueError:
                    pass

    # ── Lone items: assign UUIDs ──────────────────────────────────────────────
    lone_counter = 1
    for li in range(len(merge_items) - lone, len(merge_items)):
        lone_id  = merge_items[li].get('objectid', '')
        lone_obj = obj_by_id.get(lone_id)
        if lone_obj is not None:
            obj_uuid_str = format(obj_uuid_counter, '08x') + '-71cb-4c03-9d28-80fed5dfa1dc'
            lone_obj.set(f'{{{NS_PROD}}}UUID', obj_uuid_str)
        obj_uuid_counter += 1
        identify_id_counter += 442
        survivor_identify_ids[lone_id] = identify_id_counter

        if has_settings and sett_root is not None and lone_id in sett_obj_by_id:
            ln = sett_obj_by_id[lone_id].find('metadata[@key="name"]')
            if ln is not None:
                ln.set('value', f'{ln.get("value", "")}_Lone_{lone_counter}')

        if lone_obj is not None:
            for lc in lone_obj.findall(f'{{{NS_CORE}}}components/{{{NS_CORE}}}component'):
                lc_path = (
                    lc.get(f'{{{NS_PROD}}}path') or
                    lc.get('p:path') or
                    lc.get('path') or ''
                )
                if lc_path:
                    if not lc_path.startswith('/'):
                        lc_path = '/' + lc_path
                    if lc_path not in source_to_master:
                        source_to_master[lc_path] = lc_path
                        used_model_paths.add(lc_path)
                    lc.set(f'{{{NS_PROD}}}path', source_to_master[lc_path])
        lone_counter += 1

    # ── Clean killed instances from plate config ──────────────────────────────
    if has_settings and sett_root is not None:
        plate_el = sett_root.find('.//plate')
        if plate_el is not None:
            for inst in list(plate_el.findall('model_instance')):
                meta = inst.find('metadata[@key="object_id"]')
                if meta is not None and meta.get('value') in killed_ids:
                    plate_el.remove(inst)

    # ── GLOBAL OBJECT ID RENUMBERING (fixes "0 objects" parse bug in Bambu) ───
    # Internal mesh objects must come before printable assembly objects.
    # Rebuild as sequential 1..N so Bambu Studio can parse the file correctly.
    printable_ids_set: set[str] = set()
    if has_settings and sett_root is not None:
        for obj_node in sett_root.iter():
            tag = obj_node.tag.split('}')[-1] if '}' in obj_node.tag else obj_node.tag
            if tag == 'object':
                nid = obj_node.get('id')
                if nid:
                    printable_ids_set.add(nid)

    surviving_objects = list(resources_el)
    # Sort: internal meshes (not in printable_ids_set) first, then printable, all by int id
    surviving_objects.sort(key=lambda o: (
        1 if o.get('id', '') in printable_ids_set else 0,
        _safe_int(o.get('id', '0'))
    ))

    id_map: dict[str, str] = {}
    new_id_counter = 1
    for obj in surviving_objects:
        old_id = obj.get('id', '')
        new_id = str(new_id_counter)
        id_map[old_id] = new_id
        obj.set('id', new_id)
        # Physically move object to end of resources so order matches sort
        resources_el.remove(obj)
        resources_el.append(obj)
        new_id_counter += 1

    # Remap build item objectids
    for item in root.findall(f'.//{{{NS_CORE}}}build/{{{NS_CORE}}}item'):
        old_id = item.get('objectid', '')
        if old_id in id_map:
            item.set('objectid', id_map[old_id])

    # Remap component objectids
    for comp in root.findall(f'.//{{{NS_CORE}}}components/{{{NS_CORE}}}component'):
        old_id = comp.get('objectid', '')
        if old_id in id_map:
            comp.set('objectid', id_map[old_id])

    # Remap model_settings IDs
    if has_settings and sett_root is not None:
        for node in sett_root.iter():
            tag = node.tag.split('}')[-1] if '}' in node.tag else node.tag
            if tag == 'object':
                old_id = node.get('id', '')
                if old_id in id_map:
                    node.set('id', id_map[old_id])
            elif tag == 'part':
                old_id = node.get('id', '')
                if old_id in id_map:
                    node.set('id', id_map[old_id])
        for asm in sett_root.findall('.//assemble/assemble_item'):
            old_id = asm.get('object_id', '')
            if old_id in id_map:
                asm.set('object_id', id_map[old_id])
        plate_el = sett_root.find('.//plate')
        if plate_el is not None:
            for meta in plate_el.findall('model_instance/metadata[@key="object_id"]'):
                old_id = meta.get('value', '')
                if old_id in id_map:
                    meta.set('value', id_map[old_id])

    # ── Finalize Metadata: update identify_id in plate instances ─────────────
    if has_settings and sett_root is not None:
        plate_el = sett_root.find('.//plate')
        if plate_el is not None:
            for inst in list(plate_el.findall('model_instance')):
                meta_id = inst.find('metadata[@key="object_id"]')
                if meta_id is not None:
                    new_obj_id = meta_id.get('value', '')
                    # Reverse-map new_obj_id -> original_id to find identify_id
                    original_id = None
                    for k, v in id_map.items():
                        if v == new_obj_id:
                            original_id = k
                            break
                    if original_id is not None:
                        if original_id in killed_ids:
                            plate_el.remove(inst)
                        elif original_id in survivor_identify_ids:
                            id_node = inst.find('metadata[@key="identify_id"]')
                            if id_node is not None:
                                id_node.set('value', str(survivor_identify_ids[original_id]))

    # ── cut_information.xml: remap/remove object IDs ──────────────────────────
    if cut_info_path and cut_info_path.exists():
        try:
            _register_ns(cut_info_path)
            cut_tree = ET.parse(cut_info_path)
            cut_root = cut_tree.getroot()
            cut_modified = False
            for co in list(cut_root.iter()):
                tag = co.tag.split('}')[-1] if '}' in co.tag else co.tag
                if tag == 'object':
                    old_id = co.get('id', '')
                    if old_id in id_map:
                        co.set('id', id_map[old_id])
                        cut_modified = True
                    else:
                        for parent in cut_root.iter():
                            if co in list(parent):
                                try:
                                    parent.remove(co)
                                except ValueError:
                                    pass
                                cut_modified = True
                                break
            if cut_modified:
                cut_tree.write(cut_info_path, encoding='utf-8', xml_declaration=True)
        except Exception as e:
            print(f'[merge] WARNING: cut_information.xml update failed: {e}')

    # ── slice_info.config: remap/remove object IDs, sync parts ───────────────
    if slice_info_path and slice_info_path.exists():
        try:
            _register_ns(slice_info_path)
            slice_tree = ET.parse(slice_info_path)
            slice_root = slice_tree.getroot()
            slice_modified = False

            for s_obj in list(slice_root.iter()):
                tag = s_obj.tag.split('}')[-1] if '}' in s_obj.tag else s_obj.tag
                if tag != 'object':
                    continue
                old_id = s_obj.get('id', '')
                if old_id in killed_ids:
                    for parent in slice_root.iter():
                        if s_obj in list(parent):
                            try:
                                parent.remove(s_obj)
                            except ValueError:
                                pass
                            slice_modified = True
                            break
                elif old_id in id_map:
                    new_mapped_id = id_map[old_id]
                    s_obj.set('id', new_mapped_id)
                    slice_modified = True

                    if old_id in survivor_faces:
                        fc_node = s_obj.find('metadata[@face_count]')
                        if fc_node is not None:
                            fc_node.set('face_count', str(survivor_faces[old_id]))
                        name_node = s_obj.find('metadata[@key="name"]')
                        if name_node is not None:
                            name_node.set('value', survivor_names[old_id])

                        # Sync parts from model_settings
                        if has_settings and sett_root is not None:
                            match_sett = None
                            for sn in sett_root.iter():
                                stag = sn.tag.split('}')[-1] if '}' in sn.tag else sn.tag
                                if stag == 'object' and sn.get('id') == new_mapped_id:
                                    match_sett = sn
                                    break
                            if match_sett is not None:
                                for sp in list(s_obj.findall('part')):
                                    s_obj.remove(sp)
                                for mp in match_sett.findall('part'):
                                    s_obj.append(copy.deepcopy(mp))

            # Remap model_instance object_id references
            for inst in list(slice_root.findall('.//model_instance')):
                meta_id = inst.find('metadata[@key="object_id"]')
                if meta_id is not None:
                    old_id = meta_id.get('value', '')
                    if old_id in killed_ids:
                        for parent in slice_root.iter():
                            if inst in list(parent):
                                try:
                                    parent.remove(inst)
                                except ValueError:
                                    pass
                                slice_modified = True
                                break
                    elif old_id in id_map:
                        meta_id.set('value', id_map[old_id])
                        slice_modified = True
                        if old_id in survivor_identify_ids:
                            id_node = inst.find('metadata[@key="identify_id"]')
                            if id_node is not None:
                                id_node.set('value', str(survivor_identify_ids[old_id]))

            if slice_modified:
                slice_tree.write(slice_info_path, encoding='utf-8', xml_declaration=True)
        except Exception as e:
            print(f'[merge] WARNING: slice_info.config update failed: {e}')

    # ── Save XMLs ─────────────────────────────────────────────────────────────
    tree.write(model_file, encoding='utf-8', xml_declaration=True)
    if has_settings and sett_tree is not None and settings_path is not None:
        sett_tree.write(settings_path, encoding='utf-8', xml_declaration=True)

    # ── VLH: rewrite with renumbered IDs (or delete if no data) ──────────────
    if has_settings and vlh_data_string is not None and vlh_path.exists():
        # Collect printable IDs in numeric order (after renumbering)
        printable_new_ids: list[int] = []
        if sett_root is not None:
            for node in sett_root.iter():
                tag = node.tag.split('}')[-1] if '}' in node.tag else node.tag
                if tag == 'object':
                    nid = node.get('id')
                    if nid:
                        try:
                            printable_new_ids.append(int(nid))
                        except ValueError:
                            pass
        printable_new_ids.sort()
        new_vlh_lines = [f'object_id={fid}|{vlh_data_string}' for fid in printable_new_ids]
        vlh_path.write_text('\r\n'.join(new_vlh_lines), encoding='utf-8')
    elif vlh_path.exists():
        vlh_path.unlink(missing_ok=True)

    # ── Final sweep: protect any file still referenced anywhere in XML ────────
    for node in root.iter():
        for attr_name, attr_val in node.attrib.items():
            local = attr_name.split('}')[-1] if '}' in attr_name else attr_name
            if local == 'path' and attr_val:
                rel_path = attr_val if attr_val.startswith('/') else '/' + attr_val
                used_model_paths.add(rel_path)

    # ── Garbage-collect unused .model files ───────────────────────────────────
    if objects_dir.exists():
        for f in objects_dir.glob('*.model'):
            check = f'/3D/Objects/{f.name}'
            if check not in used_model_paths:
                f.unlink(missing_ok=True)

    # ── Rebuild .rels ─────────────────────────────────────────────────────────
    if rels_path:
        rels_lines = ['<?xml version="1.0" encoding="UTF-8"?>',
                      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        for i, p in enumerate(sorted(used_model_paths), start=1):
            rels_lines.append(
                f'  <Relationship Target="{p}" Id="rel-{i}" '
                f'Type="http://schemas.microsoft.com/3dmanufacturing/2013/01/3dmodel"/>'
            )
        rels_lines.append('</Relationships>')
        rels_path.write_text('\r\n'.join(rels_lines), encoding='utf-8')

    # ── Purge stale UI cache (forces Bambu to regenerate thumbnails) ──────────
    meta_dir = work_dir / 'Metadata'
    if meta_dir.exists():
        for pat in ('pick_*.png', 'plate_*.png', 'plate_*.json'):
            for stale in meta_dir.glob(pat):
                stale.unlink(missing_ok=True)

    # ── Repack to output ──────────────────────────────────────────────────────
    # [Content_Types].xml must be first per OPC spec; .rels stored uncompressed.
    all_files = sorted(
        (f for f in work_dir.rglob('*') if f.is_file()),
        key=lambda f: (0 if f.name == '[Content_Types].xml' else
                       1 if f.suffix == '.rels' else 2,
                       str(f)),
    )
    if output_path.exists():
        output_path.unlink()
    with zipfile.ZipFile(output_path, 'w') as zf:
        for f in all_files:
            rel = f.relative_to(work_dir).as_posix()
            compress = (zipfile.ZIP_STORED
                        if f.name == '[Content_Types].xml' or f.suffix == '.rels'
                        else zipfile.ZIP_DEFLATED)
            zf.write(f, rel, compress_type=compress)

    _log(f'[merge] Done -> {output_path.name}')
    _log(f'[merge] surviving objects (renumbered): {new_id_counter - 1}')
    _log(f'[merge] id_map: {id_map}')

    # ── Debug: dump model XML header and ZIP contents ─────────────────────────
    _log('\n[merge][DEBUG] 3dmodel.model first 1200 chars:')
    try:
        _log(model_file.read_text(encoding='utf-8', errors='replace')[:1200])
    except Exception as e:
        _log(f'  (could not read: {e})')
    _log('\n[merge][DEBUG] Output ZIP entries:')
    try:
        with zipfile.ZipFile(output_path, 'r') as _dbgz:
            for _info in _dbgz.infolist():
                _log(f'  {_info.compress_type:4d}  {_info.file_size:8d}  {_info.filename}')
    except Exception as e:
        _log(f'  (could not open: {e})')

    _dbg.close()
    return 0


def _safe_int(s: str) -> int:
    try:
        return int(s)
    except (ValueError, TypeError):
        return 0


def main() -> int:
    parser = argparse.ArgumentParser(description='N-way 3MF merge worker')
    parser.add_argument('--work-dir',    required=True)
    parser.add_argument('--input-path',  required=True)
    parser.add_argument('--output-path', required=True)
    parser.add_argument('--do-colors',   default='0')
    args = parser.parse_args()
    return merge(
        Path(args.work_dir),
        Path(args.input_path),
        Path(args.output_path),
        do_colors=args.do_colors != '0',
    )


if __name__ == '__main__':
    sys.exit(main())
