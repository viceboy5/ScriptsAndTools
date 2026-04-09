#!/usr/bin/env python3
"""
isolate_worker.py - Isolate the center object from a merged 3MF's extracted directory.

Usage:
    python isolate_worker.py --work-dir <extracted_3mf_dir> --output-path <Final.3mf>

Mirrors PS1 isolate_final_worker.ps1.
"""
import argparse
import math
import os
import re
import sys
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


# ── Transform helpers ─────────────────────────────────────────────────────────

def parse_tx(s: str) -> list[float]:
    if not s or not s.strip():
        return [1,0,0, 0,1,0, 0,0,1, 0,0,0]
    return [float(v) for v in s.strip().split()]


# ── XML helpers ───────────────────────────────────────────────────────────────

def _find_file(base: Path, rel: str) -> Path | None:
    p = base / rel.replace('\\', '/').lstrip('/')
    if p.exists():
        return p
    p2 = base / rel
    if p2.exists():
        return p2
    return None


def _save_xml(tree: ET.ElementTree, path: Path) -> None:
    tree.write(path, encoding='utf-8', xml_declaration=True)


def _find_all(root: ET.Element, tag: str) -> list[ET.Element]:
    return root.findall(f'.//{{{NS_CORE}}}{tag}') + root.findall(f'.//{tag}')


# ── Main logic ────────────────────────────────────────────────────────────────

def isolate(work_dir: Path, output_path: Path) -> int:
    """
    Finds the majority face-count object closest to (128, 128) and strips
    everything else, then repacks to output_path.
    Returns 0 on success, 1 on failure.
    """
    # Locate files
    model_files = list(work_dir.rglob('3dmodel.model'))
    if not model_files:
        print('[isolate] ERROR: 3dmodel.model not found')
        return 1
    model_file = model_files[0]
    objects_dir = work_dir / '3D' / 'Objects'
    rels_path = _find_file(work_dir, '3D/_rels/3dmodel.model.rels')
    settings_path = _find_file(work_dir, 'Metadata/model_settings.config')
    cut_info_path = _find_file(work_dir, 'Metadata/cut_information.xml')

    # Parse main model
    ET.register_namespace('', NS_CORE)
    ET.register_namespace('p', NS_PROD)
    _register_ns(model_file)
    tree = ET.parse(model_file)
    root = tree.getroot()

    ns = {'m': NS_CORE, 'p': NS_PROD}
    build_el = root.find(f'{{{NS_CORE}}}build') or root.find('build')
    if build_el is None:
        print('[isolate] ERROR: <build> not found in model')
        return 1

    build_items = list(build_el)
    obj_by_id: dict[str, ET.Element] = {}
    resources_el = root.find(f'{{{NS_CORE}}}resources') or root.find('resources')
    if resources_el is not None:
        for obj in resources_el:
            oid = obj.get('id')
            if oid:
                obj_by_id[oid] = obj

    # Parse model_settings.config
    has_settings = settings_path and settings_path.exists()
    sett_obj_by_id: dict[str, ET.Element] = {}
    sett_root: ET.Element | None = None
    sett_tree: ET.ElementTree | None = None
    if has_settings:
        try:
            sett_tree = ET.parse(settings_path)
            sett_root = sett_tree.getroot()
            for node in sett_root.iter():
                if node.tag in ('object', f'{{{NS_CORE}}}object') or node.get('id'):
                    nid = node.get('id')
                    if nid:
                        sett_obj_by_id[nid] = node
        except Exception:
            has_settings = False

    # ── 1. Find majority face count ──────────────────────────────────────────
    fc_map: dict[str, int] = {}
    for item in build_items:
        oid = item.get('objectid', '')
        fc = 'unknown'
        if has_settings and oid in sett_obj_by_id:
            fc_node = sett_obj_by_id[oid].find('.//metadata[@key="face_count"]')
            if fc_node is not None:
                fc = fc_node.get('value', 'unknown')
        fc_map[fc] = fc_map.get(fc, 0) + 1

    majority_fc = max(fc_map, key=lambda k: fc_map[k]) if fc_map else 'unknown'

    # ── 2. Filter target items ───────────────────────────────────────────────
    target_items: list[ET.Element] = []
    for item in build_items:
        oid = item.get('objectid', '')
        is_target = True
        if has_settings and oid in sett_obj_by_id:
            fc_node = sett_obj_by_id[oid].find('.//metadata[@key="face_count"]')
            fc = fc_node.get('value', 'unknown') if fc_node is not None else 'unknown'
            if fc != majority_fc:
                is_target = False
            name_node = sett_obj_by_id[oid].find('.//metadata[@key="name"]')
            if name_node is not None:
                val = name_node.get('value', '')
                if 'text' in val.lower() or 'version' in val.lower():
                    is_target = False
        if is_target:
            target_items.append(item)

    if not target_items:
        print('[isolate] ERROR: No target items found')
        return 1

    # ── 3. Find closest target to (128, 128) ─────────────────────────────────
    closest_item: ET.Element | None = None
    min_dist = float('inf')
    for item in target_items:
        tx = parse_tx(item.get('transform', ''))
        dist = (tx[9] - 128) ** 2 + (tx[10] - 128) ** 2
        if dist < min_dist:
            min_dist = dist
            closest_item = item

    if closest_item is None:
        return 1

    survivor_id = closest_item.get('objectid', '')

    # ── 4. Eradicate everything else ─────────────────────────────────────────
    killed_ids: set[str] = set()
    for item in list(build_items):
        oid = item.get('objectid', '')
        if item is not closest_item:
            killed_ids.add(oid)
            build_el.remove(item)
            if oid in obj_by_id and resources_el is not None:
                try:
                    resources_el.remove(obj_by_id[oid])
                except ValueError:
                    pass

    # ── 5. Clean metadata ─────────────────────────────────────────────────────
    if has_settings and sett_root is not None:
        # Clean plate instances
        plate_el = sett_root.find('.//plate')
        if plate_el is not None:
            found_survivor = False
            for inst in list(plate_el):
                meta_id = inst.find('.//metadata[@key="object_id"]')
                if meta_id is not None:
                    if meta_id.get('value') == survivor_id and not found_survivor:
                        found_survivor = True
                    else:
                        plate_el.remove(inst)
                else:
                    plate_el.remove(inst)

        # Remove orphaned object blocks
        for kid in killed_ids:
            if kid in sett_obj_by_id:
                parent = None
                for el in sett_root.iter():
                    if sett_obj_by_id[kid] in list(el):
                        parent = el
                        break
                if parent is not None:
                    try:
                        parent.remove(sett_obj_by_id[kid])
                    except ValueError:
                        pass

        # Clean assemble trackers
        assemble = sett_root.find('.//assemble')
        if assemble is not None:
            for kid in killed_ids:
                for item in list(assemble):
                    if item.get('object_id') == kid:
                        assemble.remove(item)
            # Fix survivor assemble transform
            surv_asm = assemble.find(f'.//assemble_item[@object_id="{survivor_id}"]')
            if surv_asm is not None and closest_item is not None:
                tx = parse_tx(closest_item.get('transform', ''))
                surv_asm.set('transform', f'1 0 0 0 1 0 0 0 1 {tx[9]} {tx[10]} {tx[11]}')

    # ── 6. Clean cut_information.xml ─────────────────────────────────────────
    if cut_info_path and cut_info_path.exists():
        try:
            cut_tree = ET.parse(cut_info_path)
            cut_root = cut_tree.getroot()
            for obj in list(cut_root.iter()):
                if obj.get('id') in killed_ids:
                    for parent in cut_root.iter():
                        if obj in list(parent):
                            parent.remove(obj)
                            break
            cut_tree.write(cut_info_path, encoding='utf-8', xml_declaration=True)
        except Exception:
            pass

    # ── 7. Clean unused .model files, rebuild .rels ───────────────────────────
    preserved_paths: set[str] = set()
    if survivor_id in obj_by_id:
        surv_obj = obj_by_id[survivor_id]
        for comp in surv_obj.iter():
            path = comp.get(f'{{{NS_PROD}}}path') or comp.get('p:path') or comp.get('path', '')
            if path:
                preserved_paths.add(path)

    if objects_dir.exists():
        for f in objects_dir.iterdir():
            if f.suffix == '.model':
                check = f'/3D/Objects/{f.name}'
                if check not in preserved_paths:
                    f.unlink(missing_ok=True)

    if rels_path and rels_path.exists():
        rels_xml = (
            "<?xml version='1.0' encoding='UTF-8'?>"
            "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
        )
        for i, p in enumerate(preserved_paths, start=1):
            rels_xml += (
                f"<Relationship Target='{p}' Id='rel-ign-{i}' "
                f"Type='http://schemas.microsoft.com/3dmanufacturing/2013/01/3dmodel'/>"
            )
        rels_xml += '</Relationships>'
        rels_path.write_text(rels_xml, encoding='utf-8')

    # Save main model
    tree.write(model_file, encoding='utf-8', xml_declaration=True)
    if has_settings and sett_tree is not None:
        sett_tree.write(settings_path, encoding='utf-8', xml_declaration=True)

    # ── 8. Repack to output zip ───────────────────────────────────────────────
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

    print(f'[isolate] Done -> {output_path.name}')
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description='Isolate center object from merged 3MF')
    parser.add_argument('--work-dir',     required=True, help='Extracted 3MF directory')
    parser.add_argument('--output-path',  required=True, help='Output Final.3mf path')
    args = parser.parse_args()
    return isolate(Path(args.work_dir), Path(args.output_path))


if __name__ == '__main__':
    sys.exit(main())
