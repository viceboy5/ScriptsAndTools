import maya.cmds as cmds
import math

ANGLE_THRESHOLD = 20.0  # degrees


def _angle_between(v1, v2):
    dot = sum(a * b for a, b in zip(v1, v2))
    mag1 = math.sqrt(sum(a * a for a in v1))
    mag2 = math.sqrt(sum(a * a for a in v2))
    if mag1 == 0 or mag2 == 0:
        return 0.0
    cos_angle = max(min(dot / (mag1 * mag2), 1.0), -1.0)
    return math.degrees(math.acos(cos_angle))


def soften_edges_by_angle(selection=None, angle_threshold=30.0):
    if selection is None:
        selection = cmds.ls(selection=True, long=True)
    if not selection:
        cmds.warning("Select one or more mesh transforms.")
        return

    for obj in selection:
        shapes = cmds.listRelatives(obj, shapes=True, fullPath=True) or []
        for shape in shapes:
            if cmds.nodeType(shape) != "mesh":
                continue

            # 1) Soften everything first
            cmds.polySoftEdge(shape, angle=180, constructionHistory=False)

            # 2) Cache face normals
            face_count = cmds.polyEvaluate(shape, face=True)
            face_normals = {}
            for i in range(face_count):
                info = cmds.polyInfo(f"{shape}.f[{i}]", faceNormals=True)[0]
                face_normals[i] = [float(n) for n in info.split()[2:5]]

            # 3) Evaluate edges
            edge_info = cmds.polyInfo(shape, edgeToFace=True) or []
            hard_edges = []

            for info in edge_info:
                parts = info.split()
                edge_id = int(parts[1].replace(":", ""))
                face_ids = [int(p) for p in parts[2:] if p.isdigit()]
                if len(face_ids) != 2:
                    continue
                n1 = face_normals[face_ids[0]]
                n2 = face_normals[face_ids[1]]
                angle = _angle_between(n1, n2)
                if angle >= angle_threshold:
                    hard_edges.append(f"{shape}.e[{edge_id}]")

            # 4) Harden sharp edges only
            if hard_edges:
                cmds.select(hard_edges, r=True)
                cmds.polySoftEdge(angle=0, constructionHistory=False)

    cmds.select(selection, r=True)
    print(f"Done: edges = {angle_threshold}° hardened.")


# Run it on current selection
soften_edges_by_angle(selection=None, angle_threshold=ANGLE_THRESHOLD)
