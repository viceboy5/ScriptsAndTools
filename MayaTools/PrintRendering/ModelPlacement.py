import maya.cmds as cmds
import math

# ----------------------------------------
# Config
# ----------------------------------------
HARDEN_ANGLE_THRESHOLD = 30.0  # degrees


# ----------------------------------------
# Normal Utilities
# ----------------------------------------
def _angle_between(v1, v2):
    dot = sum(a * b for a, b in zip(v1, v2))
    mag1 = math.sqrt(sum(a * a for a in v1))
    mag2 = math.sqrt(sum(a * a for a in v2))
    if mag1 == 0 or mag2 == 0:
        return 0.0
    cos_angle = max(min(dot / (mag1 * mag2), 1.0), -1.0)
    return math.degrees(math.acos(cos_angle))


def harden_edges_by_angle(mesh, angle_threshold=30.0):
    """Hardens only edges whose face angle >= threshold."""
    hard_edges = []

    face_count = cmds.polyEvaluate(mesh, face=True)
    face_normals = {}

    for i in range(face_count):
        info = cmds.polyInfo(f"{mesh}.f[{i}]", faceNormals=True)[0]
        face_normals[i] = [float(n) for n in info.split()[2:5]]

    edge_info = cmds.polyInfo(mesh, edgeToFace=True) or []

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
            hard_edges.append(f"{mesh}.e[{edge_id}]")

    if hard_edges:
        cmds.select(hard_edges, r=True)
        cmds.polySoftEdge(angle=0, constructionHistory=False)


# ----------------------------------------
# Pivot / Merge / Separate
# ----------------------------------------
def merge_separate_with_pivot_to_origin(rotate_axes=None, rotation_degrees=-90.0):
    """
    rotate_axes:      list of axis chars to rotate, e.g. ['x'], ['x', 'z'], []
    rotation_degrees: degrees applied to each checked axis
    """
    if rotate_axes is None:
        rotate_axes = []

    sel = cmds.ls(selection=True, long=True)
    if not sel:
        cmds.warning("Select a mesh transform or mesh shape first.")
        return

    transform = sel[0]
    if cmds.objectType(transform) == "mesh":
        parents = cmds.listRelatives(transform, parent=True, fullPath=True)
        if not parents:
            cmds.warning("Selected mesh shape has no transform parent.")
            return
        transform = parents[0]

    shapes = cmds.listRelatives(transform, shapes=True, fullPath=True) or []
    if not any(cmds.objectType(s) == "mesh" for s in shapes):
        cmds.warning("Selection is not a mesh transform.")
        return

    # 1) Rotate on selected axes
    rot_x = rotation_degrees if 'x' in rotate_axes else 0
    rot_y = rotation_degrees if 'y' in rotate_axes else 0
    rot_z = rotation_degrees if 'z' in rotate_axes else 0
    if rot_x or rot_y or rot_z:
        cmds.xform(transform, relative=True, rotation=(rot_x, rot_y, rot_z), ws=True)

    # 2) Center pivot
    cmds.xform(transform, cp=True)

    # 3) Pivot to bbox min Y (world space floor)
    bbox = cmds.exactWorldBoundingBox(transform)
    minX, minY, minZ, maxX, maxY, maxZ = bbox
    cur_pivot = cmds.xform(transform, q=True, ws=True, rp=True)
    new_pivot = [cur_pivot[0], minY, cur_pivot[2]]
    cmds.xform(transform, ws=True, rp=new_pivot, sp=new_pivot)

    # 4) Move object so pivot is at origin
    pivot_world = cmds.xform(transform, q=True, ws=True, rp=True)
    offset = [-pivot_world[0], -pivot_world[1], -pivot_world[2]]
    cmds.xform(transform, ws=True, t=offset, r=False)

    # 4) Merge vertices
    verts = cmds.polyListComponentConversion(transform, toVertex=True)
    verts = cmds.filterExpand(verts, sm=31)
    if not verts:
        cmds.warning("No vertices found.")
        return
    cmds.select(verts, r=True)
    cmds.polyMergeVertex(distance=0.0001)

    # 5) Reset normals, then re-harden sharp edges
    cmds.polySoftEdge(transform, angle=180, constructionHistory=False)
    harden_edges_by_angle(transform, HARDEN_ANGLE_THRESHOLD)

    # 6) Unlock normals before separating
    cmds.polyNormalPerVertex(transform, unFreezeNormal=True)

    # 7) Separate shells
    separated = cmds.polySeparate(transform, constructionHistory=False)

    # 8) Center pivot and apply Arnold subdivision on each piece
    if separated:
        for piece in separated:
            cmds.xform(piece, cp=True)
            shapes = cmds.listRelatives(piece, shapes=True, fullPath=True) or []
            for shape in shapes:
                if cmds.objectType(shape) == "mesh":
                    cmds.setAttr("{}.aiSubdivType".format(shape), 1)       # catclark
                    cmds.setAttr("{}.aiSubdivIterations".format(shape), 2)
        cmds.select(separated, r=True)

    print("Finished Pivot/Merge: rotate_axes={}, degrees={}".format(rotate_axes, rotation_degrees))


# ----------------------------------------
# Cluster
# ----------------------------------------
def create_cluster():
    sel = cmds.ls(selection=True)
    if not sel:
        cmds.warning("Select objects or components to create a cluster.")
        return
    result = cmds.cluster(sel)
    print("Created cluster: {}".format(result))


# ----------------------------------------
# Cluster Rotate Left / Right
# ----------------------------------------
def rotate_left_group():
    sel = cmds.ls(selection=True, long=True)
    if not sel or len(sel) < 2:
        cmds.warning("Select a cluster handle and objects for Rotate Left.")
        return

    cluster_handle = sel[0]
    objects = sel[1:]

    cluster_pos = cmds.xform(cluster_handle, q=True, ws=True, rp=True)
    grp = cmds.group(empty=True, name="RotateLeft_Group")
    cmds.xform(grp, ws=True, t=cluster_pos)

    for obj in objects:
        cmds.parent(obj, grp)

    cmds.xform(grp, ws=True, rotation=(0, 0, 90))
    cmds.delete(cluster_handle)

    print("Rotate Left: Objects grouped under '{}'.".format(grp))


def rotate_right_group():
    sel = cmds.ls(selection=True, long=True)
    if not sel:
        cmds.warning("Select objects for Rotate Right.")
        return

    objects = sel

    if not cmds.objExists("RotateLeft_Group"):
        cmds.warning("No RotateLeft_Group found. Run Rotate Left first.")
        return

    left_pos = cmds.xform("RotateLeft_Group", q=True, ws=True, rp=True)
    mirrored_pos = [-left_pos[0], left_pos[1], left_pos[2]]

    grp = cmds.group(empty=True, name="RotateRight_Group")
    cmds.xform(grp, ws=True, t=mirrored_pos)

    for obj in objects:
        cmds.parent(obj, grp)

    cmds.xform(grp, ws=True, rotation=(0, 0, -90))

    print("Rotate Right: Objects grouped under '{}'.".format(grp))


# ----------------------------------------
# Model Prep UI
# ----------------------------------------
def show_model_prep_ui():
    window = "modelPrepUI"
    if cmds.window(window, exists=True):
        cmds.deleteUI(window)

    cmds.window(window, title="Model Prep", sizeable=True, widthHeight=(300, 320))
    main_col = cmds.columnLayout(adj=True, rowSpacing=15)

    # Pivot / Merge Section
    cmds.frameLayout(label="Pivot / Merge / Separate", collapsable=True, parent=main_col)
    cmds.columnLayout(adj=True, rowSpacing=5)

    cmds.button(
        label="Maintain Orientation",
        height=30,
        command=lambda *_: merge_separate_with_pivot_to_origin()
    )

    cmds.separator(height=8, style='in')
    cmds.text(label="Rotation Axes:", align='left')

    row = cmds.rowLayout(numberOfColumns=3, columnWidth3=(80, 80, 80))
    cb_x = cmds.checkBox(label="X", value=True)
    cb_y = cmds.checkBox(label="Y", value=False)
    cb_z = cmds.checkBox(label="Z", value=False)
    cmds.setParent('..')

    cmds.rowLayout(numberOfColumns=2, columnWidth2=(80, 100))
    cmds.text(label="Degrees:")
    deg_field = cmds.floatField(value=-90.0, precision=1)
    cmds.setParent('..')

    def _run_rotate(*_):
        axes = []
        if cmds.checkBox(cb_x, q=True, value=True):
            axes.append('x')
        if cmds.checkBox(cb_y, q=True, value=True):
            axes.append('y')
        if cmds.checkBox(cb_z, q=True, value=True):
            axes.append('z')
        deg = cmds.floatField(deg_field, q=True, value=True)
        merge_separate_with_pivot_to_origin(rotate_axes=axes, rotation_degrees=deg)

    cmds.button(label="Rotate", height=30, command=_run_rotate)
    cmds.setParent("..")

    # Cluster Rotate Section
    cmds.frameLayout(label="Cluster Rotate", collapsable=True, parent=main_col)
    cmds.columnLayout(adj=True, rowSpacing=5)
    cmds.button(label="Create Cluster", height=30, command=lambda *_: create_cluster())
    cmds.separator(height=8, style='in')
    cmds.button(label="Rotate Left (90° Z)", height=30, command=lambda *_: rotate_left_group())
    cmds.button(label="Rotate Right (-90° Z)", height=30, command=lambda *_: rotate_right_group())
    cmds.setParent("..")

    cmds.showWindow(window)


# ----------------------------------------
# Shelf Button Call
# ----------------------------------------
show_model_prep_ui()
