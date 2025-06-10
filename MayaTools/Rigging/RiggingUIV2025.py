import maya.cmds as cmds
from PySide6 import QtWidgets, QtCore
import maya.OpenMayaUI as omui
import shiboken6
import re
import sys
script_folder = r"D:\GitRepos\ScriptsAndTools\MayaTools\Rigging"
if script_folder not in sys.path:
    sys.path.append(script_folder)
import CreateRK

def maya_main_window():
    """
    Get the main Maya window as a QMainWindow instance.
    """
    main_window_ptr = omui.MQtUtil.mainWindow()
    return shiboken6.wrapInstance(int(main_window_ptr), QtWidgets.QWidget)

def create_group_and_ctrl(joint, color_index):
    # Create an empty group
    empty_group = cmds.group(empty=True, name=f"{joint}_Ctrl_Grp")

    # Match the group's transformations to the joint
    cmds.matchTransform(empty_group, joint, pos=True, rot=True, scale=True)

    # Create a NURBS circle
    circle = cmds.circle(name=f"{joint}_Ctrl")[0]

    # Parent the circle under the empty group
    cmds.parent(circle, empty_group)

    # Match the circle's transformations to the joint
    cmds.matchTransform(circle, joint, pos=True, rot=True, scale=True)

    # Rotate the circle 90 degrees on the Z-axis
    cmds.rotate(0, 90, 0, circle, relative=True)

    # Freeze the circle's rotations
    cmds.makeIdentity(circle, apply=True, rotate=True)

    # Enable drawing overrides for the circle
    cmds.setAttr(f"{circle}.overrideEnabled", 1)

    # Set the drawing override color
    cmds.setAttr(f"{circle}.overrideColor", color_index)

    return empty_group, circle

def create_ctrls_for_joint_chain(top_joint, color_index):
    # Create an empty group and control for the top joint
    top_empty_group, top_circle = create_group_and_ctrl(top_joint, color_index)

    # Get the joint chain excluding the top joint
    joint_chain = cmds.listRelatives(top_joint, allDescendents=True, type="joint") or []

    # Iterate through the joint chain
    for joint in joint_chain:
        # Create an empty group and control circle
        create_group_and_ctrl(joint, color_index)

def create_controls_clicked(color_name):
    color_map = {"Blue": 6, "Red": 13, "Green": 14}
    color_index = color_map.get(color_name, 6)  # Default to blue if color not found
    # Get the selected joint
    selected_joint = cmds.ls(selection=True, type="joint")

    if not selected_joint:
        cmds.warning("Please select a joint to start from.")
    else:
        # Call the function to create groups and control circles for the joint chain
        create_ctrls_for_joint_chain(selected_joint[0], color_index)

def create_joint_at_cluster_transforms():
    # Get all cluster handles in the scene
    cluster_handles = cmds.ls(type='clusterHandle')

    if not cluster_handles:
        cmds.warning("No cluster handles found in the scene.")
        return

    joint_names = []  # List to keep track of created joints

    for cluster_handle in cluster_handles:
        # Get the transform node of the cluster handle
        transform_node = cmds.listRelatives(cluster_handle, parent=True, type='transform')

        if not transform_node:
            cmds.warning(f"No transform node found for cluster handle {cluster_handle}")
            continue

        # Get the world position of the transform node
        cluster_position = cmds.xform(transform_node[0], query=True, translation=True, worldSpace=True)

        # Create a joint at the transform node's world position
        joint_name = cmds.joint(position=cluster_position)
        joint_names.append(joint_name)  # Store the joint name
        print(f"Created joint {joint_name} at {cluster_position} for transform node {transform_node[0]}")

    # Unparent each joint after creation
    for joint in joint_names:
        cmds.parent(joint, world=True)  # Unparent the joint to world
        print(f"Unparented joint {joint}")

def RenameSequentially(txt):
    count = txt.count("#")

    if count == 0:
        print("# should be used to indicate the numbering in between the Name and NodeType.")
        return

    nums_placeholder = "#" * count
    x = txt.find(nums_placeholder)

    if x == -1:
        print("# should only be used in between the Name and NodeType. Ensure all arguments are named appropriately.")
    else:
        parts = txt.partition(nums_placeholder)

        sels = cmds.ls(sl=True)
        for i, sel in enumerate(sels):
            newNum = str(i + 1).zfill(count)  # Correctly format the numbering
            new_name = parts[0] + newNum + parts[2]
            cmds.rename(sel, new_name)

def find_matching_joint(ctrl_name):
    # Extract joint name from the control name
    joint_name = ctrl_name.replace("_Ctrl", "")

    # Check if the joint exists
    if cmds.objExists(joint_name):
        return joint_name
    else:
        return None

def create_constraints(ctrl_name, joint_name):
    # Create parent constraint
    parent_constraint = cmds.parentConstraint(ctrl_name, joint_name, mo=True)[0]

    # Create scale constraint
    scale_constraint = cmds.scaleConstraint(ctrl_name, joint_name, mo=True)[0]

    return parent_constraint, scale_constraint

def find_and_create_constraints():
    # List all transform nodes in the scene
    transforms = cmds.ls(type='transform')

    for ctrl in transforms:
        # Check if the control name ends with '_Ctrl'
        if ctrl.endswith('_Ctrl'):
            # Find matching joint
            joint_name = find_matching_joint(ctrl)

            if joint_name:
                # Create constraints
                create_constraints(ctrl, joint_name)
            else:
                print(f"Joint not found for control: {ctrl}")

def create_broken_fk_constraints():
    # Get selection, separate parent control and child control
    sels = cmds.ls(sl=True)  # [parent control, child control]
    if len(sels) < 2:
        cmds.warning("Select a parent control and a child control.")
        return
    parent_ctrl = sels[0]
    child_ctrl = sels[1]

    # Get the parent group of the child control
    child_ctrl_grp = cmds.listRelatives(child_ctrl, parent=True)[0]  # [child control's parent node]

    # Create constraints
    p_constraint1 = cmds.parentConstraint(parent_ctrl, child_ctrl_grp, mo=True, skipRotate=['x', 'y', 'z'], weight=1)[0]  # constrain translates
    p_constraint2 = cmds.parentConstraint(parent_ctrl, child_ctrl_grp, mo=True, skipTranslate=['x', 'y', 'z'], weight=1)[0]  # constrain rotates
    cmds.scaleConstraint(parent_ctrl, child_ctrl_grp, weight=1)

    # Create attributes on the child control
    if not cmds.attributeQuery('FollowTranslate', node=child_ctrl, exists=True):
        cmds.addAttr(child_ctrl, ln='FollowTranslate', at='double', min=0, max=1, dv=1)
        cmds.setAttr('%s.FollowTranslate' % (child_ctrl), e=True, keyable=True)
    if not cmds.attributeQuery('FollowRotate', node=child_ctrl, exists=True):
        cmds.addAttr(child_ctrl, ln='FollowRotate', at='double', min=0, max=1, dv=1)
        cmds.setAttr('%s.FollowRotate' % (child_ctrl), e=True, keyable=True)

    # Connect attributes from child control to constraint weights
    cmds.connectAttr('%s.FollowTranslate' % child_ctrl, '%s.w0' % (p_constraint1), f=True)
    cmds.connectAttr('%s.FollowRotate' % child_ctrl, '%s.w0' % (p_constraint2), f=True)




def create_ui():
    # Create a window
    window = QtWidgets.QWidget(parent=maya_main_window())
    window.setWindowTitle("Rigging Automation")
    window.setGeometry(200, 200, 400, 400)

    # Set window flags for resizing and moving
    window.setWindowFlags(QtCore.Qt.Window | QtCore.Qt.WindowStaysOnTopHint)

    # Create a layout
    layout = QtWidgets.QVBoxLayout()

    # Button for creating controls
    color_dropdown = QtWidgets.QComboBox()
    color_dropdown.addItem("Blue")
    color_dropdown.addItem("Red")
    color_dropdown.addItem("Green")
    layout.addWidget(color_dropdown)

    create_ctrls_button = QtWidgets.QPushButton("Create Controls")
    create_ctrls_button.clicked.connect(lambda: create_controls_clicked(color_dropdown.currentText()))
    layout.addWidget(create_ctrls_button)

    # Button for clusters to joints
    clusters_to_joints_button = QtWidgets.QPushButton("Clusters to Joints")
    clusters_to_joints_button.clicked.connect(create_joint_at_cluster_transforms)
    layout.addWidget(clusters_to_joints_button)

    # Sequential renamer input
    rename_input = QtWidgets.QLineEdit()
    rename_input.setPlaceholderText("Enter name with # for numbering (e.g. FK_Jnt_##)")
    layout.addWidget(rename_input)

    # Button for sequential renamer
    sequential_renamer_button = QtWidgets.QPushButton("Sequential Renamer")
    sequential_renamer_button.clicked.connect(lambda: RenameSequentially(rename_input.text()))
    layout.addWidget(sequential_renamer_button)

    # Button for creating constraints
    create_constraints_button = QtWidgets.QPushButton("Create Constraints for Matching Joints")
    create_constraints_button.clicked.connect(find_and_create_constraints)
    layout.addWidget(create_constraints_button)

    # Button for broken FK constraints
    broken_fk_constraints_button = QtWidgets.QPushButton("BrokenFK Constraints")
    broken_fk_constraints_button.clicked.connect(create_broken_fk_constraints)
    layout.addWidget(broken_fk_constraints_button)

    # Button for creating RK
    create_rk_button = QtWidgets.QPushButton("Create RK")
    create_rk_button.clicked.connect(lambda: CreateRK.build_rk_system())
    layout.addWidget(create_rk_button)

    # Set the layout on the window
    window.setLayout(layout)

    # Show the window
    window.show()

# Run the UI
create_ui()