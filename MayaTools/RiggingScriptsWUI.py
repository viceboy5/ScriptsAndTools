import maya.cmds as cmds
from PySide2 import QtWidgets, QtCore
import maya.OpenMayaUI as omui
import shiboken2
import re

def maya_main_window():
    """
    Get the main Maya window as a QMainWindow instance.
    """
    main_window_ptr = omui.MQtUtil.mainWindow()
    return shiboken2.wrapInstance(int(main_window_ptr), QtWidgets.QWidget)

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

# Main function and helper functions from CreateRKFromFKWithConstraints.py
def duplicate_joint_chains():
    # Get the selected joint (root of the hierarchy)
    selected_joint = cmds.ls(selection=True, type='joint')

    if not selected_joint:
        cmds.warning("Please select a joint.")
        return

    root_joint = selected_joint[0]

    # Get the original hierarchy (including the root joint)
    original_joints = [root_joint] + (cmds.listRelatives(root_joint, allDescendents=True, type='joint') or [])

    # Store the names of the original joints in a list (in hierarchical order)
    original_names = [joint.split('|')[-1] for joint in original_joints]

    # Function to duplicate and rename joints
    def duplicate_and_rename_chain(suffix):
        # List to store duplicated joints
        duplicated_joints = []

        # Duplicate each joint individually and store it
        for original_joint in original_joints:
            # Duplicate the current joint without hierarchy
            duplicated_joint = cmds.duplicate(original_joint, parentOnly=True)[0]
            duplicated_joints.append(duplicated_joint)

        # Create a dictionary to store original joint and its corresponding duplicated joint
        joint_mapping = {original: duplicated for original, duplicated in zip(original_joints, duplicated_joints)}

        # Rebuild the hierarchy based on numeric suffix order
        for original_joint in original_joints:
            # Extract the numeric suffix from the original joint name
            match = re.search(r'(\d+)$', original_joint)
            if match:
                suffix_number = int(match.group(1))
                # Check if the original joint has a parent
                parent_joint = cmds.listRelatives(original_joint, parent=True)
                if parent_joint:
                    parent_joint = joint_mapping[parent_joint[0]]  # Get the corresponding duplicated parent joint
                    # Reparent the duplicated joint to its duplicated parent joint
                    cmds.parent(joint_mapping[original_joint], parent_joint)

        # Rename the duplicated joints according to the original names, replacing 'FK' with the provided suffix
        renamed_joints = []
        for original_name, duplicated_joint in zip(original_names, duplicated_joints):
            # Replace 'FK' with the given suffix while keeping the numeric suffix
            new_name = re.sub(r'FK(.*?)(\d+)', rf'{suffix}\1\2', original_name)

            # Rename the duplicated joint
            renamed_joint = cmds.rename(duplicated_joint, new_name)
            renamed_joints.append(renamed_joint)

        # Sort renamed joints based on their numeric suffix
        renamed_joints.sort(key=lambda j: int(re.search(r'(\d+)$', j).group(1)))  # Sort by suffix number

        return renamed_joints  # Return the renamed joints for constraint creation

    # Create IK chain and get the renamed joints
    ik_joints = duplicate_and_rename_chain("IK")

    # Create RK chain and get the renamed joints
    rk_joints = duplicate_and_rename_chain("RK")

    # Change joint radius for FK, IK, and RK chains
    joint_data = {'FK': original_joints, 'IK': ik_joints, 'RK': rk_joints}

    # Set radius for FK, IK, and RK joints
    for fk_joint in joint_data['FK']:
        cmds.setAttr(f"{fk_joint}.radius", 0.75)

    for ik_joint in joint_data['IK']:
        cmds.setAttr(f"{ik_joint}.radius", 0.5)

    for rk_joint in joint_data['RK']:
        cmds.setAttr(f"{rk_joint}.radius", 1.0)  # RK joints set to 1.0

    # Create constraints for the RK joints
    constraints = create_constraints(joint_data)

    # Set drawing overrides for joint colors
    set_joint_colors(joint_data)

    # Create IK handle from the first IK joint to the last IK joint
    create_ik_handle(ik_joints)

    # Create control attribute for the RK joints
    control_attribute_name = create_control_attribute(original_names[0])

    # Create reverse node and tie it to the control attribute
    create_reverse_node(control_attribute_name)

    # Connect reverse node to IK weights and control attribute to FK weights
    connect_weights_to_reverse_and_control(constraints, control_attribute_name)

    # Print the names of the created constraints with weight aliases
    print("Created Constraints and Weight Aliases:")
    constraint_dict = {}
    for constraint, weight_alias in constraints:
        if constraint not in constraint_dict:
            constraint_dict[constraint] = []
        constraint_dict[constraint].extend(weight_alias)

    for constraint, weight_aliases in constraint_dict.items():
        print(f"Constraint: {constraint}, Weight Aliases: {list(set(weight_aliases))}")


def create_constraints(joint_data):
    constraints = []  # List to store the names of constraints and their weight aliases
    for rk_joint in joint_data['RK']:
        # Extract the numeric suffix from the RK joint name
        match = re.search(r'(\d+)$', rk_joint)
        if match:
            suffix = match.group(1)

            # Construct the corresponding FK and IK joint names
            fk_joint = rk_joint.replace('RK', 'FK')
            ik_joint = rk_joint.replace('RK', 'IK')

            # Debugging output
            print(f"Creating constraints for RK joint: {rk_joint}")
            print(f"Corresponding FK joint: {fk_joint}, IK joint: {ik_joint}")

            # Check if FK joint exists and create constraints if it does
            if cmds.objExists(fk_joint):
                # Create parent constraint from FK to RK joint
                parent_constraint = cmds.parentConstraint(fk_joint, rk_joint, maintainOffset=True)
                # Create scale constraint from FK to RK joint
                scale_constraint = cmds.scaleConstraint(fk_joint, rk_joint, maintainOffset=True)
                # Get the weight aliases
                parent_weight_alias = cmds.parentConstraint(parent_constraint, query=True, weightAliasList=True)
                scale_weight_alias = cmds.scaleConstraint(scale_constraint, query=True, weightAliasList=True)
                # Store the constraint names and weight aliases
                constraints.append((parent_constraint[0], parent_weight_alias))
                constraints.append((scale_constraint[0], scale_weight_alias))
                print(f"Created constraints from {fk_joint} to {rk_joint}")

            # Check if IK joint exists and create constraints if it does
            if cmds.objExists(ik_joint):
                # Create parent constraint from IK to RK joint
                parent_constraint = cmds.parentConstraint(ik_joint, rk_joint, maintainOffset=True)
                # Create scale constraint from IK to RK joint
                scale_constraint = cmds.scaleConstraint(ik_joint, rk_joint, maintainOffset=True)
                # Get the weight aliases
                parent_weight_alias = cmds.parentConstraint(parent_constraint, query=True, weightAliasList=True)
                scale_weight_alias = cmds.scaleConstraint(scale_constraint, query=True, weightAliasList=True)
                # Store the constraint names and weight aliases
                constraints.append((parent_constraint[0], parent_weight_alias))
                constraints.append((scale_constraint[0], scale_weight_alias))
                print(f"Created constraints from {ik_joint} to {rk_joint}")

    return constraints  # Return the list of constraint names and their weight aliases


def set_joint_colors(joint_data):
    # FK joints should be blue (index 6), IK joints red (index 13), and RK joints green (index 14)

    for fk_joint in joint_data['FK']:
        cmds.setAttr(f"{fk_joint}.overrideEnabled", 1)
        cmds.setAttr(f"{fk_joint}.overrideColor", 6)  # Blue

    for ik_joint in joint_data['IK']:
        cmds.setAttr(f"{ik_joint}.overrideEnabled", 1)
        cmds.setAttr(f"{ik_joint}.overrideColor", 13)  # Red

    for rk_joint in joint_data['RK']:
        cmds.setAttr(f"{rk_joint}.overrideEnabled", 1)
        cmds.setAttr(f"{rk_joint}.overrideColor", 14)  # Green


def create_ik_handle(ik_joints):
    # Print each IK joint in order
    print("IK Joints in order:")
    for joint in ik_joints:
        print(joint)

    # Create an IK handle from the first IK joint to the last IK joint
    if len(ik_joints) >= 2:
        start_joint = ik_joints[0]
        end_joint = ik_joints[-1]
        ik_handle_name = cmds.ikHandle(sj=start_joint, ee=end_joint)[0]  # Removed the 'sol' flag
        print(f"Created IK handle: {ik_handle_name} from {start_joint} to {end_joint}")


def create_control_attribute(original_name):
    # Check if the Transform_Ctrl exists
    transform_ctrl = cmds.ls("Transform_Ctrl", type='transform')
    if not transform_ctrl:
        # Create a NURBS curve at the origin if it doesn't exist
        transform_ctrl = cmds.circle(name="Transform_Ctrl", center=(0, 0, 0), normal=(0, 1, 0), radius=1.0)[0]
        print("Created Transform_Ctrl at the origin.")
    else:
        transform_ctrl = transform_ctrl[0]  # Get the first found transform_ctrl

    # Generate the attribute name based on the original joint name
    attribute_name = re.sub(r'FK_', '', original_name)
    attribute_name = re.sub(r'Jnt_\d*', '', attribute_name)
    attribute_name = f"{attribute_name.replace('_', '')}_IKFK"  # Final attribute name

    # Add the attribute to the transform control
    cmds.addAttr(transform_ctrl, longName=attribute_name, attributeType='float', min=0, max=1, defaultValue=0, keyable=True)
    print(f"Added attribute {attribute_name} to {transform_ctrl}")

    return attribute_name


def create_reverse_node(control_attribute_name):
    # Check if the reverse node already exists
    reverse_node_name = f"{control_attribute_name}_Rev"
    if not cmds.objExists(reverse_node_name):
        # Create the reverse node
        reverse_node = cmds.createNode('reverse', name=reverse_node_name)
        print(f"Created reverse node: {reverse_node}")

        # Connect the control attribute to the reverse node
        cmds.connectAttr(f"Transform_Ctrl.{control_attribute_name}", f"{reverse_node}.inputX")
        print(f"Connected Transform_Ctrl.{control_attribute_name} to {reverse_node}.inputX")
    else:
        print(f"Reverse node {reverse_node_name} already exists.")


def connect_weights_to_reverse_and_control(constraints, control_attribute_name):
    if control_attribute_name is None:
        cmds.warning("Control attribute not provided.")
        return

    # Find the reverse node
    reverse_node = f"{control_attribute_name}_Rev"
    if not cmds.objExists(reverse_node):
        cmds.warning(f"Reverse node {reverse_node} does not exist.")
        return

    # Ensure the Transform_Ctrl attribute is connected to the reverse node inputX
    if not cmds.isConnected(f"Transform_Ctrl.{control_attribute_name}", f"{reverse_node}.inputX"):
        cmds.connectAttr(f"Transform_Ctrl.{control_attribute_name}", f"{reverse_node}.inputX")
        print(f"Connected Transform_Ctrl.{control_attribute_name} to {reverse_node}.inputX")
    else:
        print(f"Connection Transform_Ctrl.{control_attribute_name} to {reverse_node}.inputX already exists")

    # Connect the reverse node outputX to IK weights and control attribute to FK weights
    for constraint, weight_aliases in constraints:
        for alias in weight_aliases:
            # Check for existing connections before making new ones
            if 'IK' in alias:  # IK weight
                if not cmds.isConnected(f"{reverse_node}.outputX", f"{constraint}.{alias}"):
                    cmds.connectAttr(f"{reverse_node}.outputX", f"{constraint}.{alias}")
                    print(f"Connected {reverse_node}.outputX to {constraint}.{alias} (IK weight)")
                else:
                    print(f"Connection {reverse_node}.outputX to {constraint}.{alias} already exists (IK weight)")
            elif 'FK' in alias:  # FK weight
                if not cmds.isConnected(f"Transform_Ctrl.{control_attribute_name}", f"{constraint}.{alias}"):
                    cmds.connectAttr(f"Transform_Ctrl.{control_attribute_name}", f"{constraint}.{alias}")
                    print(f"Connected Transform_Ctrl.{control_attribute_name} to {constraint}.{alias} (FK weight)")
                else:
                    print(f"Connection Transform_Ctrl.{control_attribute_name} to {constraint}.{alias} already exists (FK weight)")


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
    create_rk_button.clicked.connect(duplicate_joint_chains)
    layout.addWidget(create_rk_button)

    # Set the layout on the window
    window.setLayout(layout)

    # Show the window
    window.show()

# Run the UI
create_ui()