import maya.cmds as cmds

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


def change_outline_color(color_index):
    try:
        sels = cmds.ls(sl=True)
        for sel in sels:
            shape = cmds.listRelatives(sel)
            cmds.setAttr(shape[0] + ".overrideEnabled", 1)
            cmds.setAttr(shape[0] + ".overrideColor", color_index)
    except:
        print("Input must be between 1 and 32 to assign new Index Color")

def run_control_creation(color_name):
    success = create_controls_clicked(color_name)
    if success:
        print(f"Controls created with color: {color_name}")
    else:
        print("Control creation failed due to no selection.")
