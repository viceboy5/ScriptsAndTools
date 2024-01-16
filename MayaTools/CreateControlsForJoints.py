import maya.cmds as cmds

def create_group_and_ctrl(joint):
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

    return empty_group, circle

def create_ctrls_for_joint_chain(top_joint):
    # Get the joint chain
    joint_chain = cmds.listRelatives(top_joint, allDescendents=True, type="joint")

    if not joint_chain:
        cmds.warning("No joints found in the chain.")
        return

    # Iterate through the joint chain
    for joint in joint_chain:
        # Create an empty group and control circle
        create_group_and_ctrl(joint)

# Get the selected joint
selected_joint = cmds.ls(selection=True, type="joint")

if not selected_joint:
    cmds.warning("Please select a joint to start from.")
else:
    # Call the function to create groups and control circles for the joint chain
    create_ctrls_for_joint_chain(selected_joint[0])
