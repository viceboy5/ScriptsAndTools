import maya.cmds as cmds


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


# Call the function to find and create constraints
find_and_create_constraints()
