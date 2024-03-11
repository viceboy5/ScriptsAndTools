import maya.cmds as cmds

def find_matching_joint(curve_name):
    # Extract joint name from the curve name
    joint_name = curve_name.replace("_Ctrl", "")

    # Check if the joint exists
    if cmds.objExists(joint_name):
        return joint_name
    else:
        return None

def create_constraints(curve_name, joint_name):
    # Create parent constraint
    parent_constraint = cmds.parentConstraint(curve_name, joint_name, mo=True)[0]

    # Create scale constraint
    scale_constraint = cmds.scaleConstraint(curve_name, joint_name, mo=True)[0]

    return parent_constraint, scale_constraint

def find_and_create_constraints():
    # List all curves in the scene
    curves = cmds.ls(type='nurbsCurve')

    for curve in curves:
        # Check if the curve name ends with '_Ctrl'
        if curve.endswith('_Ctrl'):
            # Find matching joint
            joint_name = find_matching_joint(curve)

            if joint_name:
                # Create constraints
                create_constraints(curve, joint_name)
            else:
                print(joint_name)

# Call the function to find and create constraints
find_and_create_constraints()
