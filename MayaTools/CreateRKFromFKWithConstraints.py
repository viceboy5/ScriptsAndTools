import maya.cmds as cmds
import re


def duplicate_and_create_constraints():
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

        # Dictionary to store joint names with their numeric suffixes
        joint_suffix_map = {}

        # Duplicate each joint individually and store it
        for original_joint in original_joints:
            # Duplicate the current joint without hierarchy
            duplicated_joint = cmds.duplicate(original_joint, parentOnly=True)[0]
            duplicated_joints.append(duplicated_joint)

            # Extract the numeric suffix from the original joint name
            match = re.search(r'(\d+)$', original_joint)
            if match:
                joint_suffix_map[duplicated_joint] = int(match.group(1))

        # Sort the duplicated joints based on their numeric suffixes
        sorted_joints = sorted(duplicated_joints, key=lambda joint: joint_suffix_map[joint])

        # Rebuild the hierarchy based on the sorted joints
        for i in range(1, len(sorted_joints)):
            # Reparent each joint to its previous joint in the sorted order
            cmds.parent(sorted_joints[i], sorted_joints[i - 1])

        # Rename the duplicated joints according to the original names, replacing 'FK' with the provided suffix
        renamed_joints = []
        for original_name, duplicated_joint in zip(original_names, duplicated_joints):
            # Replace 'FK' with the given suffix while keeping the numeric suffix
            new_name = re.sub(r'FK(.*?)(\d+)', rf'{suffix}\1\2', original_name)

            # Rename the duplicated joint
            cmds.rename(duplicated_joint, new_name)
            renamed_joints.append(new_name)  # Store the new name

        return renamed_joints  # Return the sorted joint names for constraint creation

    # Create IK chain and get the sorted joint names
    ik_joints = duplicate_and_rename_chain("IK")

    # Create RK chain and get the sorted joint names
    rk_joints = duplicate_and_rename_chain("RK")

    # Call the function to create constraints
    create_constraints(ik_joints, rk_joints)


def create_constraints(ik_joints, rk_joints):
    for rk_joint in rk_joints:
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
                cmds.parentConstraint(fk_joint, rk_joint, maintainOffset=True)
                # Create scale constraint from FK to RK joint
                cmds.scaleConstraint(fk_joint, rk_joint, maintainOffset=True)
                print(f"Created constraints from {fk_joint} to {rk_joint}")

            # Check if IK joint exists and create constraints if it does
            if cmds.objExists(ik_joint):
                # Create parent constraint from IK to RK joint
                cmds.parentConstraint(ik_joint, rk_joint, maintainOffset=True)
                # Create scale constraint from IK to RK joint
                cmds.scaleConstraint(ik_joint, rk_joint, maintainOffset=True)
                print(f"Created constraints from {ik_joint} to {rk_joint}")


# Execute the function
duplicate_and_create_constraints()
