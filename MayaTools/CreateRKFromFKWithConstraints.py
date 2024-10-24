import maya.cmds as cmds
import re


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
    # Find the Transform_Ctrl in the scene
    transform_ctrl = cmds.ls("Transform_Ctrl", type='transform')
    if not transform_ctrl:
        cmds.warning("Transform_Ctrl not found in the scene.")
        return None

    transform_ctrl = transform_ctrl[0]  # Get the first found transform_ctrl

    # Generate the attribute name based on the original joint name
    attribute_name = re.sub(r'FK_', '', original_name)
    attribute_name = re.sub(r'Jnt_\d*', '', attribute_name)
    attribute_name = f"{attribute_name.replace('_', '')}_IKFK"  # Final attribute name

    # Add the attribute to the transform control
    cmds.addAttr(transform_ctrl, longName=attribute_name, attributeType='float', min=0, max=1, defaultValue=0)

    # Make the attribute hidden, keyable, and viewable in the channel box
    cmds.setAttr(f"{transform_ctrl}.{attribute_name}", keyable=True)
    print(f"Created attribute {attribute_name} on Transform_Ctrl")

    return attribute_name


def create_reverse_node(control_attribute_name):
    if control_attribute_name is None:
        return

    # Create the reverse node
    reverse_node = cmds.createNode("reverse", name=f"{control_attribute_name}_Rev")
    cmds.connectAttr(f"Transform_Ctrl.{control_attribute_name}", f"{reverse_node}.inputX")

    print(f"Created reverse node: {reverse_node} for attribute: {control_attribute_name}")


# Call the main function to run the script
duplicate_joint_chains()
