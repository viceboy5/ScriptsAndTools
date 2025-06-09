import maya.cmds as cmds
import re

def duplicate_joint_chain(base_joint, suffix):
    """
    Duplicates the joint chain starting from base_joint and renames it using the given suffix ("IK" or "RK").
    Requires that "FK" is in the original joint names.
    Returns a list of duplicated joint names.
    """
    original_chain = cmds.listRelatives(base_joint, allDescendents=True, type='joint', fullPath=True) or []
    original_chain = [base_joint] + original_chain
    original_chain = cmds.ls(original_chain, long=True)  # ensure proper hierarchy order
    duplicated_chain = []
    joint_mapping = {}

    for joint in original_chain:
        base_name = joint.split('|')[-1]
        if 'FK' not in base_name:
            cmds.error(f"Joint '{base_name}' must contain 'FK' in its name to be duplicated with suffix '{suffix}'.")
        new_name = base_name.replace('FK', suffix)
        dup = cmds.duplicate(joint, parentOnly=True, name=new_name)[0]
        joint_mapping[joint] = dup
        duplicated_chain.append(dup)

    for joint in original_chain:
        parent = cmds.listRelatives(joint, parent=True, type='joint', fullPath=True)
        if parent and parent[0] in joint_mapping:
            cmds.parent(joint_mapping[joint], joint_mapping[parent[0]])

    return duplicated_chain

def get_weight_aliases(constraint_node):
    try:
        # Query aliases as tuples of (alias, full attribute)
        alias_info = cmds.aliasAttr(constraint_node, q=True)
        if not alias_info:
            return []

        # Convert to dict for easier filtering
        alias_dict = dict(zip(alias_info[::2], alias_info[1::2]))

        # Return only aliases that correspond to weight attributes
        weight_aliases = [alias for alias, attr in alias_dict.items() if '.w[' in attr]
        return weight_aliases

    except Exception as e:
        print(f"# Warning: Failed to get weight aliases for {constraint_node}: {e}")
        return []


def create_constraints(fk_chain, ik_chain, rk_chain):
    constraints_info = []

    for fk, ik, rk in zip(fk_chain, ik_chain, rk_chain):
        base_name = rk.replace('_RK_', '_')

        # Parent constraint FK ➜ RK
        if cmds.objExists(fk):
            parent_fk = cmds.parentConstraint(fk, rk, mo=True, name=f"{rk}_parentConstraint")[0]
            parent_fk_aliases = cmds.parentConstraint(parent_fk, q=True, weightAliasList=True) or []
            constraints_info.append((parent_fk, parent_fk_aliases, 'FK'))
            print(f"Parent constraint created: {parent_fk}, weight aliases: {parent_fk_aliases}, source: FK")

        # Parent constraint IK ➜ RK
        if cmds.objExists(ik):
            parent_ik = cmds.parentConstraint(ik, rk, mo=True, name=f"{rk}_parentConstraint")[0]
            parent_ik_aliases = cmds.parentConstraint(parent_ik, q=True, weightAliasList=True) or []
            constraints_info.append((parent_ik, parent_ik_aliases, 'IK'))
            print(f"Parent constraint created: {parent_ik}, weight aliases: {parent_ik_aliases}, source: IK")

        # Scale constraint FK ➜ RK
        if cmds.objExists(fk):
            scale_fk = cmds.scaleConstraint(fk, rk, mo=True, name=f"{rk}_scaleConstraint")[0]
            scale_fk_aliases = cmds.scaleConstraint(scale_fk, q=True, weightAliasList=True) or []
            constraints_info.append((scale_fk, scale_fk_aliases, 'FK'))
            print(f"Scale constraint created: {scale_fk}, weight aliases: {scale_fk_aliases}, source: FK")

        # Scale constraint IK ➜ RK
        if cmds.objExists(ik):
            scale_ik = cmds.scaleConstraint(ik, rk, mo=True, name=f"{rk}_scaleConstraint")[0]
            scale_ik_aliases = cmds.scaleConstraint(scale_ik, q=True, weightAliasList=True) or []
            constraints_info.append((scale_ik, scale_ik_aliases, 'IK'))
            print(f"Scale constraint created: {scale_ik}, weight aliases: {scale_ik_aliases}, source: IK")

    return constraints_info

def set_joint_colors(joint_data):
    """
    Sets override colors for FK, IK, and RK joints.
    FK: Blue (6), IK: Red (13), RK: Green (14)
    """
    color_map = {'FK': 6, 'IK': 13, 'RK': 14}
    for chain_type, joints in joint_data.items():
        for joint in joints:
            if cmds.objExists(joint):
                cmds.setAttr(f"{joint}.overrideEnabled", 1)
                cmds.setAttr(f"{joint}.overrideColor", color_map[chain_type])

def set_joint_radii(joint_data):
    """
    Sets custom joint radii for visual differentiation.
    FK: 0.75, IK: 0.5, RK: 1.0
    """
    radius_map = {'FK': 0.75, 'IK': 0.5, 'RK': 1.0}
    for chain_type, joints in joint_data.items():
        for joint in joints:
            if cmds.objExists(joint):
                cmds.setAttr(f"{joint}.radius", radius_map[chain_type])

def create_ik_handle(ik_chain):
    if len(ik_chain) >= 2:
        # Find the terminal joint in the IK chain
        end_joint = ik_chain[0]
        while True:
            children = cmds.listRelatives(end_joint, children=True, type='joint', fullPath=True)
            if not children:
                break
            end_joint = children[0]
        start_joint = ik_chain[0]
        handle = cmds.ikHandle(sj=start_joint, ee=end_joint, sol='ikRPsolver')[0]
        print(f"Created IK handle: {handle} from {start_joint} to {end_joint}")

def process_selected_joint():
    selection = cmds.ls(selection=True, type='joint')
    if not selection:
        cmds.warning("Please select a root joint to process.")
        return

    fk_root = selection[0]
    fk_chain = cmds.listRelatives(fk_root, allDescendents=True, type='joint', fullPath=True) or []
    fk_chain = [fk_root] + fk_chain
    fk_chain = cmds.ls(fk_chain, long=True)

    ik_chain = duplicate_joint_chain(fk_root, "IK")
    rk_chain = duplicate_joint_chain(fk_root, "RK")

    joint_data = {'FK': fk_chain, 'IK': ik_chain, 'RK': rk_chain}

    constraints = create_constraints(fk_chain, ik_chain, rk_chain)
    print("Constraints info:", constraints)

    set_joint_colors(joint_data)
    set_joint_radii(joint_data)
    create_ik_handle(ik_chain)

    # Create control attribute on Transform_Ctrl based on FK root name
    control_attr_name = create_control_attribute(fk_root)

    # Create reverse node connected to that attribute
    create_reverse_node(control_attr_name)

    # Connect constraint weights to control attribute and reverse node
    connect_weights_to_reverse_and_control(constraints, control_attr_name)
    print("Finished connecting weights.")

    print("Created IK and RK duplicate chains with constraints, colors, radii, IK handle, control attribute, reverse node, and connected weights.")

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


if __name__ == "__main__":
    process_selected_joint()




