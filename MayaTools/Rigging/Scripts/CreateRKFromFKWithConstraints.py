import maya.cmds as cmds
import re

def duplicate_joint_chain(base_joint, suffix):
    original_chain = cmds.listRelatives(base_joint, allDescendents=True, type='joint', fullPath=True) or []
    original_chain = [base_joint] + original_chain
    original_chain = cmds.ls(original_chain, long=True)
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
        alias_info = cmds.aliasAttr(constraint_node, q=True)
        if not alias_info:
            return []
        alias_dict = dict(zip(alias_info[::2], alias_info[1::2]))
        weight_aliases = [alias for alias, attr in alias_dict.items() if '.w[' in attr]
        return weight_aliases
    except Exception as e:
        print(f"# Warning: Failed to get weight aliases for {constraint_node}: {e}")
        return []

def create_constraints(fk_chain, ik_chain, rk_chain):
    constraints_info = []

    for fk, ik, rk in zip(fk_chain, ik_chain, rk_chain):
        base_name = rk.replace('_RK_', '_')

        if cmds.objExists(fk):
            parent_fk = cmds.parentConstraint(fk, rk, mo=True, name=f"{rk}_parentConstraint")[0]
            parent_fk_aliases = cmds.parentConstraint(parent_fk, q=True, weightAliasList=True) or []
            constraints_info.append((parent_fk, parent_fk_aliases, 'FK'))

        if cmds.objExists(ik):
            parent_ik = cmds.parentConstraint(ik, rk, mo=True, name=f"{rk}_parentConstraint")[0]
            parent_ik_aliases = cmds.parentConstraint(parent_ik, q=True, weightAliasList=True) or []
            constraints_info.append((parent_ik, parent_ik_aliases, 'IK'))

        if cmds.objExists(fk):
            scale_fk = cmds.scaleConstraint(fk, rk, mo=True, name=f"{rk}_scaleConstraint")[0]
            scale_fk_aliases = cmds.scaleConstraint(scale_fk, q=True, weightAliasList=True) or []
            constraints_info.append((scale_fk, scale_fk_aliases, 'FK'))

        if cmds.objExists(ik):
            scale_ik = cmds.scaleConstraint(ik, rk, mo=True, name=f"{rk}_scaleConstraint")[0]
            scale_ik_aliases = cmds.scaleConstraint(scale_ik, q=True, weightAliasList=True) or []
            constraints_info.append((scale_ik, scale_ik_aliases, 'IK'))

    return constraints_info

def set_joint_colors(joint_data):
    color_map = {'FK': 6, 'IK': 13, 'RK': 14}
    for chain_type, joints in joint_data.items():
        for joint in joints:
            if cmds.objExists(joint):
                cmds.setAttr(f"{joint}.overrideEnabled", 1)
                cmds.setAttr(f"{joint}.overrideColor", color_map[chain_type])

def set_joint_radii(joint_data):
    radius_map = {'FK': 0.75, 'IK': 0.5, 'RK': 1.0}
    for chain_type, joints in joint_data.items():
        for joint in joints:
            if cmds.objExists(joint):
                cmds.setAttr(f"{joint}.radius", radius_map[chain_type])

def create_ik_handle(ik_chain):
    if len(ik_chain) >= 2:
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

    set_joint_colors(joint_data)
    set_joint_radii(joint_data)
    create_ik_handle(ik_chain)

    control_attr_name = create_control_attribute(fk_root)
    reverse_node = create_reverse_node(control_attr_name)
    connect_weights_to_reverse_and_control(constraints, control_attr_name, reverse_node)

    print("Finished rig setup with FK/IK/RK chains, constraints, control attribute, and reverse node.")

def create_control_attribute(original_name):
    transform_ctrl = cmds.ls("Transform_Ctrl", type='transform')
    if not transform_ctrl:
        transform_ctrl = cmds.circle(name="Transform_Ctrl", center=(0, 0, 0), normal=(0, 1, 0), radius=1.0)[0]
        print("Created Transform_Ctrl at the origin.")
    else:
        transform_ctrl = transform_ctrl[0]

    attribute_name = re.sub(r'FK_', '', original_name)
    attribute_name = re.sub(r'Jnt_\d*', '', attribute_name)
    attribute_name = f"{attribute_name.replace('_', '')}_IKFK"

    if not cmds.attributeQuery(attribute_name, node=transform_ctrl, exists=True):
        cmds.addAttr(transform_ctrl, longName=attribute_name, attributeType='float', min=0, max=1, defaultValue=0, keyable=True)
        print(f"Added attribute {attribute_name} to {transform_ctrl}")

    return attribute_name

def create_reverse_node(control_attribute_name):
    reverse_node_name = f"{control_attribute_name}_Rev"
    if not cmds.objExists(reverse_node_name):
        reverse_node = cmds.createNode('reverse', name=reverse_node_name)
        cmds.connectAttr(f"Transform_Ctrl.{control_attribute_name}", f"{reverse_node}.inputX")
        print(f"Created reverse node: {reverse_node} and connected input.")
    else:
        reverse_node = reverse_node_name
        print(f"Reverse node {reverse_node} already exists.")
    return reverse_node

def connect_weights_to_reverse_and_control(constraints, control_attribute_name, reverse_node):
    if control_attribute_name is None or reverse_node is None:
        cmds.warning("Control attribute or reverse node not provided.")
        return

    if not cmds.isConnected(f"Transform_Ctrl.{control_attribute_name}", f"{reverse_node}.inputX"):
        cmds.connectAttr(f"Transform_Ctrl.{control_attribute_name}", f"{reverse_node}.inputX")

    connected = set()

    for constraint, weight_aliases, source_type in constraints:
        for alias in weight_aliases:
            plug = f"{constraint}.{alias}"
            if plug in connected:
                continue  # Skip already connected alias
            if source_type == 'IK':
                if not cmds.isConnected(f"{reverse_node}.outputX", plug):
                    cmds.connectAttr(f"{reverse_node}.outputX", plug)
                    connected.add(plug)
            elif source_type == 'FK':
                if not cmds.isConnected(f"Transform_Ctrl.{control_attribute_name}", plug):
                    cmds.connectAttr(f"Transform_Ctrl.{control_attribute_name}", plug)
                    connected.add(plug)


if __name__ == "__main__":
    process_selected_joint()
