import maya.cmds as cmds
import re

def duplicate_with_constraints_and_driven_keys(node):
    # Duplicate the node and its hierarchy
    duplicated_node = cmds.duplicate(node, renameChildren=True)[0]

    # Get all nodes in the duplicated hierarchy
    all_nodes = cmds.listRelatives(duplicated_node, allDescendents=True, fullPath=True) or []
    all_nodes.append(duplicated_node)

    # Rename nodes by removing the "1" from the end
    for node in all_nodes:
        short_name = node.split('|')[-1]  # Get the short name of the node
        new_name = re.sub(r'1$', '', short_name)
        cmds.rename(node, new_name)

    # Get all nodes in the duplicated hierarchy again after renaming
    all_nodes = cmds.listRelatives(duplicated_node, allDescendents=True, fullPath=True) or []
    all_nodes.append(duplicated_node)

    # Print the full path names of all nodes after renaming
    for node in all_nodes:
        print(f"Full path name after renaming: {node}")

    # Store the full path name of the node ending in COG_Ctrl
    cog_ctrl = None
    for node in all_nodes:
        if node.endswith("COG_Ctrl"):
            cog_ctrl = node
            break

    # Debug print statement to check if cog_ctrl is set correctly
    print(f"Identified COG_Ctrl: {cog_ctrl}")

    # Get the parent constraints in the duplicated node's hierarchy again after renaming
    duplicated_constraints = cmds.listRelatives(duplicated_node, type='parentConstraint', allDescendents=True, fullPath=True) or []

    # Delete the duplicated constraints
    for duplicated_constraint in duplicated_constraints:
        cmds.delete(duplicated_constraint)

    # Store the full path names of the Transform_Ctrl and COG_Ctrl_Grp
    body_part_name = re.sub(r'\d+$', '', duplicated_node.split('|')[-1])
    transform_ctrl = f"{duplicated_node}|{body_part_name}_Transform_Ctrl"
    cog_ctrl_grp = f"{transform_ctrl}|COG_Ctrl_Grp"

    return duplicated_node, transform_ctrl, cog_ctrl_grp, cog_ctrl

def create_parent_constraint(duplicated_hierarchy, transform_ctrl, cog_ctrl_grp):
    # Define the nodes to be selected
    nodes_to_select = [
        "Apollo1:Prop_Ctrl",
        "Dionysus_Asset_Rig:L_Hand_Prop_Ctrl",
        "Dionysus_Asset_Rig:R_Hand_Prop_Ctrl",
        transform_ctrl,
        cog_ctrl_grp
    ]

    # Ensure all nodes exist
    for node in nodes_to_select:
        if not cmds.objExists(node):
            cmds.warning(f"{node} does not exist.")
            return

    # Select the nodes
    cmds.select(nodes_to_select)

    # Create the parent constraint
    constraint = cmds.parentConstraint(nodes_to_select, maintainOffset=True)[0]
    print(f"Created parent constraint: {constraint}")

    return constraint

def setup_driven_keys(constraint, cog_ctrl):
    # Print the names of the constraint and cog_ctrl
    print(f"Constraint: {constraint}, Cog_Ctrl: {cog_ctrl}")

    # Check if the "Follow" attribute exists
    if not cmds.attributeQuery("Follow", node=cog_ctrl, exists=True):
        cmds.error(f"The 'Follow' attribute does not exist on {cog_ctrl}.")

    # Define the follow options and corresponding weights
    follow_options = ["Transform", "Dio Right Hand", "Dio Left Hand", "Apollo Hand"]  # Swapped the order

    weight_aliases = cmds.parentConstraint(constraint, query=True, weightAliasList=True)

    # Set driven keys for each follow option
    for i, option in enumerate(follow_options):
        for j, alias in enumerate(weight_aliases):
            weight_value = 1.0 if i == j else 0.0
            cmds.setDrivenKeyframe(f"{constraint}.{alias}", currentDriver=f"{cog_ctrl}.Follow", driverValue=i, value=weight_value)
            print(f"Set driven key for {constraint}.{alias} with {cog_ctrl}.Follow={i} to {weight_value}")

# Example usage
selected_nodes = cmds.ls(selection=True)
for node in selected_nodes:
    duplicated_hierarchy, transform_ctrl, cog_ctrl_grp, cog_ctrl = duplicate_with_constraints_and_driven_keys(node)
    constraint = create_parent_constraint(duplicated_hierarchy, transform_ctrl, cog_ctrl_grp)
    setup_driven_keys(constraint, cog_ctrl)