import maya.cmds as cmds


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


# Run the function
create_joint_at_cluster_transforms()
