import maya.cmds as cmds


def add_influence_to_constraint():
    """
    Adds a new influence to a selected parent or scale constraint.
    The first selected object should be the constraint, and the second should be the new influence.
    """
    selection = cmds.ls(selection=True)

    if len(selection) != 2:
        cmds.error("Please select exactly two objects: a constraint and the new influence.")
        return

    constraint = selection[0]
    new_influence = selection[1]

    # Check if the selected object is a supported constraint type
    constraint_type = cmds.objectType(constraint)
    if constraint_type not in ['parentConstraint', 'scaleConstraint']:
        cmds.error("The first selected object must be a parent or scale constraint.")
        return

    # Check if the new influence is a valid transform node
    if not cmds.objectType(new_influence, isType='transform'):
        cmds.error("The second selected object must be a transform node.")
        return

    try:
        if constraint_type == 'parentConstraint':
            # Add the new influence to a parentConstraint
            cmds.parentConstraint(new_influence, constraint, e=True, weight=1.0)
        elif constraint_type == 'scaleConstraint':
            # Add the new influence to a scaleConstraint
            cmds.scaleConstraint(new_influence, constraint, e=True, weight=1.0)

        print(f"Successfully added {new_influence} to {constraint}.")
    except RuntimeError as e:
        cmds.error(f"Failed to add influence: {e}")


# Run the function
add_influence_to_constraint()
