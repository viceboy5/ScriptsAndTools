import maya.cmds as cmds


def add_influences_to_constraint():
    # Define the influence objects to add
    influence_names = [
        "Dionysus_Asset_Rig:Two_Handed_Prop_Ctrl",
        "Dionysus_Asset_Rig:R_Hand_Prop_Ctrl",
        "Dionysus_Asset_Rig:L_Hand_Prop_Ctrl",
        "Apollo:Prop_Ctrl"
    ]

    # Check for selection
    selection = cmds.ls(selection=True)
    if not selection:
        cmds.warning("Please select a constraint.")
        return

    constraint = selection[0]

    # Determine the type of constraint
    constraint_types = ["parentConstraint", "pointConstraint", "orientConstraint", "scaleConstraint", "aimConstraint"]
    constraint_type = None

    for c_type in constraint_types:
        if cmds.objectType(constraint, isType=c_type):
            constraint_type = c_type
            break

    if not constraint_type:
        cmds.warning(f"Selected object '{constraint}' is not a supported constraint type.")
        return

    # Add each influence if it exists
    for influence in influence_names:
        if cmds.objExists(influence):
            try:
                if constraint_type == "parentConstraint":
                    cmds.parentConstraint(influence, constraint, edit=True, weight=0)
                elif constraint_type == "pointConstraint":
                    cmds.pointConstraint(influence, constraint, edit=True, weight=0)
                elif constraint_type == "orientConstraint":
                    cmds.orientConstraint(influence, constraint, edit=True, weight=0)
                elif constraint_type == "scaleConstraint":
                    cmds.scaleConstraint(influence, constraint, edit=True, weight=0)
                elif constraint_type == "aimConstraint":
                    cmds.aimConstraint(influence, constraint, edit=True, weight=0)

                print(f"Added influence: {influence} to {constraint} ({constraint_type})")
            except Exception as e:
                print(f"Error adding influence {influence}: {e}")
        else:
            print(f"Influence not found: {influence}")


# Run the function
add_influences_to_constraint()
