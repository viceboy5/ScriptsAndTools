import maya.cmds as cmds

def create_driven_key():
    # Ensure the correct selection
    sel = cmds.ls(selection=True)
    if len(sel) != 2:
        cmds.error("Please select exactly two objects: the non-parent constraint and the parent constraint.")

    non_parent_constraint = sel[0]
    parent_constraint = sel[1]

    # Verify that the second selection is a parent constraint
    if not cmds.objectType(parent_constraint, isType="parentConstraint"):
        cmds.error(f"The second selected object ({parent_constraint}) must be a parent constraint.")

    # Add the follow attribute to the non-parent constraint object
    attr_name = "follow"
    if not cmds.attributeQuery(attr_name, node=non_parent_constraint, exists=True):
        cmds.addAttr(non_parent_constraint, longName=attr_name, attributeType="enum",
                     enumName="Transform:Dio Two Hand:Dio Right Hand:Dio Left Hand:Apollo Hand",
                     keyable=True)

    # Map the enum options to the weights of the parent constraint
    weight_mappings = {
        "Transform": "Transform_CtrlW0",
        "Dio Two Hand": "Two_Handed_Prop_CtrlW1",
        "Dio Right Hand": "R_Hand_Prop_CtrlW2",
        "Dio Left Hand": "L_Hand_Prop_CtrlW3",
        "Apollo Hand": "Prop_CtrlW4"
    }

    # Get the weight aliases for the parent constraint
    weight_aliases = cmds.parentConstraint(parent_constraint, query=True, weightAliasList=True)
    print(f"Weight aliases found: {weight_aliases}")

    # Print each weight alias individually
    for alias in weight_aliases:
        print(f"Found weight alias: {alias}")

    # Create driven keys
    for i, (option, expected_alias) in enumerate(weight_mappings.items()):
        # Find the actual alias that matches the expected alias
        matching_alias = next((alias for alias in weight_aliases if expected_alias in alias), None)
        if not matching_alias:
            cmds.error(f"Weight alias matching '{expected_alias}' not found in parent constraint {parent_constraint}.")

        # Create driven keyframes
        for alias in weight_aliases:
            weight_value = 1.0 if alias == matching_alias else 0.0
            cmds.setDrivenKeyframe(f"{parent_constraint}.{alias}",
                                   currentDriver=f"{non_parent_constraint}.{attr_name}",
                                   driverValue=i, value=weight_value)

    print(f"Driven keys successfully created between {non_parent_constraint}.follow and {parent_constraint} weights.")

# Run the function
create_driven_key()
