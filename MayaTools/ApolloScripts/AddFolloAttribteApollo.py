import maya.cmds as cmds


def add_follow_attribute():
    # Get the current selection
    selection = cmds.ls(selection=True)

    if not selection:
        cmds.warning("No objects selected. Please select at least one object.")
        return

    # Define the enum options
    enum_options = "Transform:Dio Two Hand:Dio Right Hand:Dio Left Hand:Apollo Hand"

    for obj in selection:
        # Check if the attribute already exists
        if not cmds.attributeQuery("Follow", node=obj, exists=True):
            # Add the enum attribute
            cmds.addAttr(obj, longName="Follow", attributeType="enum", enumName=enum_options, keyable=True)
            print(f"Added 'Follow' attribute to {obj}.")
        else:
            print(f"'Follow' attribute already exists on {obj}.")


# Run the function
add_follow_attribute()
