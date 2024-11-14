import maya.cmds as cmds

def copy_translate_keys():
    # Get the selected objects
    selected_objects = cmds.ls(selection=True)

    if len(selected_objects) != 2:
        cmds.error("Please select exactly two objects.")
        return

    source_object = selected_objects[0]
    target_object = selected_objects[1]

    # Define the translate attributes
    translate_attrs = [f"{source_object}.translateX", f"{source_object}.translateY", f"{source_object}.translateZ"]

    for attr in translate_attrs:
        # Get keyframe times and values
        key_times = cmds.keyframe(attr, query=True, timeChange=True)
        key_values = cmds.keyframe(attr, query=True, valueChange=True)

        if key_times and key_values:
            # Get the attribute name without the object prefix
            attr_name = attr.split('.')[-1]
            target_attr = f"{target_object}.{attr_name}"

            # Ensure the target attribute exists
            if not cmds.objExists(target_attr):
                cmds.warning(f"Attribute {target_attr} does not exist on {target_object}.")
                continue

            # Set keyframes on the target object
            for time, value in zip(key_times, key_values):
                cmds.setKeyframe(target_attr, time=time, value=value)

            # Set the tangents to stepped
            cmds.keyTangent(target_attr, edit=True, time=(key_times[0], key_times[-1]), inTangentType='step', outTangentType='step')

    print(f"Copied translate keyframes from {source_object} to {target_object} and set tangents to stepped.")

# Run the function
copy_translate_keys()