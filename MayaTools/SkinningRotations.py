import maya.cmds as cmds

def set_rotation_keys():
    # Set initial keyframes at frame 0
    cmds.setKeyframe(at='rotateX', v=0)
    cmds.setKeyframe(at='rotateY', v=0)
    cmds.setKeyframe(at='rotateZ', v=0)
  # Adjust 300 to the desired end frame

    # Set rotation keys for Z axis
    cmds.setKeyframe(at='rotateZ', v=75, t=15)
    cmds.setKeyframe(at='rotateZ', v=0, t=30)
    cmds.setKeyframe(at='rotateZ', v=-45, t=45)
    cmds.setKeyframe(at='rotateZ', v=0, t=60)

    cmds.setKeyframe(at='rotateY', v = 0, t= 60)
    cmds.setKeyframe(at='rotateY', v=45, t=75)
    cmds.setKeyframe(at='rotateY', v=0, t=90)
    cmds.setKeyframe(at='rotateY', v=-45, t=105)
    cmds.setKeyframe(at='rotateY', v=0, t=120)

    cmds.setKeyframe(at='translateX', v=0, t=120)
    cmds.setKeyframe(at='translateX', v=120, t=135)



# Call the function to set the rotation keys
set_rotation_keys()
