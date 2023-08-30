import maya.cmds as cmds

def createJoints():
    """
    Creates a joint at each selection transform.

    Returns: [joints]
    """
    sels = cmds.ls(sl=True)
    newJoints = []

    for sel in sels:
        pos = cmds.xform(sel, q=True, worldSpace=True, translation=True)
        cmds.select(cl=True)

        jnt = cmds.joint()
        newJoints.append(jnt)
        cmds.xform(jnt, worldSpace=True, translation=pos)
    cmds.select(newJoints, r=True)

    return newJoints

createJoints()
