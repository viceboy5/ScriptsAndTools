import maya.cmds as cmds
def change_outline_color(color):

    try:
        sels = cmds.ls(sl=True)
        for sel in sels:
            shape = cmds.listRelatives(sel)
            cmds.setAttr(shape[0] + ".overrideEnabled", 1)
            cmds.setAttr(shape[0] + ".overrideColor", color)
    except:
        print("Input must be between 1 and 32 to assign new Index Color")

#"color" must be between 1 and 32
change_outline_color(6)

