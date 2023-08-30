import maya.cmds as cmds



def RenameSequentially(txt):

    count = txt.count("#")
    nums = "#" * count

    x = txt.find(nums)
    if x == -1:
        print("# should only be used in between the Name and NodeType.  Ensure all arguments are named appropriately")
    else:
        parts = txt.partition(nums)

        sels = cmds.ls(sl=True)
        for sel in sels:
            i = sels.index(sel)
            newNum = i + 1
            nums = nums.replace(nums, str(newNum))
            nums = nums.zfill(count)

            cmds.rename([sel], parts[0] + nums + parts[2])


RenameSequentially("Leg_#######_Jnt")



