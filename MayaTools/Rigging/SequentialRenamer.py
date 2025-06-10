import maya.cmds as cmds


def RenameSequentially(txt):
    count = txt.count("#")

    if count == 0:
        print("# should be used to indicate the numbering in between the Name and NodeType.")
        return

    nums_placeholder = "#" * count
    x = txt.find(nums_placeholder)

    if x == -1:
        print("# should only be used in between the Name and NodeType. Ensure all arguments are named appropriately.")
    else:
        parts = txt.partition(nums_placeholder)

        sels = cmds.ls(sl=True)
        for i, sel in enumerate(sels):
            newNum = str(i + 1).zfill(count)  # Correctly format the numbering
            new_name = parts[0] + newNum + parts[2]
            cmds.rename(sel, new_name)