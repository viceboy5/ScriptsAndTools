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


# UI function to take input and trigger renaming
def showRenameUI():
    if cmds.window("renameWindow", exists=True):
        cmds.deleteUI("renameWindow")

    window = cmds.window("renameWindow", title="Rename Sequentially", widthHeight=(300, 100))
    cmds.columnLayout(adjustableColumn=True)

    # Input field
    nameField = cmds.textField(placeholderText="Enter name with # for numbering (e.g. FK_Jnt_##)")

    # Button to trigger renaming
    def onRenameButtonClicked(*args):
        nameString = cmds.textField(nameField, query=True, text=True)
        RenameSequentially(nameString)

    cmds.button(label="Rename", command=onRenameButtonClicked)

    cmds.showWindow(window)


# Call the UI
showRenameUI()
