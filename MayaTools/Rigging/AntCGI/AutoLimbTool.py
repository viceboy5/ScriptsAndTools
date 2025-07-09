#-----------------------------------------------
#
#antCGI Auto Limb Tool V1, created from part 21 antCGI video tutorial
#
#-----------------------------------------------

import maya.cmds as cmds

def autoLimbTool(*args):
    # Setup the variables which could come from the UI

    # Is this the front or rear Leg
    whichLeg = cmds.optionMenu("LegMenu", q=1, v=1)

    if whichLeg == "Front":
        isRearleg = 0
    else:
        isRearleg = 1

    #Check the checkboxes
    rollCheck = cmds.checkBox("RollCheck", q=1, v=1)
    stretchCheck = cmds.checkBox("StretchCheck", q=1, v=1)

    # How many joints are we working with?
    limbJoints = 4

    # Use this information to start to generate the names we need
    if isRearleg:
        limbType = "Rear"
        print("Working on the REAR leg.")
    else:
        limbType = "Front"
        print("Working on the FRONT leg.")

    # Check the selection is valid
    selectionCheck = cmds.ls(sl=1, type="joint")

    #Error check to make sure a joint is selected
    if not selectionCheck:
        cmds.error("Please select the root joint.")
    else:
        jointRoot = cmds.ls(sl=1, type="joint")[0]
        print(jointRoot)

    #Now we have a selected joint we can check for the prefix to see which side it is
    whichSide = jointRoot[0:2]

    # Make sure the prefix is usable
    if not "L_" in whichSide:
        if not "R_" in whichSide:
            cmds.error("Please use a joint with a usable prefix of either L_ or R_.")

    #Now build the names we need
    limbName = whichSide + "Leg_" + limbType

    mainControl = limbName + "_Ctrl"
    pawControlName = limbName + "_IK_Ctrl"
    kneeControlName = limbName + "_Tibia_Ctrl"
    hockControlName = limbName + "_Hock_Ctrl"
    rootControlName = limbName + "_Root_Ctrl"

    # Build the list of joints we are working with, using the root as a start point

    # Find it's children
    jointHierarchy = cmds.listRelatives(jointRoot, ad =1, type="joint")

    # Add the selected joint into the front of the list
    jointHierarchy.append(jointRoot)

    # Reverse the list so we can work in order
    jointHierarchy.reverse()

    # Clear the selection
    cmds.select(cl=1)

    #Duplicate the main joint chain and rename each joint

    #First define what joint chains we need
    newJointList = ["_IK", "_FK", "_Stretch"]


    # Add the extra driver joints if this is the rear leg
    if isRearleg:
        newJointList.append("_Driver")


    #Build the Joints
    for newJoint in newJointList:
        for i in range(limbJoints):
            newJointName = jointHierarchy[i] + newJoint

            cmds.joint(n=newJointName)
            cmds.matchTransform(newJointName, jointHierarchy[i])
            cmds.makeIdentity(newJointName, a=1, t=0, r=1, s=0)

        cmds.select(cl=1)

    #Constrain the main joints to the ik and fk joints so we can blend between them
    for i in range(limbJoints):
        cmds.parentConstraint((jointHierarchy[i] + "_IK"), (jointHierarchy[i] + "_FK"), jointHierarchy[i], w=1, mo=0)

    # -----------------------------------------------------------------------------------------------------------------
    # Setup FK
    # -----------------------------------------------------------------------------------------------------------------

    # Connect the FK controls in to the new joints
    for i in range(limbJoints):
        cmds.parentConstraint((jointHierarchy[i] + "_FK_Ctrl"), (jointHierarchy[i] + "_FK"), w=1, mo=0)

    #-----------------------------------------------------------------------------------------------------------------
    #Setup IK
    #-----------------------------------------------------------------------------------------------------------------

    #If it's the rear leg, create the ik handle from the femur to the metacarpus
    if isRearleg:
        cmds.ikHandle(n=(limbName + "_Driver_IKHandle"), sol= "ikRPsolver", sj=(jointHierarchy[0] + "_Driver"), ee=(jointHierarchy[3] + "_Driver") )

    # Next, create the main IK handle from the femur to the metacarpus
    cmds.ikHandle(n=(limbName + "_Knee_IKHandle"), sol="ikRPsolver", sj=(jointHierarchy[0] + "_IK"),ee=(jointHierarchy[2] + "_IK"))

    # Finally create the hock IK Handle, from the metatarsus to the metacarpus
    cmds.ikHandle(n=(limbName + "_Hock_IKHandle"), sol="ikSCsolver", sj=(jointHierarchy[2] + "_IK"),ee=(jointHierarchy[3] + "_IK"))

    #Create the Hock control offset group
    cmds.group((limbName + "_Knee_IKHandle"), n=(limbName + "_Knee_Ctrl"))
    cmds.group((limbName + "_Knee_Ctrl"), n=(limbName + "_Knee_Ctrl_Grp"))

    #Find the ankle pivot
    anklePivot = cmds.xform(jointHierarchy[3], q=1, ws=1, piv=1)

    # Set the groups pivot to match the ankle position
    cmds.xform(((limbName + "_Knee_Ctrl"), (limbName + "_Knee_Ctrl_Grp")), ws=1, piv=(anklePivot[0], anklePivot[1], anklePivot[2]))

    #Parent the ik Handle and the group to the paw control
    cmds.parent( (limbName + "_Hock_IKHandle"), pawControlName)

    #If it's the rear leg, adjust the hierarchy so the driver leg controls the ik handles
    if isRearleg:
        cmds.parent((limbName + "_Knee_Ctrl_Grp"), (jointHierarchy[2] + "_Driver"))
        cmds.parent((limbName + "_Hock_IKHandle"), (jointHierarchy[3] + "_Driver"))
        cmds.parent((limbName + "_Driver_IKHandle"), pawControlName)

    else:
        cmds.parent(limbName + "_Knee_Ctrl_Grp", limbName + "_Root_Ctrl")
        cmds.pointConstraint( pawControlName, (limbName + "_Knee_Ctrl_Grp"), w=1)

    # Make the paw control drive the ankle joint to maintain it's orientation
    cmds.orientConstraint( pawControlName, (jointHierarchy[3] + "_IK"), w=1)

    # Add the pole vector to the driver IK handle if it's the rear leg.  If it's the front add it to the knee ik handle
    if isRearleg:
        cmds.poleVectorConstraint(kneeControlName, (limbName + "_Driver_IKHandle"), w=1)
    else:
        cmds.poleVectorConstraint(kneeControlName, (limbName + "_Knee_IKHandle"), w=1)

    #Add Hock Control
    if isRearleg:
        multiValue = 10
    else:
        multiValue = 15

    cmds.shadingNode("multiplyDivide", au=1, n=(limbName + "_Hock_Multi"))
    cmds.connectAttr((hockControlName + ".translate"), (limbName + "_Hock_Multi.input1"), f=1)
    cmds.connectAttr((limbName + "_Hock_Multi.outputZ"), (limbName + "_Knee_Ctrl.rotateZ"), f=1)
    cmds.connectAttr((limbName + "_Hock_Multi.outputY"), (limbName + "_Knee_Ctrl.rotateX"), f=1)

    if whichSide == "L_":
        cmds.setAttr((limbName + "_Hock_Multi.input2Z"), multiValue)
        cmds.setAttr((limbName + "_Hock_Multi.input2Y"), multiValue)
    else:
        cmds.setAttr((limbName + "_Hock_Multi.input2Z"), multiValue)
        cmds.setAttr((limbName + "_Hock_Multi.input2Y"), multiValue * -1)

    # -----------------------------------------------------------------------------------------------------------------
    # Add IK and FK Blending
    # -----------------------------------------------------------------------------------------------------------------

    for i in range(limbJoints):
        getConstraint = cmds.listConnections(jointHierarchy[i], type="parentConstraint")[0]
        getWeights = cmds.parentConstraint(getConstraint, q=1, wal=1)

        cmds.connectAttr((mainControl + ".FK_IK_Switch"), (getConstraint + "." + getWeights[0]), f=1)
        cmds.connectAttr((limbName + "_Rev.outputX"), (getConstraint + "." + getWeights[1]), f=1)

    # Add a group for the limb
    cmds.group(em=1, n=(limbName + "_Grp"))

    #Move it to the root position and freeze the transforms
    cmds.matchTransform((limbName + "_Grp"), jointRoot)
    cmds.makeIdentity((limbName + "_Grp"), a=1, t=1, r=1, s=0)

    #Parent the joints to the new group
    cmds.parent((jointRoot + "_IK"), (jointRoot + "_FK"), (jointRoot + "_Stretch"), (limbName + "_Grp"))

    if isRearleg:
        cmds.parent((jointRoot + "_Driver"), (limbName + "_Grp"))

    #Make the new group follow the root control
    cmds.parentConstraint(rootControlName, (limbName + "_Grp"), w=1, mo=1)

    #Move the group to the rig systems folder
    cmds.parent((limbName + "_Grp"), "Rig_Systems")

    #Clear the selection
    cmds.select(cl=1)

    #-----------------------------------------------------------------------------------------------------------------
    #Make Stretchy
    #-----------------------------------------------------------------------------------------------------------------

    if stretchCheck:

        #Create locator which dictates the end position
        cmds.spaceLocator(n=(limbName +"_StretchEndPos_Loc"))

        #Move it to the end joint
        cmds.matchTransform((limbName +"_StretchEndPos_Loc"), jointHierarchy[3])

        #Parent the locator to the paw control
        cmds.parent((limbName + "_StretchEndPos_Loc"), pawControlName)

        #Start to build the distance nodes
        #First, we will need to add all the distance nodes together, so we need a plusMinusAverage node
        cmds.shadingNode("plusMinusAverage", au=1, n=(limbName + "_Length"))

        for i in range(limbJoints):
            #Ignore the last joint or it will try to use the toes
            if i is not limbJoints - 1:
                cmds.shadingNode("distanceBetween", au=1, n=(jointHierarchy[i] + "_DistNode"))
                cmds.connectAttr((jointHierarchy[i] + "_Stretch.worldMatrix"), (jointHierarchy[i] + "_DistNode.inMatrix1"), f=1)
                cmds.connectAttr((jointHierarchy[i+1] + "_Stretch.worldMatrix"), (jointHierarchy[i] + "_DistNode.inMatrix2"), f=1)

                cmds.connectAttr((jointHierarchy[i] + "_Stretch.rotatePivotTranslate"), (jointHierarchy[i] + "_DistNode.point1"), f=1)
                cmds.connectAttr((jointHierarchy[i + 1] + "_Stretch.rotatePivotTranslate"), (jointHierarchy[i] + "_DistNode.point2"),f=1)

                cmds.connectAttr((jointHierarchy[i] + "_DistNode.distance"), (limbName + "_Length.input1D[" + str(i) + "]"),f=1)

        #Now get the distance from the root to the stretch end locator. We use this to check if the leg should stretch
        cmds.shadingNode("distanceBetween", au=1, n=(limbName + "_Stretch_DistNode"))

        cmds.connectAttr((jointHierarchy[0] + "_Stretch.worldMatrix"), (limbName + "_Stretch_DistNode.inMatrix1"), f=1)
        cmds.connectAttr((limbName +"_StretchEndPos_Loc.worldMatrix"), (limbName + "_Stretch_DistNode.inMatrix2"), f=1)

        cmds.connectAttr((jointHierarchy[0] + "_Stretch.rotatePivotTranslate"), (limbName + "_Stretch_DistNode.point1"),f=1)
        cmds.connectAttr((limbName +"_StretchEndPos_Loc.rotatePivotTranslate"),(limbName + "_Stretch_DistNode.point2"), f=1)

        #Create nodes to check for stretching, and to control how the stretch works

        #Scale factor compares the length of the leg with the stretch locator, so we can see when the leg is actually stretching
        cmds.shadingNode("multiplyDivide", au=1, n=(limbName + "_Scale_Factor"))

        cmds.shadingNode("condition", au=1, n=(limbName + "_Condition"))

        #Adjust the node Settings
        cmds.setAttr((limbName + "_Scale_Factor.operation"), 2)
        cmds.setAttr((limbName + "_Condition.operation"), 2)
        cmds.setAttr((limbName + "_Condition.secondTerm"), 1)

        #Connect the stretch distance to the scale factor multiply divide node

        cmds.connectAttr((limbName + "_Stretch_DistNode.distance"), (limbName + "_Scale_Factor.input1X"), f=1)

        #Connect the full leg distance to the scale factor multiply divide node

        cmds.connectAttr((limbName + "_Length.output1D"), (limbName + "_Scale_Factor.input2X"), f=1)

        #Next, connect the stretch factor node to the first term in the condition node

        cmds.connectAttr((limbName + "_Scale_Factor.outputX"), (limbName + "_Condition.firstTerm"), f=1)

        #Also, connect it to the color if true attribute, so we can use this as the stretch value
        cmds.connectAttr((limbName + "_Scale_Factor.outputX"), (limbName + "_Condition.colorIfTrueR"), f=1)

        #Now connect the stretch value to the ik and driver joints, so they stretch
        for i in range(limbJoints):
            cmds.connectAttr((limbName + "_Condition.outColorR"), (jointHierarchy[i] + "_IK.scaleX"), f=1)

            #Also affect the driver skeleton, if this is the rear leg
            if isRearleg:
                cmds.connectAttr((limbName + "_Condition.outColorR"), (jointHierarchy[i] + "_Driver.scaleX"), f=1)

        # Add the ability to turn the stretchiness off
        cmds.shadingNode("blendColors", au=1, n=(limbName + "_BlendColors"))
        cmds.setAttr((limbName + "_BlendColors.color2"), 1,0,0, type = "double3")

        cmds.connectAttr((limbName + "_Scale_Factor.outputX"), (limbName + "_BlendColors.color1R"), f=1)
        cmds.connectAttr((limbName + "_BlendColors.outputR"), (limbName + "_Condition.colorIfTrueR"), f=1)

        #Connect to the paw control attribute
        cmds.connectAttr((pawControlName + ".Stretchiness"), (limbName + "_BlendColors.blender"), f=1)

        #Wire up the attributes so we can control how the stretch works
        cmds.setAttr((pawControlName + ".Stretch_Type"), 0)
        cmds.setAttr((limbName + "_Condition.operation"), 1) # Not Equal

        cmds.setDrivenKeyframe((limbName + "_Condition.operation"), cd= (pawControlName + ".Stretch_Type"))

        cmds.setAttr((pawControlName + ".Stretch_Type"), 1)
        cmds.setAttr((limbName + "_Condition.operation"), 3)  # Greater Than

        cmds.setDrivenKeyframe((limbName + "_Condition.operation"), cd=(pawControlName + ".Stretch_Type"))

        cmds.setAttr((pawControlName + ".Stretch_Type"), 2)
        cmds.setAttr((limbName + "_Condition.operation"), 5)  # Less or Equal

        cmds.setDrivenKeyframe((limbName + "_Condition.operation"), cd=(pawControlName + ".Stretch_Type"))

        cmds.setAttr((pawControlName + ".Stretch_Type"), 1)

        #Clear the Selection
        cmds.select(cl=1)

        # -----------------------------------------------------------------------------------------------------------------
        # Volume Preservation
        # -----------------------------------------------------------------------------------------------------------------

        # Create the main multiply divide node which will calculate the volume
        cmds.shadingNode("multiplyDivide", au=1, n=(limbName + "_Volume"))

        #Set the operation to Power
        cmds.setAttr(limbName + "_Volume.operation", 3)

        #Connect the main stretch value to the volume node
        cmds.connectAttr((limbName + "_BlendColors.outputR"), (limbName + "_Volume.input1X"), f=1)

        #Connect the condition node so we can control scaling
        cmds.connectAttr((limbName + "_Volume.outputX"), (limbName + "_Condition.colorIfTrueG"), f=1)

        #Connect to the fibula joint
        cmds.connectAttr((limbName + "_Condition.outColorG"), (jointHierarchy[1] + ".scaleY"), f=1)
        cmds.connectAttr((limbName + "_Condition.outColorG"), (jointHierarchy[1] + ".scaleZ"), f=1)

        # Connect to the Metatarsus joint
        cmds.connectAttr((limbName + "_Condition.outColorG"), (jointHierarchy[2] + ".scaleY"), f=1)
        cmds.connectAttr((limbName + "_Condition.outColorG"), (jointHierarchy[2] + ".scaleZ"), f=1)

        #Connect to the main volume attribute
        cmds.connectAttr((mainControl + ".Volume_Offset"), (limbName + "_Volume.input2X"), f=1)

    # -----------------------------------------------------------------------------------------------------------------
    # Add Roll Joints and Systems
    # -----------------------------------------------------------------------------------------------------------------

    if rollCheck:

        #Upper Leg Systems

        #Check which side we are working on so we can move things to the correct side
        if whichSide == "L_":
            flipSide = 1
        else:
            flipSide =-1


        #Create the main roll and follow joints
        rollJointList = [jointHierarchy[0], jointHierarchy[3], jointHierarchy[0], jointHierarchy[0]]

        for i in range(len(rollJointList)):
            #Setup the Joint Names
            if i >2:
                rollJointName = rollJointList[i] + "_Follow_Tip"
            elif i >1:
                rollJointName = rollJointList[i] + "_Follow"
            else:
                rollJointName = rollJointList[i] + "_Roll"

            cmds.joint(n=rollJointName, rad=3)
            cmds.matchTransform(rollJointName, rollJointList[i])
            cmds.makeIdentity(rollJointName, a=1, t=0, r=1, s=0)

            if i <2:
                cmds.parent(rollJointName, rollJointList[i])
            elif i >2:
                cmds.parent(rollJointName, rollJointList[2] + "_Follow")

            cmds.select(cl=1)

            #Show the rotational axes to help us visualize the rotations
            #cmds.toggle(rollJointName, la=1)

        #Let's work on the femur first and adjust the follow joints
        cmds.pointConstraint(jointHierarchy[0], jointHierarchy[1], rollJointList[2] + "_Follow_Tip", w=1, mo=0, n="tempPC")
        cmds.delete("tempPC")

        #Now move them out
        cmds.move(0,0,-5 * flipSide, (rollJointList[2] + "_Follow"), r=1, os=1, wd=1)

        #Create the aim locator which the femur roll joint will always follow
        cmds.spaceLocator(n=(rollJointList[0] + "_Roll_Aim"))
        cmds.matchTransform((rollJointList[0] + "_Roll_Aim"), (rollJointList[2] + "_Follow"))
        cmds.parent((rollJointList[0] + "_Roll_Aim"), (rollJointList[2] + "_Follow"))

        #Move the locator out too
        cmds.move(0,0,-5 * flipSide, (rollJointList[0] + "_Roll_Aim"), r=1, os=1, wd=1)

        #Make the roll joint aim at the fibula joint,but also keep looking at the aim locator for reference
        cmds.aimConstraint(jointHierarchy[1], (rollJointList[0] + "_Roll"), w=1, aim=(1,0,0), u=(0,0,-1), wut="object", wuo=rollJointList[0] + "_Roll_Aim", mo=1)

        #Add the IK handle so the follow joints follow the leg
        cmds.ikHandle(n=(limbName + "_Follow_IKHandle"), sol="ikRPsolver", sj=(rollJointList[2] + "_Follow"), ee=rollJointList[2] + "_Follow_Tip")

        #Now move it to the fibula and parent it too
        cmds.parent((limbName + "_Follow_IKHandle"), (jointHierarchy[1]))
        cmds.matchTransform((limbName + "_Follow_IKHandle"), (jointHierarchy[1]))

        #Reset the pole vector
        cmds.setAttr(limbName + "_Follow_IKHandle.poleVectorX", 0)
        cmds.setAttr(limbName + "_Follow_IKHandle.poleVectorY", 0)
        cmds.setAttr(limbName + "_Follow_IKHandle.poleVectorZ", 0)

        #Lower Leg Systems
        # Create the aim locator which the Metacarpus roll joint will always follow
        cmds.spaceLocator(n=(rollJointList[1] + "_Roll_Aim"))

        #move it to the ankle joint and parent it to ankle joint too
        cmds.matchTransform((rollJointList[1] + "_Roll_Aim"), (rollJointList[1] + "_Roll"))
        cmds.parent((rollJointList[1] + "_Roll_Aim"), jointHierarchy[3])

        #Move the locator out
        cmds.move(5 * flipSide,0,0, (rollJointList[1] + "_Roll_Aim"), r=1, os=1, wd=1)

        #Make the ankle joint aim at the Metatarsus joint,but also keep looking at the aim locator for reference
        cmds.aimConstraint(jointHierarchy[2], (rollJointList[1] + "_Roll"), w=1, aim=(0,1,0), u=(1,0,0), wut="object", wuo=rollJointList[1] + "_Roll_Aim", mo=1)

        #Update the hierarchy, parenting the follow joints to the main group
        cmds.parent((rollJointList[0] + "_Follow"), (limbName + "_Grp"))

        #Clear the Selection
        cmds.select(cl=1)


def autoLimbToolUI():

    # First we check if the window exists and if it does, delete it
    if cmds.window("AutoLimbToolUI", ex=1):
        cmds.deleteUI("AutoLimbToolUI")

    #Create the window
    window = cmds.window("AutoLimbToolUI", t="Auto Limb Tool v1.0", w=200, h=200, mnb=0, mxb=0)

    #Create the main layour
    mainLayout = cmds.formLayout(nd=100)

    #Leg Menu
    legMenu = cmds.optionMenu("LegMenu", l = "Which Leg?", h=20, ann="Which side are we working on?")

    cmds.menuItem(l="Front")
    cmds.menuItem(l="Rear")

    #Checkboxes
    rollCheck = cmds.checkBox("RollCheck", l="Roll Joints", h=20, ann="Add Roll Joints?", v=0)
    stretchCheck = cmds.checkBox("StretchCheck", l="Stretchy", h=20, ann="Add Stretchiness?", v=0)


    #Separators
    separator01 = cmds.separator(h=5)
    separator02 = cmds.separator(h=5)

    #Buttons
    button = cmds.button(l="[Rig Leg]", c= autoLimbTool)

    #Adjust Layout
    cmds.formLayout(mainLayout, e=1,
                    attachForm = [(legMenu, 'top', 5), (legMenu, 'left', 5), (legMenu, 'right', 5),
                                  (separator01, 'left', 5), (separator01, 'right', 5),
                                  (separator02, 'left', 5), (separator02, 'right', 5),
                                  (button, 'bottom', 5), (button, 'left', 5), (button, 'right', 5)
                                  ],

                    attachControl = [(separator01, 'top', 5, legMenu),
                                     (rollCheck, 'top', 5, separator01),
                                     (stretchCheck, 'top', 5, separator01),

                                     (separator02, 'top', 5, rollCheck),
                                     (button, 'top', 5, separator02),
                                    ],
                    attachPosition = [(rollCheck, 'left', 0, 15),
                                      (stretchCheck, 'right', 0, 85)
                                     ]

    )



    #Show the window
    cmds.showWindow(window)
