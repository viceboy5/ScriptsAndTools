#-----------------------------------------------
#
#antCGI Auto Limb Tool V1, created from part 21 antCGI video tutorial
#
#-----------------------------------------------

import maya.cmds as cmds

def autoLimbTool():
    # Setup the variables which could come from the UI

    # Is this the front or rear Leg
    isRearleg = 1

    # How many joints are we working with?
    limbJoints = 4

    # Use this information to start to generate the names we need
    if isRearleg:
        limbType = "Rear"
        print("Working on the REAR leg.")
    else:
        limbType = "front"
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

    #Setup FK

    # Connect the FK controls in to the new joints
    for i in range(limbJoints):
        cmds.parentConstraint((jointHierarchy[i] + "_FK_Ctrl"), (jointHierarchy[i] + "_FK"), w=1, mo=0)

    #Setup IK

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
    cmds.parent( (limbName + "_Knee_Ctrl_Grp"), (limbName + "_Hock_IKHandle"), pawControlName)

    #If it's the rear leg, adjust the hierarchy so the driver leg controls the ik handles
    if isRearleg:
        cmds.parent((limbName + "_Knee_Ctrl_Grp"), (jointHierarchy[2] + "_Driver"))
        cmds.parent((limbName + "_Hock_IKHandle"), (jointHierarchy[3] + "_Driver"))
        cmds.parent((limbName + "_Driver_IKHandle"), pawControlName)

    # Make the paw control drive the ankle joint to maintain it's orientation
    cmds.orientConstraint( pawControlName, (jointHierarchy[3] + "_IK"), w=1)

    # Add the pole vector to the driver IK handle if it's the rear leg.  If it's the front add it to the knee ik handle
    if isRearleg:
        cmds.poleVectorConstraint(kneeControlName, (limbName + "_Driver_IKHandle"), w=1)
    else:
        cmds.poleVectorConstraint(kneeControlName, (limbName + "_Knee_IKHandle"), w=1)


