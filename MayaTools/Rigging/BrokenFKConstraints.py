import maya.cmds as cmds
# created with professor Clayton Lantz
# make viewport selection, parent control and then child control


def create_broken_fk_constraints():
    # get selection, separate parent control and child control
    sels = cmds.ls(sl=True) # [parent control, child control]
    parent_ctrl = sels[0]
    child_ctrl = sels[1]

    # get the parent group of the child control
    child_ctrl_grp = cmds.listRelatives(child_ctrl, parent=True)[0]  #[child control's parent node]

    # create constraints
    p_constraint1 = cmds.parentConstraint( parent_ctrl, child_ctrl_grp, mo=True, skipRotate=['x', 'y', 'z'], weight =1)[0] #constrain translates
    p_constraint2 = cmds.parentConstraint(parent_ctrl, child_ctrl_grp, mo=True, skipTranslate=['x', 'y', 'z'], weight =1)[0] #constrain rotates
    cmds.scaleConstraint(parent_ctrl, child_ctrl_grp, weight =1)

    # create attributes on the child control
    if not cmds.attributeQuery('FollowTranslate', node=child_ctrl, exists =True):
        cmds.addAttr(child_ctrl, ln = 'FollowTranslate', at = 'double', min =0, max=1, dv=1)
        cmds.setAttr('%s.FollowTranslate' % (child_ctrl), e=True, keyable=True)
    if not cmds.attributeQuery('FollowRotate', node=child_ctrl, exists =True):
        cmds.addAttr(child_ctrl, ln = 'FollowRotate', at = 'double', min =0, max=1, dv=1)
        cmds.setAttr('%s.FollowRotate' % (child_ctrl), e=True, keyable=True)

    # connect attributes from child control to constraint weights
    cmds.connectAttr('%s.FollowTranslate' % child_ctrl, '%s.w0' % (p_constraint1), f=True)
    cmds.connectAttr('%s.FollowRotate' % child_ctrl, '%s.w0' % (p_constraint2), f=True)
