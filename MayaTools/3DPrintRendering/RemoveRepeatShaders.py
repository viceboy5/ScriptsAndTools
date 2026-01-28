import maya.cmds as cmds

def get_core_name(name):
    """
    Strip common prefixes and namespaces to get the 'core' shader name.
    Example:
    thBeaver_GoldSunluSilk_SilkGoldDis -> GoldSunluSilk_SilkGoldDis
    NewRenderEnvironment:GoldSunluSilk_SilkGoldDis -> GoldSunluSilk_SilkGoldDis
    """
    # Remove namespace
    if ":" in name:
        name = name.split(":", 1)[1]
    # Remove prefix up to first underscore if present
    if "_" in name:
        parts = name.split("_")
        if len(parts) > 1:
            name = "_".join(parts[1:])
    return name


def replace_shaders_and_cleanup():
    """
    Replaces local shaders with referenced ones, deletes all non-referenced displacements,
    and removes unused local shaders and shading groups.
    """
    shading_engines = cmds.ls(type="shadingEngine")

    # Build lookup: core shader name -> referenced shading group
    referenced_sgs = {}
    for sg in shading_engines:
        if not cmds.referenceQuery(sg, isNodeReferenced=True):
            continue
        shaders = cmds.listConnections(sg + ".surfaceShader", source=True, destination=False)
        if not shaders:
            continue
        core_name = get_core_name(shaders[0])
        referenced_sgs[core_name] = sg

    if not referenced_sgs:
        cmds.warning("No referenced shaders found.")
        return

    reassigned = 0

    # Step 1: Replace local shaders
    for sg in shading_engines:
        if cmds.referenceQuery(sg, isNodeReferenced=True):
            continue

        members = cmds.sets(sg, query=True)
        if not members:
            continue

        shaders = cmds.listConnections(sg + ".surfaceShader", source=True, destination=False)
        if not shaders:
            continue

        core_name = get_core_name(shaders[0])

        if core_name not in referenced_sgs:
            continue

        new_sg = referenced_sgs[core_name]

        # Reassign geometry
        cmds.sets(members, edit=True, forceElement=new_sg)
        reassigned += 1

    print(f"Reassigned {reassigned} local shaders to referenced versions.")

    # Step 2: Remove all non-referenced displacement nodes
    removed_disp = 0
    for sg in cmds.ls(type="shadingEngine"):
        # Find any displacement nodes connected
        disp_nodes = cmds.listConnections(sg + ".displacementShader", source=True, destination=False) or []
        for disp in disp_nodes:
            if not cmds.referenceQuery(disp, isNodeReferenced=True):
                # Disconnect and delete
                try:
                    cmds.disconnectAttr(disp + ".outDisplacement", sg + ".displacementShader")
                except:
                    pass  # safe if already disconnected
                cmds.delete(disp)
                removed_disp += 1

    print(f"Removed {removed_disp} non-referenced displacement nodes.")

    # Step 3: Cleanup unused local shaders and shading groups
    deleted_sgs = 0
    deleted_shaders = 0

    for sg in cmds.ls(type="shadingEngine"):
        if cmds.referenceQuery(sg, isNodeReferenced=True):
            continue

        members = cmds.sets(sg, query=True)
        if members:
            continue

        shaders = cmds.listConnections(sg + ".surfaceShader", source=True, destination=False)

        cmds.delete(sg)
        deleted_sgs += 1

        if shaders:
            shader = shaders[0]
            if not cmds.referenceQuery(shader, isNodeReferenced=True):
                cmds.delete(shader)
                deleted_shaders += 1

    print(f"Deleted {deleted_sgs} unused local shading groups.")
    print(f"Deleted {deleted_shaders} unused local shaders.")


# =====================
# RUN
# =====================
replace_shaders_and_cleanup()
