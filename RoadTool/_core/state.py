#! python3
# -*- coding: utf-8 -*-
"""Create/update state model for Trocha (road_tool_plan.md S5).

Reads/writes the RG_ROAD_* user strings that key a centerline to its
generated slab: self-healing by GUID scan when the stored child id goes
stale, plus an orphan-lookup helper for the re-link fallback when the
centerline itself was replaced (rev.2 review item #7).

Touches RhinoDoc (via rhinoscriptsyntax) - unlike geometry.py/junctions.py,
this module is document state, not pure geometry.
"""

import math

import rhinoscriptsyntax as rs

from RoadTool._core import config as _cfg


def is_road_centerline(obj_id):
    return rs.GetUserText(obj_id, _cfg.TAG_ROAD) == "1"


def read_params(center_id):
    """(width, thickness, terrain_id) stored on a tagged centerline.

    Returns (None, None, None) for any field that's missing/unparseable.
    """
    w = rs.GetUserText(center_id, _cfg.TAG_WIDTH)
    t = rs.GetUserText(center_id, _cfg.TAG_THICK)
    terrain_id = rs.GetUserText(center_id, _cfg.TAG_TERRAIN) or None
    try:
        width = float(w) if w else None
    except ValueError:
        width = None
    try:
        thickness = float(t) if t else None
    except ValueError:
        thickness = None
    return width, thickness, terrain_id


def write_centerline_tags(center_id, width, thickness, terrain_id, child_id):
    rs.SetUserText(center_id, _cfg.TAG_ROAD, "1")
    rs.SetUserText(center_id, _cfg.TAG_WIDTH, repr(width))
    rs.SetUserText(center_id, _cfg.TAG_THICK, repr(thickness))
    rs.SetUserText(center_id, _cfg.TAG_TERRAIN, str(terrain_id))
    rs.SetUserText(center_id, _cfg.TAG_CHILD, str(child_id))


def write_slab_tag(slab_id, center_id):
    rs.SetUserText(slab_id, _cfg.TAG_PARENT, str(center_id))


def child_of(center_id):
    """The slab GUID stored on *center_id*'s RG_ROAD_CHILD tag, or None."""
    return rs.GetUserText(center_id, _cfg.TAG_CHILD) or None


def resolve_child(center_id):
    """Find the live slab for *center_id*.

    Tries the stored RG_ROAD_CHILD GUID first; if it's stale (slab deleted
    by hand), falls back to scanning for any object whose RG_ROAD_PARENT
    equals this centerline's GUID (road_tool_plan.md S5 self-healing).
    """
    stored = child_of(center_id)
    if stored and rs.IsObject(stored):
        return stored
    center_str = str(center_id)
    for obj_id in rs.AllObjects():
        if rs.GetUserText(obj_id, _cfg.TAG_PARENT) == center_str:
            return obj_id
    return None


def find_orphan_near(curve):
    """Re-link fallback (road_tool_plan.md S5, rev.2 #7).

    Among slabs whose RG_ROAD_PARENT no longer resolves to a live
    centerline, return the id of the one whose bounding-box center is
    closest (in plan) to *curve*'s midpoint, or None if there are none.
    """
    best_id, best_dist = None, None
    mid = curve.PointAtNormalizedLength(0.5)
    for obj_id in rs.AllObjects():
        parent = rs.GetUserText(obj_id, _cfg.TAG_PARENT)
        if not parent or rs.IsObject(parent):
            continue  # not a slab, or its parent still exists (not an orphan)
        brep = rs.coercebrep(obj_id)
        if brep is None:
            continue
        center = brep.GetBoundingBox(True).Center
        dist = math.hypot(center.X - mid.X, center.Y - mid.Y)
        if best_dist is None or dist < best_dist:
            best_id, best_dist = obj_id, dist
    return best_id


def remove_road(center_id):
    """Delete the tagged slab (if any) and clear the RG_ROAD_* tags on *center_id*."""
    slab_id = resolve_child(center_id)
    if slab_id and rs.IsObject(slab_id):
        rs.DeleteObject(slab_id)
    for key in (_cfg.TAG_ROAD, _cfg.TAG_WIDTH, _cfg.TAG_THICK, _cfg.TAG_TERRAIN, _cfg.TAG_CHILD):
        rs.SetUserText(center_id, key, None)


def all_tagged_centerlines():
    return [oid for oid in rs.AllObjects() if is_road_centerline(oid)]
