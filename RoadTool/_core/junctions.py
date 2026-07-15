#! python3
# -*- coding: utf-8 -*-
"""Junction merge for Trocha (road_tool_plan.md S10).

Because roads are solids, crossing/overlapping road slabs merge with a plain
Boolean union into one continuous body. Kept separate from the per-road
create/update flow so individual slabs stay editable and update-linked; the
merged solid is a derived, re-runnable output.

Pure RhinoCommon in/out - no RhinoDoc access.
"""

import Rhino.Geometry as rg


def merge_slabs(breps, tol):
    """Boolean-union *breps* into as few solids as possible.

    Returns (merged_breps, ok). ``ok`` is False if the union call itself
    failed (invalid input, no result at all) - note that Boolean union on
    non-overlapping breps still succeeds and simply returns them unchanged,
    which is not treated as failure.
    """
    if not breps or len(breps) < 2:
        return list(breps), False
    result = rg.Brep.CreateBooleanUnion(breps, tol)
    if not result:
        return list(breps), False
    return list(result), True
