#! python3
# -*- coding: utf-8 -*-
"""Defaults and tag keys for Trocha (see road_tool_plan.md S5, S7).

Pure Python - no RhinoCommon, no Eto - so this loads and is testable outside
Rhino. Geometry tolerances are derived from ``doc.ModelAbsoluteTolerance`` at
the call site (Trocha.py passes it in); nothing here is hardcoded per-unit.
"""


class TrochaConfig(object):
    """Geometry-pipeline knobs for one build_slab() call (road_tool_plan.md S7).

    fit_tol / top_rise / margin default off *sample_step* when not given
    explicitly - sample_step is a physical, document-scaled length the user
    already sets in real model units (like width/thickness), unlike
    ModelAbsoluteTolerance, which is a numerical-precision setting that can
    legitimately be loose (0.1+) on a large site model without meaning
    "the road should deviate/float by 10s of centimeters" (bug report
    2026-07-15: tolerance-scaled defaults inflated thickness 0.1 -> 1.316m
    and floated the slab off the terrain on a loose-tolerance file).
    *tolerance* is kept only as a tiny floor (avoids degenerate near-zero
    values) and for its real remaining use: the geometric-coincidence `tol`
    argument passed into the Brep offset/Boolean operations in geometry.py.
    """

    def __init__(self, tolerance, sample_step=None, fit_tol=None, top_rise=None,
                 margin=None, smooth_center=True, corner_radius=None, max_drift=None):
        self.tolerance = tolerance
        self.sample_step = sample_step if sample_step is not None else DEFAULT_SAMPLE_STEP
        # Floor multiplier is deliberately gentle (0.1x, not e.g. 2x): it must
        # only catch a degenerate near-zero sample_step, not re-dominate the
        # sample_step-based default at realistic "loose" tolerances (~0.1,
        # exactly the value that caused the 2026-07-15 bug) - a steeper
        # multiplier reintroduces the same inflation at a smaller scale.
        self.fit_tol = fit_tol if fit_tol is not None else max(self.sample_step * 0.05, tolerance * 0.1)
        self.top_rise = top_rise if top_rise is not None else max(self.sample_step * 0.02, tolerance * 0.1)
        self.margin = margin if margin is not None else max(self.sample_step * 0.02, tolerance * 0.1)
        self.smooth_center = smooth_center
        self.corner_radius = corner_radius
        # How far the smoothed centerline may stray (in plan) from the drawn
        # one. None -> resolved per road in build_slab() as
        # MAX_DRIFT_WIDTH_FACTOR * width, since what counts as "off the road"
        # scales with the road (bug 2026-09-30: an unchecked Rebuild moved
        # centerlines up to 3.65m sideways, off the road bench on a hillside).
        self.max_drift = max_drift

    def resolved_max_drift(self, width):
        if self.max_drift is not None:
            return self.max_drift
        return max(width * MAX_DRIFT_WIDTH_FACTOR, self.tolerance)


DEFAULT_WIDTH = 4.0
DEFAULT_THICKNESS = 0.25
DEFAULT_SAMPLE_STEP = 2.0
MAX_DRIFT_WIDTH_FACTOR = 0.05  # smoothed centerline may drift <= 5% of road width
DEFAULT_LAYER = "RoadTool::Roads"
MERGED_LAYER = "RoadTool::Roads::Merged"
DISPLAY_MODE_NAME = "RG_Technical_Colour_NoEdges"

# User-string tag keys (road_tool_plan.md S5). The RG_ prefix matches the
# suite's launch-key namespace but is a different system (per-object user
# strings vs. launch.py dict keys) - no collision.
TAG_ROAD = "RG_ROAD"
TAG_WIDTH = "RG_ROAD_WIDTH"
TAG_THICK = "RG_ROAD_THICK"
TAG_TERRAIN = "RG_ROAD_TERRAIN"
TAG_CHILD = "RG_ROAD_CHILD"
TAG_PARENT = "RG_ROAD_PARENT"
TAG_MERGED = "RG_ROAD_MERGED"
TAG_MERGE_SOURCES = "RG_ROAD_MERGE_SOURCES"  # "|"-joined centerline GUIDs a merged solid was built from
