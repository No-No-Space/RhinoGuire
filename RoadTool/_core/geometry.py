#! python3
# -*- coding: utf-8 -*-
"""Pure geometry pipeline for Trocha - builds a road slab Brep from a
centerline, a terrain model, a width and a thickness.

RhinoCommon geometry in/out only - no RhinoDoc access (no ``sc.doc``, no
``rhinoscriptsyntax``). Takes an already-built ``TerrainTools._core.terrain
.TerrainModel`` instance rather than an object id, so it can be called with a
bare in-memory Mesh and exercised outside a live document. See
RoadTool/road_tool_plan.md S4/S4a for the design rationale.
"""

import math

import Rhino.Geometry as rg


class SlabResult(object):
    """Output of build_slab(): the Brep (or None on failure) plus warnings.

    ``brep`` may be a closed solid slab, or - if thickening failed even after
    retrying at a smaller thickness (road_tool_plan.md S8) - the bare top
    surface with a flag explaining why. Always check ``ok`` before trusting
    the result is a usable solid.
    """

    def __init__(self, brep=None, top_rise=None, thickness=None, flags=None):
        self.brep = brep
        self.top_rise = top_rise
        self.thickness = thickness
        self.flags = flags or []

    @property
    def ok(self):
        return self.brep is not None and bool(self.brep.IsSolid)


# ---------------------------------------------------------------------------
# Pipeline steps (road_tool_plan.md S4, numbered to match)
# ---------------------------------------------------------------------------

def _prep_centerline(curve, smooth_center):
    """S4.1 - optionally rebuild a kinked/polyline centerline to a smooth
    degree-3 curve. Returns *curve* unchanged if rebuilding isn't requested
    or fails.
    """
    if not smooth_center:
        return curve
    nc = curve.ToNurbsCurve()
    if nc is None:
        return curve
    try:
        point_count = max(nc.Points.Count, 10)
        rebuilt = nc.Rebuild(point_count, 3, True)
    except Exception:
        return curve
    return rebuilt or curve


def _stations(curve, step):
    """S4.2 - sample *curve* at ~*step* spacing. Returns (points, unit tangents)."""
    length = curve.GetLength()
    if length <= 0:
        return [], []
    count = max(int(math.ceil(length / max(step, 1e-6))), 2)
    params = curve.DivideByCount(count, True)
    if not params:
        params = [curve.Domain.T0, curve.Domain.T1]
    pts, tans = [], []
    for t in params:
        pt = curve.PointAt(t)
        tan = curve.TangentAt(t)
        if tan is None or not tan.IsValid or tan.Length < 1e-9:
            tan = rg.Vector3d(1.0, 0.0, 0.0)
        else:
            tan.Unitize()
        pts.append(pt)
        tans.append(tan)
    return pts, tans


def _offset_rails(pts, tans, width):
    """S4.3 - plan-offset Left/Right rail points from station points + tangents."""
    half = width / 2.0
    left, right = [], []
    for pt, tan in zip(pts, tans):
        perp = rg.Vector3d.CrossProduct(tan, rg.Vector3d.ZAxis)
        if perp.Length < 1e-9:
            perp = rg.Vector3d.XAxis
        else:
            perp.Unitize()
        left.append(pt + perp * half)
        right.append(pt - perp * half)
    return left, right


def _sample_terrain(points, terrain_model, z_raise):
    """S4.4 - project each point onto the terrain via TerrainModel.project_z,
    raised by *z_raise*. Returns (samples, n_missed); samples[i] is None on a
    miss (station outside the terrain footprint).
    """
    out = []
    missed = 0
    for pt in points:
        z = terrain_model.project_z(pt.X, pt.Y)
        if z is None:
            out.append(None)
            missed += 1
        else:
            out.append(rg.Point3d(pt.X, pt.Y, z + z_raise))
    return out, missed


def _fill_gaps(samples):
    """road_tool_plan.md S8 mitigation for terrain misses: clamp leading/
    trailing gaps to the nearest valid sample, linearly interpolate interior
    gaps between the two bounding valid samples.
    """
    n = len(samples)
    idxs = [i for i, s in enumerate(samples) if s is not None]
    if not idxs:
        return list(samples)
    filled = list(samples)
    first, last = idxs[0], idxs[-1]
    for i in range(0, first):
        filled[i] = samples[first]
    for i in range(last + 1, n):
        filled[i] = samples[last]
    for a, b in zip(idxs, idxs[1:]):
        if b - a <= 1:
            continue
        pa, pb = samples[a], samples[b]
        for i in range(a + 1, b):
            t = (i - a) / float(b - a)
            filled[i] = rg.Point3d(
                pa.X + (pb.X - pa.X) * t,
                pa.Y + (pb.Y - pa.Y) * t,
                pa.Z + (pb.Z - pa.Z) * t,
            )
    return filled


def _fit_edge(points, fit_tol):
    """S4.5 - fit a smooth degree-3 curve through *points*."""
    if len(points) < 2:
        return None
    crv = rg.Curve.CreateInterpolatedCurve(points, 3)
    if crv is None:
        return None
    try:
        fitted = crv.Fit(3, fit_tol, 0.0)
    except Exception:
        fitted = None
    return fitted or crv


def _max_deviation(fitted_curve, raw_points):
    """S4a - max distance from *fitted_curve* to any of *raw_points* (this is
    ``d`` in the contact rule).
    """
    d = 0.0
    for pt in raw_points:
        ok, t = fitted_curve.ClosestPoint(pt)
        if not ok:
            continue
        cp = fitted_curve.PointAt(t)
        dist = pt.DistanceTo(cp)
        if dist > d:
            d = dist
    return d


def _loft_top(left_curve, right_curve):
    """S4.6 - loft the two edge curves into a single crease-free top face."""
    breps = rg.Brep.CreateFromLoft(
        [left_curve, right_curve], rg.Point3d.Unset, rg.Point3d.Unset,
        rg.LoftType.Normal, False)
    if not breps:
        return None
    return breps[0]


def _thicken(top_brep, thickness, tol):
    """S4.7 - offset the top face into a closed solid slab, growing downward.

    Brep.CreateFromLoft does not guarantee which way the resulting face's
    normal points - it depends on which direction the centerline was drawn
    (CW vs CCW) - so a fixed offset sign can grow the solid *upward* instead,
    leaving the original top curve as the slab's underside (bug report
    2026-07-15: a uniform gap under the slab equal to top_rise). Try both
    directions and keep whichever solid actually extends downward.
    """
    if top_brep is None or top_brep.Faces.Count == 0:
        return None
    face = top_brep.Faces[0]

    def _try(offset):
        try:
            return rg.Brep.CreateFromOffsetFace(face, offset, tol, False, True)
        except Exception:
            return None

    down = _try(-thickness)
    up = _try(thickness)
    candidates = [b for b in (down, up) if b is not None and b.IsSolid]
    if not candidates:
        return down or up  # let build_slab's IsSolid check + retry loop handle failure
    return min(candidates, key=lambda b: b.GetBoundingBox(True).Min.Z)


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def build_slab(center_curve, terrain_model, width, thickness, cfg):
    """Build a road slab Brep draped on *terrain_model* from *center_curve*.

    Pure RhinoCommon in/out - no RhinoDoc access. Returns a SlabResult; check
    ``.ok`` before trusting ``.brep`` is a closed solid (see SlabResult docs
    for the top-surface-only fallback case, road_tool_plan.md S8).
    """
    tol = cfg.tolerance
    flags = []

    curve = _prep_centerline(center_curve, cfg.smooth_center) or center_curve

    pts, tans = _stations(curve, cfg.sample_step)
    if len(pts) < 2:
        return SlabResult(flags=["Centerline too short to station."])

    left_pts, right_pts = _offset_rails(pts, tans, width)

    # First pass at raw terrain Z (no raise) to measure the approximation
    # deviation d that the contact rule (S4a) is built on.
    left_raw, left_missed = _sample_terrain(left_pts, terrain_model, 0.0)
    right_raw, right_missed = _sample_terrain(right_pts, terrain_model, 0.0)
    if left_missed or right_missed:
        flags.append(
            "%d station(s) fell outside the terrain footprint - interpolated."
            % (left_missed + right_missed))
    left_raw = _fill_gaps(left_raw)
    right_raw = _fill_gaps(right_raw)
    if not left_raw or not right_raw:
        return SlabResult(flags=flags + ["Terrain sampling failed along the whole centerline."])

    left_fit0 = _fit_edge(left_raw, cfg.fit_tol)
    right_fit0 = _fit_edge(right_raw, cfg.fit_tol)
    if left_fit0 is None or right_fit0 is None:
        return SlabResult(flags=flags + ["Could not fit a smooth edge curve."])
    d = max(_max_deviation(left_fit0, left_raw), _max_deviation(right_fit0, right_raw))

    # S4a contact rule: top_rise >= d clears convex bumps; thickness buries
    # the underside through the worst dip.
    top_rise = max(cfg.top_rise, d)
    resolved_thickness = max(thickness, top_rise + d + cfg.margin)
    if resolved_thickness > thickness * 1.5:
        flags.append(
            "Thickness increased to %.4g to keep the road embedded (requested %.4g)."
            % (resolved_thickness, thickness))

    left_z, _ = _sample_terrain(left_pts, terrain_model, top_rise)
    right_z, _ = _sample_terrain(right_pts, terrain_model, top_rise)
    left_z = _fill_gaps(left_z)
    right_z = _fill_gaps(right_z)

    left_edge = _fit_edge(left_z, cfg.fit_tol)
    right_edge = _fit_edge(right_z, cfg.fit_tol)
    if left_edge is None or right_edge is None:
        return SlabResult(flags=flags + ["Could not fit the raised edge curve."])

    top = _loft_top(left_edge, right_edge)
    if top is None:
        return SlabResult(flags=flags + [
            "Loft failed - check for self-intersecting offsets (tight corners, S8)."])

    attempt_thickness = resolved_thickness
    slab = _thicken(top, attempt_thickness, tol)
    while (slab is None or not slab.IsSolid) and attempt_thickness > resolved_thickness * 0.25:
        attempt_thickness *= 0.5
        slab = _thicken(top, attempt_thickness, tol)

    if slab is None or not slab.IsSolid:
        flags.append(
            "Could not thicken to a closed solid - returning the top surface only; "
            "reduce thickness or width (S8).")
        return SlabResult(brep=top, top_rise=top_rise, thickness=resolved_thickness, flags=flags)

    if attempt_thickness != resolved_thickness:
        flags.append(
            "Thickness reduced to %.4g to avoid self-overlap (requested %.4g)."
            % (attempt_thickness, resolved_thickness))

    return SlabResult(brep=slab, top_rise=top_rise, thickness=attempt_thickness, flags=flags)
