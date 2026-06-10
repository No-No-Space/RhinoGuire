#! python3
# -*- coding: utf-8 -*-
"""Design-heightfield builders for TerrainTools (see DECISIONS D5).

Both pads and ways are reduced to a *design Z grid* over a shared regular XY
grid, alongside the existing-ground Z grid. The per-node design elevation is
computed analytically with a slope clamp that auto-stops at the daylight line.

Imports RhinoCommon (via terrain.py), so this only loads inside Rhino.
``GradeResult`` itself is plain data — volumes.py consumes it by duck typing
without importing this module, keeping volumes RhinoCommon-free.
"""

import math

import Rhino.Geometry as rg

from TerrainTools._core import terrain as _terrain


# ---------------------------------------------------------------------------
# Result container (plain data — no RhinoCommon needed to read it)
# ---------------------------------------------------------------------------

class GradeResult(object):
    """A design + terrain heightfield over a regular grid.

    Attributes:
      x0, y0, cell, nx, ny  : grid origin (lower-left node), spacing, node counts
      z_design  [j][i]      : design elevation (float or None)
      z_terrain [j][i]      : existing-ground elevation (float or None)
      region_mask[j][i]     : True where the node is inside a feature or its
                              active grading skirt (drives meshing + volumes)
      flags                 : list[str] of human-readable warnings
      kind                  : 'pad' or 'corridor'
      stations              : list[dict] of per-station areas (corridors only)
      station_spacing       : along-centerline sample step (corridors only)
    """

    def __init__(self, x0, y0, cell, nx, ny, kind="pad"):
        self.x0, self.y0, self.cell = x0, y0, cell
        self.nx, self.ny = nx, ny
        self.kind = kind
        self.z_design   = [[None] * nx for _ in range(ny)]
        self.z_terrain  = [[None] * nx for _ in range(ny)]
        self.region_mask = [[False] * nx for _ in range(ny)]
        self.flags = []
        self.stations = []
        self.station_spacing = None

    def node_xy(self, i, j):
        return self.x0 + i * self.cell, self.y0 + j * self.cell


# ---------------------------------------------------------------------------
# Core per-node slope clamp (D5)
# ---------------------------------------------------------------------------

_FLAT_TOL = 1e-9


def design_z(z_t, z_edge, d, m_cut, m_fill):
    """Design elevation at a skirt node (D5 rule).

    z_t    : terrain Z at the node (must not be None here)
    z_edge : feature elevation at the nearest edge point
    d      : horizontal distance from node to feature edge (>= 0)
    m_cut, m_fill : positive gradients (V/H); <= 0 means 'no skirt that side'.
    """
    if z_t > z_edge + _FLAT_TOL:
        if m_cut <= 0.0:
            return z_t            # no cut skirt -> leave ground (vertical edge)
        return min(z_t, z_edge + d * m_cut)
    if z_t < z_edge - _FLAT_TOL:
        if m_fill <= 0.0:
            return z_t            # no fill skirt
        return max(z_t, z_edge - d * m_fill)
    return z_t


# ---------------------------------------------------------------------------
# Pads
# ---------------------------------------------------------------------------

def _pad_elevation(terr, curve, mode, cell):
    """Resolve a pad's platform elevation.

    mode: a float/int  -> used verbatim as the explicit Z.
          'mean'|'min'|'max' -> statistic of terrain sampled inside the curve.
    Returns (pad_z, n_samples). pad_z is None if no terrain falls inside.
    """
    if isinstance(mode, (int, float)):
        return float(mode), 0

    bb = curve.GetBoundingBox(True)
    nx = max(int(math.ceil((bb.Max.X - bb.Min.X) / cell)) + 1, 2)
    ny = max(int(math.ceil((bb.Max.Y - bb.Min.Y) / cell)) + 1, 2)
    tol = 1e-6
    zs = []
    for j in range(ny):
        y = bb.Min.Y + j * cell
        for i in range(nx):
            x = bb.Min.X + i * cell
            pt = rg.Point3d(x, y, 0.0)
            if _terrain.point_in_curve(curve, pt, tol):
                z = terr.project_z(x, y)
                if z is not None:
                    zs.append(z)
    if not zs:
        return None, 0
    if mode == "min":
        return min(zs), len(zs)
    if mode == "max":
        return max(zs), len(zs)
    return sum(zs) / len(zs), len(zs)   # mean (default)


def _auto_reach(terr, curves, pad_zs, m_cut, m_fill, max_reach):
    """Estimate the grading reach from the worst terrain-vs-pad elevation diff."""
    m_min = min([m for m in (m_cut, m_fill) if m > 0.0] or [1.0])
    dz_max = 0.0
    for curve, pad_z in zip(curves, pad_zs):
        if pad_z is None:
            continue
        bb = curve.GetBoundingBox(True)
        for corner in bb.GetCorners():
            z = terr.project_z(corner.X, corner.Y)
            if z is not None:
                dz_max = max(dz_max, abs(z - pad_z))
        zc = terr.project_z((bb.Min.X + bb.Max.X) / 2.0, (bb.Min.Y + bb.Max.Y) / 2.0)
        if zc is not None:
            dz_max = max(dz_max, abs(zc - pad_z))
    if dz_max <= 0.0:
        dz_max = max(bb.Max.Z - bb.Min.Z, 1.0)
    reach = dz_max / m_min * 1.15 + 2.0 * 0.0   # 15% headroom past first estimate
    return min(reach, max_reach)


def grade_pads(terr, boundary_curves, pad_z_mode, m_cut, m_fill, cell, max_reach):
    """Build a GradeResult for one or more building pads.

    boundary_curves : list of closed planar curves.
    pad_z_mode      : explicit float, or 'mean'|'min'|'max' (per-pad statistic).
    m_cut, m_fill   : canonical gradients (V/H, > 0).
    cell            : grid spacing (model units).
    max_reach       : cap on how far the skirt grid extends past each pad.

    Merge rule for overlapping skirts: nearest-feature wins (a node's skirt is
    governed by the pad whose edge is closest). Documented in DECISIONS.
    """
    if not boundary_curves:
        raise ValueError("grade_pads needs at least one boundary curve.")

    pad_zs = [_pad_elevation(terr, c, pad_z_mode, cell)[0] for c in boundary_curves]
    if all(z is None for z in pad_zs):
        raise ValueError("No terrain found under the pad boundary — cannot resolve elevation.")
    # Fall back to the mean of resolved pads for any that returned None.
    valid = [z for z in pad_zs if z is not None]
    fallback = sum(valid) / len(valid)
    pad_zs = [z if z is not None else fallback for z in pad_zs]

    reach = _auto_reach(terr, boundary_curves, pad_zs, m_cut, m_fill, max_reach)
    x0, y0, nx, ny, cell = _terrain.make_grid_bounds(boundary_curves[0]
                                                     if len(boundary_curves) == 1
                                                     else _union_box(boundary_curves),
                                                     reach, cell)

    res = GradeResult(x0, y0, cell, nx, ny, kind="pad")
    tol = 1e-6
    off_terrain = 0
    edge_on_slope = 0

    for j in range(ny):
        y = y0 + j * cell
        for i in range(nx):
            x = x0 + i * cell
            pt = rg.Point3d(x, y, 0.0)
            z_t = terr.project_z(x, y)
            res.z_terrain[j][i] = z_t

            # Inside any pad?
            inside_z = None
            for c, pz in zip(boundary_curves, pad_zs):
                if _terrain.point_in_curve(c, pt, tol):
                    inside_z = pz
                    break

            if inside_z is not None:
                res.z_design[j][i] = inside_z
                res.region_mask[j][i] = True
                if z_t is None:
                    off_terrain += 1
                continue

            # Skirt: nearest pad governs.
            best_d, best_edge = None, None
            for c, pz in zip(boundary_curves, pad_zs):
                d, _z_curve, _cp = _terrain.dist_and_edge_z(c, pt)
                if d is None:
                    continue
                if best_d is None or d < best_d:
                    best_d, best_edge = d, pz
            if best_d is None:
                res.z_design[j][i] = z_t
                continue

            if z_t is None:
                res.z_design[j][i] = None      # outside terrain footprint
                off_terrain += 1
                continue

            dz = design_z(z_t, best_edge, best_d, m_cut, m_fill)
            res.z_design[j][i] = dz
            if abs(dz - z_t) > 1e-6:
                res.region_mask[j][i] = True
                if i == 0 or j == 0 or i == nx - 1 or j == ny - 1:
                    edge_on_slope += 1

    if off_terrain:
        res.flags.append("%d grid node(s) fall outside the terrain footprint — "
                         "volumes there are excluded." % off_terrain)
    if edge_on_slope:
        res.flags.append("Grading still on a slope at the grid edge in %d place(s) — "
                         "increase Max reach for a clean daylight line." % edge_on_slope)
    return res


def _union_box(curves):
    bb = rg.BoundingBox.Empty
    for c in curves:
        bb.Union(c.GetBoundingBox(True))
    return bb


def grade_pad(terr, boundary_curve, pad_z, m_cut, m_fill, cell, max_reach):
    """Single-pad convenience wrapper around :func:`grade_pads`."""
    return grade_pads(terr, [boundary_curve], pad_z, m_cut, m_fill, cell, max_reach)


# ---------------------------------------------------------------------------
# Ways / corridors
# ---------------------------------------------------------------------------

class WayParams(object):
    """Parameter bundle for a way corridor.

    width      : full carriageway width (model units)
    m_cross    : crossfall gradient (V/H); 0 = flat carriageway
    crossfall  : 'crown' (falls both ways from centre) or 'single' (L->R)
    m_cut      : cut skirt gradient (V/H)
    m_fill     : fill skirt gradient (V/H)
    cell       : grid spacing
    station    : along-centerline sample step (defaults to cell if None)
    max_reach  : lateral cap past the carriageway edge
    """

    def __init__(self, width=4.0, m_cross=0.02, crossfall="crown",
                 m_cut=0.5, m_fill=0.5, cell=1.0, station=None, max_reach=30.0):
        self.width = width
        self.m_cross = m_cross
        self.crossfall = crossfall
        self.m_cut = m_cut
        self.m_fill = m_fill
        self.cell = cell
        self.station = station if station else cell
        self.max_reach = max_reach


def _carriageway_z(centerline_z, s, half_w, m_cross, crossfall):
    """Design Z on the carriageway at signed lateral offset *s* (|s| <= half_w)."""
    if crossfall == "single":
        return centerline_z - s * m_cross
    return centerline_z - abs(s) * m_cross   # crown


def grade_corridor(terr, centerline, params):
    """Build a GradeResult heightfield for a way corridor.

    Strategy: rasterize the corridor onto the same regular grid used by pads so
    volumes/CutFillReport work uniformly. For each grid node we find the nearest
    point on the centerline, derive the signed lateral offset, and apply the
    carriageway crossfall inside the width or the D5 skirt clamp outside it.
    Per-station cross-section areas are also accumulated for a mass-haul curve.

    Centerline profile v1 = 'drape on terrain' (centerline_z = project_z).
    """
    p = params
    half_w = params.width / 2.0

    reach = params.max_reach
    x0, y0, nx, ny, cell = _terrain.make_grid_bounds(centerline, half_w + reach, p.cell)
    res = GradeResult(x0, y0, cell, nx, ny, kind="corridor")
    res.station_spacing = p.station

    off_terrain = 0
    edge_on_slope = 0
    no_drape = 0

    # Cache centerline draped Z by curve parameter to avoid re-projecting.
    for j in range(ny):
        y = y0 + j * cell
        for i in range(nx):
            x = x0 + i * cell
            node = rg.Point3d(x, y, 0.0)
            z_t = terr.project_z(x, y)
            res.z_terrain[j][i] = z_t

            ok, t = centerline.ClosestPoint(node)
            if not ok:
                continue
            cp = centerline.PointAt(t)
            cz = terr.project_z(cp.X, cp.Y)         # drape
            if cz is None:
                no_drape += 1
                continue
            tan = centerline.TangentAt(t)
            n = rg.Vector3d(-tan.Y, tan.X, 0.0)
            if not n.Unitize():
                continue
            s = (node.X - cp.X) * n.X + (node.Y - cp.Y) * n.Y   # signed offset

            if abs(s) <= half_w:
                res.z_design[j][i] = _carriageway_z(cz, s, half_w, p.m_cross, p.crossfall)
                res.region_mask[j][i] = True
                if z_t is None:
                    off_terrain += 1
                continue

            # Skirt outside the carriageway.
            edge_s = half_w if s > 0 else -half_w
            z_edge = _carriageway_z(cz, edge_s, half_w, p.m_cross, p.crossfall)
            d = abs(s) - half_w
            if z_t is None:
                res.z_design[j][i] = None
                off_terrain += 1
                continue
            dz = design_z(z_t, z_edge, d, p.m_cut, p.m_fill)
            res.z_design[j][i] = dz
            if abs(dz - z_t) > 1e-6:
                res.region_mask[j][i] = True
                if i == 0 or j == 0 or i == nx - 1 or j == ny - 1:
                    edge_on_slope += 1

    res.stations = _corridor_stations(terr, centerline, params)

    if no_drape:
        res.flags.append("Centerline leaves the terrain at %d sampled node(s) — "
                         "those columns are skipped." % no_drape)
    if off_terrain:
        res.flags.append("%d grid node(s) fall outside the terrain footprint — "
                         "volumes there are excluded." % off_terrain)
    if edge_on_slope:
        res.flags.append("Grading still on a slope at the grid edge in %d place(s) — "
                         "increase Max reach." % edge_on_slope)
    return res


def _corridor_stations(terr, centerline, params):
    """Per-station cut/fill cross-section areas (for the mass-haul curve)."""
    p = params
    half_w = p.width / 2.0
    length = centerline.GetLength()
    if length <= 0:
        return []
    n_st = max(int(math.ceil(length / p.station)) + 1, 2)
    cross_step = p.cell
    reach = p.max_reach
    stations = []

    for k in range(n_st):
        dist = min(k * p.station, length)
        ok, t = centerline.NormalizedLengthParameter(dist / length)
        if not ok:
            continue
        cp = centerline.PointAt(t)
        cz = terr.project_z(cp.X, cp.Y)
        if cz is None:
            stations.append({"station": dist, "cut_area": 0.0, "fill_area": 0.0})
            continue
        tan = centerline.TangentAt(t)
        n = rg.Vector3d(-tan.Y, tan.X, 0.0)
        if not n.Unitize():
            continue
        cut_a = 0.0
        fill_a = 0.0
        s = -(half_w + reach)
        s_end = half_w + reach
        while s <= s_end:
            qx = cp.X + s * n.X
            qy = cp.Y + s * n.Y
            z_t = terr.project_z(qx, qy)
            if z_t is not None:
                if abs(s) <= half_w:
                    zd = _carriageway_z(cz, s, half_w, p.m_cross, p.crossfall)
                else:
                    edge_s = half_w if s > 0 else -half_w
                    z_edge = _carriageway_z(cz, edge_s, half_w, p.m_cross, p.crossfall)
                    zd = design_z(z_t, z_edge, abs(s) - half_w, p.m_cut, p.m_fill)
                delta = zd - z_t
                if delta > 0:
                    fill_a += delta * cross_step
                elif delta < 0:
                    cut_a += -delta * cross_step
            s += cross_step
        stations.append({"station": dist, "cut_area": cut_a, "fill_area": fill_a})
    return stations
