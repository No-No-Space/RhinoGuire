#! python3
# -*- coding: utf-8 -*-
"""Terrain sampling for TerrainTools (see DECISIONS D4).

Wraps a Surface/Mesh/SubD/Polysurface as a sampled mesh with a vertical-raycast
``project_z(x, y)`` elevation probe. The coercion + projector logic is the same
technique proven in ``MeshTools/WrapeMeshOnMesh/Sebucan.py``.

Imports RhinoCommon, so this module only loads inside Rhino.
"""

import math

import rhinoscriptsyntax as rs
import scriptcontext as sc
import Rhino
import Rhino.Geometry as rg


# ---------------------------------------------------------------------------
# Mesh coercion (ported from Sebucan.py — behaviourally identical)
# ---------------------------------------------------------------------------

def _brep_to_mesh(brep):
    """Join all mesh faces produced from a BRep into a single Mesh."""
    mesh_list = rg.Mesh.CreateFromBrep(brep, rg.MeshingParameters.Default)
    if not mesh_list:
        return None
    joined = rg.Mesh()
    for m in mesh_list:
        joined.Append(m)
    return joined


def coerce_to_mesh(obj):
    """Convert a Mesh, SubD, Surface, Trimmed Surface, or Polysurface to a Mesh.

    *obj* may be a Rhino object id (GUID) or a geometry instance.
    Returns the Mesh, or None if conversion fails.
    """
    if isinstance(obj, rg.Mesh):
        return obj

    mesh = rs.coercemesh(obj)
    if mesh is not None:
        return mesh

    brep = rs.coercebrep(obj)
    if brep is not None:
        return _brep_to_mesh(brep)

    geo = rs.coercegeometry(obj)
    if isinstance(geo, rg.SubD):
        b = geo.ToBrep()
        if b is not None:
            return _brep_to_mesh(b)
    if isinstance(geo, rg.Mesh):
        return geo
    if isinstance(geo, rg.Brep):
        return _brep_to_mesh(geo)

    return None


def obj_type_label(obj_id):
    """Return a short display string for the object type ('Mesh', 'Surface'...)."""
    if rs.IsMesh(obj_id):
        return "Mesh"
    geo = rs.coercegeometry(obj_id)
    if isinstance(geo, rg.SubD):
        return "SubD"
    if rs.IsPolysurface(obj_id):
        return "Polysurface"
    if rs.IsSurface(obj_id):
        return "Surface"
    return "Object"


def _make_projector(dest_mesh):
    """Return a ``project_z(x, y)`` closure that raycasts onto *dest_mesh*.

    Results are cached by (x, y). Returns the Z float at that plan position,
    or None if (x, y) falls outside the mesh footprint.
    """
    bbox     = dest_mesh.GetBoundingBox(True)
    z_range  = max(bbox.Max.Z - bbox.Min.Z, 1.0)
    z_top    = bbox.Max.Z + z_range + 1.0
    z_bottom = bbox.Min.Z - z_range - 1.0
    _cache   = {}

    def project_z(x, y):
        key = (round(x, 8), round(y, 8))
        if key in _cache:
            return _cache[key]
        ray = rg.Ray3d(rg.Point3d(x, y, z_top), rg.Vector3d(0, 0, -1))
        t = rg.Intersect.Intersection.MeshRay(dest_mesh, ray)
        if t >= 0.0:
            z = float(ray.PointAt(t).Z)
            _cache[key] = z
            return z
        ray = rg.Ray3d(rg.Point3d(x, y, z_bottom), rg.Vector3d(0, 0, 1))
        t = rg.Intersect.Intersection.MeshRay(dest_mesh, ray)
        if t >= 0.0:
            z = float(ray.PointAt(t).Z)
            _cache[key] = z
            return z
        _cache[key] = None
        return None

    return project_z


# ---------------------------------------------------------------------------
# TerrainModel
# ---------------------------------------------------------------------------

class TerrainModel:
    """A terrain surface coerced to a mesh with a cached vertical-raycast probe."""

    def __init__(self, obj):
        self._mesh = coerce_to_mesh(obj)
        if self._mesh is None:
            raise ValueError("Could not coerce the selected object to a terrain mesh.")
        self._bbox = self._mesh.GetBoundingBox(True)
        self._project = _make_projector(self._mesh)

    @property
    def mesh(self):
        return self._mesh

    @property
    def bbox(self):
        return self._bbox

    def project_z(self, x, y):
        """Elevation directly below/above (x, y), or None if outside footprint."""
        return self._project(x, y)

    def sample_grid(self, x0, y0, nx, ny, cell):
        """Return a 2D list [j][i] of Z (or None) over a regular grid."""
        grid = []
        for j in range(ny):
            y = y0 + j * cell
            row = []
            for i in range(nx):
                row.append(self._project(x0 + i * cell, y))
            grid.append(row)
        return grid


# ---------------------------------------------------------------------------
# Free helpers
# ---------------------------------------------------------------------------

def _bbox_of(geometry):
    """BoundingBox of a BoundingBox, a curve, a Point3d list, or a GeometryBase."""
    if isinstance(geometry, rg.BoundingBox):
        return geometry
    if isinstance(geometry, (list, tuple)):
        bb = rg.BoundingBox.Empty
        for p in geometry:
            bb.Union(p if isinstance(p, rg.Point3d) else rg.Point3d(p))
        return bb
    return geometry.GetBoundingBox(True)


def make_grid_bounds(geometry, reach, cell):
    """Grid extents = feature bbox expanded by *reach*, snapped to *cell*.

    Returns (x0, y0, nx, ny, cell). x0/y0 are the lower-left node; nx/ny are the
    node counts (so the grid spans (nx-1)*cell by (ny-1)*cell).
    """
    bb = _bbox_of(geometry)
    x0 = bb.Min.X - reach
    y0 = bb.Min.Y - reach
    x1 = bb.Max.X + reach
    y1 = bb.Max.Y + reach
    nx = int(math.ceil((x1 - x0) / cell)) + 1
    ny = int(math.ceil((y1 - y0) / cell)) + 1
    return x0, y0, max(nx, 2), max(ny, 2), cell


def point_in_curve(curve, pt, tol):
    """True if *pt* is inside the closed planar *curve* (tested on World XY)."""
    try:
        res = curve.Contains(pt, rg.Plane.WorldXY, tol)
    except TypeError:
        res = curve.Contains(pt)
    return res == rg.PointContainment.Inside


def dist_and_edge_z(curve, pt):
    """Horizontal distance from *pt* to *curve* and the Z of the closest point.

    Returns (dist_xy, z_edge, closest_point) or (None, None, None) on failure.
    """
    ok, t = curve.ClosestPoint(pt)
    if not ok:
        return None, None, None
    cp = curve.PointAt(t)
    d = math.hypot(pt.X - cp.X, pt.Y - cp.Y)
    return d, float(cp.Z), cp


def is_closed_planar(curve, tol=None):
    """Return (closed, planar) booleans for a curve."""
    if tol is None:
        tol = sc.doc.ModelAbsoluteTolerance if sc.doc else 1e-6
    closed = bool(curve.IsClosed)
    planar = bool(curve.IsPlanar(tol))
    return closed, planar


def model_unit_label():
    """Abbreviated model-unit name, e.g. 'm', 'mm', 'ft'. Empty on failure."""
    try:
        return sc.doc.GetUnitSystemName(False, False, True, True)
    except Exception:
        return ""
