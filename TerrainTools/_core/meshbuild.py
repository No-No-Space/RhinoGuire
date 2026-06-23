#! python3
# -*- coding: utf-8 -*-
"""Mesh construction + cut/fill tinting for TerrainTools.

Turns a GradeResult heightfield into a Rhino Mesh and tints it by per-vertex
cut/fill depth. Imports RhinoCommon + System.Drawing (System.Drawing is not Eto,
so _core stays UI-free). The colour *ramp* is injected by the caller as a
function returning System.Drawing.Color, so colour policy lives in the tools.
"""

import System.Drawing as sd
import Rhino.Geometry as rg


def grid_to_mesh(grade, only_region=True):
    """Quad grid -> triangulated Mesh built from z_design.

    Cells are skipped where any corner is None, or (only_region and the corner is
    outside region_mask). Returns the Mesh (may be empty).
    """
    nx, ny = grade.nx, grade.ny
    x0, y0, cell = grade.x0, grade.y0, grade.cell
    zd = grade.z_design
    mask = grade.region_mask

    vid = [[-1] * nx for _ in range(ny)]
    mesh = rg.Mesh()

    def usable(i, j):
        if zd[j][i] is None:
            return False
        if only_region and not mask[j][i]:
            return False
        return True

    for j in range(ny):
        for i in range(nx):
            if usable(i, j):
                vid[j][i] = mesh.Vertices.Add(x0 + i * cell, y0 + j * cell, zd[j][i])

    for j in range(ny - 1):
        for i in range(nx - 1):
            a = vid[j][i]
            b = vid[j][i + 1]
            c = vid[j + 1][i + 1]
            d = vid[j + 1][i]
            if a >= 0 and b >= 0 and c >= 0 and d >= 0:
                mesh.Faces.AddFace(a, b, c)
                mesh.Faces.AddFace(a, c, d)
            elif a >= 0 and b >= 0 and c >= 0:
                mesh.Faces.AddFace(a, b, c)
            elif a >= 0 and c >= 0 and d >= 0:
                mesh.Faces.AddFace(a, c, d)
            elif a >= 0 and b >= 0 and d >= 0:
                mesh.Faces.AddFace(a, b, d)
            elif b >= 0 and c >= 0 and d >= 0:
                mesh.Faces.AddFace(b, c, d)

    mesh.Normals.ComputeNormals()
    mesh.Compact()
    return mesh


def vertex_deltas(grade, mesh):
    """Per-vertex (design - terrain) by snapping each vertex to its grid node."""
    out = []
    inv = 1.0 / grade.cell
    for v in mesh.Vertices:
        i = int(round((v.X - grade.x0) * inv))
        j = int(round((v.Y - grade.y0) * inv))
        d = 0.0
        if 0 <= j < grade.ny and 0 <= i < grade.nx:
            zd = grade.z_design[j][i]
            zt = grade.z_terrain[j][i]
            if zd is not None and zt is not None:
                d = zd - zt
        out.append(d)
    return out


def tint_by_delta(mesh, grade, ramp, scale=None):
    """Assign mesh.VertexColors from per-vertex cut/fill depth.

    ramp(t) -> System.Drawing.Color for t in [-1, 1] (t<0 cut, t>0 fill).
    scale   : symmetric normaliser (max |delta|); auto from data when None.
    Returns the scale used (so a legend can label it).
    """
    deltas = vertex_deltas(grade, mesh)
    if scale is None:
        scale = max([abs(d) for d in deltas] or [1.0]) or 1.0

    mesh.VertexColors.Clear()
    for d in deltas:
        t = max(-1.0, min(1.0, d / scale))
        mesh.VertexColors.Add(ramp(t))
    return scale


def _interp_grid_z(grid_2d, grade, x, y):
    """Bilinear interpolation over a grade grid's 2-D Z array at world position (x, y)."""
    fi = (x - grade.x0) / grade.cell
    fj = (y - grade.y0) / grade.cell
    i0 = int(fi)
    j0 = int(fj)
    if i0 < 0 or j0 < 0 or i0 >= grade.nx or j0 >= grade.ny:
        return None
    i1 = min(i0 + 1, grade.nx - 1)
    j1 = min(j0 + 1, grade.ny - 1)
    tx = fi - i0
    ty = fj - j0
    z00 = grid_2d[j0][i0]
    z10 = grid_2d[j0][i1]
    z01 = grid_2d[j1][i0]
    z11 = grid_2d[j1][i1]
    vals = [z for z in (z00, z10, z01, z11) if z is not None]
    if not vals:
        return None
    if len(vals) < 4:
        return sum(vals) / len(vals)
    return (z00 * (1 - tx) * (1 - ty) +
            z10 * tx       * (1 - ty) +
            z01 * (1 - tx) * ty +
            z11 * tx       * ty)


def deform_terrain_to_grade(terrain_mesh, grade):
    """Return a copy of terrain_mesh with the design *delta* applied to vertices inside
    the analysis zone.

    Uses (z_design - z_terrain) as the displacement so that vertices outside the active
    grading zone receive a delta of exactly zero — the mesh is not distorted there and
    the seam at the analysis-zone boundary is seamless.
    """
    x_lo = grade.x0
    x_hi = grade.x0 + (grade.nx - 1) * grade.cell
    y_lo = grade.y0
    y_hi = grade.y0 + (grade.ny - 1) * grade.cell

    out = rg.Mesh()
    for vi in range(terrain_mesh.Vertices.Count):
        v = terrain_mesh.Vertices[vi]
        x, y, z_v = float(v.X), float(v.Y), float(v.Z)
        if x_lo <= x <= x_hi and y_lo <= y <= y_hi:
            z_d = _interp_grid_z(grade.z_design,  grade, x, y)
            z_t = _interp_grid_z(grade.z_terrain, grade, x, y)
            if z_d is not None and z_t is not None and abs(z_d - z_t) > 0.001:
                # Active grading zone: place vertex exactly at design elevation.
                # Using z_d directly (not z_v + delta) avoids grid-interpolation
                # error that grows with cell size, keeping pads flat and slopes correct.
                out.Vertices.Add(x, y, z_d)
            else:
                out.Vertices.Add(x, y, z_v)
        else:
            out.Vertices.Add(x, y, z_v)

    for fi in range(terrain_mesh.Faces.Count):
        f = terrain_mesh.Faces[fi]
        if f.IsQuad:
            out.Faces.AddFace(f.A, f.B, f.C, f.D)
        else:
            out.Faces.AddFace(f.A, f.B, f.C)

    out.Normals.ComputeNormals()
    out.Compact()
    return out


def default_ramp(t):
    """Blue (cut) -> light neutral (0) -> red (fill). t in [-1, 1]."""
    cut  = (42, 139, 156)    # Mar Caribe teal-blue
    mid  = (235, 230, 222)   # warm neutral
    fill = (224, 115, 92)    # salmon red-orange
    if t < 0:
        a, b, f = mid, cut, -t
    else:
        a, b, f = mid, fill, t
    r = int(round(a[0] + (b[0] - a[0]) * f))
    g = int(round(a[1] + (b[1] - a[1]) * f))
    bl = int(round(a[2] + (b[2] - a[2]) * f))
    return sd.Color.FromArgb(r, g, bl)
