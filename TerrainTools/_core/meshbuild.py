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
