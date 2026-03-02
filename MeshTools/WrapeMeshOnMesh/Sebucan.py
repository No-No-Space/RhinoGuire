#! python3
# -*- coding: utf-8 -*-
# __title__ = "Sebucan"
# __doc__ = """Version = 0.3
# Date    = 2026-03-02
# _____________________________________________________________________
# Description:
# Wraps one or more source meshes onto a destination object along the Z axis.
# Each vertex of the source mesh is projected (raycast) vertically onto the
# destination geometry, preserving its X/Y position and snapping its Z.
# Destination can be a Mesh, SubD, Surface, or Polysurface (solid included).
# Typical use case: adjusting road meshes to follow the contours of a terrain.
# The window is modeless — Rhino stays accessible while it is open.
# _____________________________________________________________________
# How-to:
# -> Run the script in Rhino 8 (RunPythonScript). The Sebucan panel opens.
# -> Click "Select Destination" and pick the target geometry (terrain, solid…).
# -> Click "Select Source Mesh(es)" and pick the mesh(es) to be wrapped (roads).
# -> Optionally enable "Replace source mesh(es)" to delete the originals.
# -> Click "Wrap!" to run the projection. New meshes are added to the same layer.
# -> Vertices outside the destination bounding box keep their original Z.
# _____________________________________________________________________
# Last update:
# - [02.03.2026] - 0.3 Adaptive refinement: splits faces where Z deviation exceeds tolerance
# - [02.03.2026] - 0.2 Destination now accepts Mesh, SubD, Surface, Polysurface
# - [02.03.2026] - 0.1 Initial release
# _____________________________________________________________________

import System
import rhinoscriptsyntax as rs
import Rhino
import Rhino.Geometry as rg
import Eto.Drawing as drawing
import Eto.Forms as forms
import scriptcontext as sc

# Selection filter: mesh + surface + polysurface + SubD
_DEST_FILTER = rs.filter.mesh | rs.filter.surface | rs.filter.polysurface
try:
    _DEST_FILTER |= rs.filter.subd
except AttributeError:
    pass  # older build without SubD filter constant


# ---------------------------------------------------------------------------
# Geometry helpers
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


def coerce_to_destination_mesh(obj_id):
    """Convert a Mesh, SubD, Surface, or Polysurface to a Mesh for ray intersection.

    Conversion order:
      Mesh       → used directly
      BRep       → Mesh.CreateFromBrep
      SubD       → SubD.ToBrep → Mesh.CreateFromBrep

    Returns the Mesh, or None if conversion fails.
    """
    mesh = rs.coercemesh(obj_id)
    if mesh is not None:
        return mesh

    brep = rs.coercebrep(obj_id)
    if brep is not None:
        return _brep_to_mesh(brep)

    geo = rs.coercegeometry(obj_id)
    if isinstance(geo, rg.SubD):
        brep = geo.ToBrep()
        if brep is not None:
            return _brep_to_mesh(brep)

    return None


def _obj_type_label(obj_id):
    """Return a short display string for the object type."""
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


# Dropdown values for "Max. passes" control (index → iteration count)
_ITER_VALUES = [1, 2, 3, 4, 5, 6, 8]


# ---------------------------------------------------------------------------
# Core geometry
# ---------------------------------------------------------------------------

def _make_projector(dest_mesh):
    """Return a project_z(x, y) closure that raycasts onto dest_mesh.

    Results are cached by (x, y) so each unique position is only projected
    once — critical for the adaptive function where midpoints are reused.
    Returns the new Z float, or None if (x, y) falls outside dest_mesh.
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


def wrap_mesh_on_mesh(source_mesh, dest_mesh):
    """Project every vertex of source_mesh onto dest_mesh along the Z axis.

    Returns: (wrapped_mesh, hit_count, miss_count)
    """
    result    = source_mesh.DuplicateMesh()
    project_z = _make_projector(dest_mesh)
    hit_count  = 0
    miss_count = 0

    for i in range(result.Vertices.Count):
        v = result.Vertices[i]
        x, y = float(v.X), float(v.Y)
        z = project_z(x, y)
        if z is not None:
            result.Vertices.SetVertex(i, rg.Point3f(x, y, z))
            hit_count += 1
        else:
            miss_count += 1

    result.Normals.ComputeNormals()
    result.Compact()
    return result, hit_count, miss_count


def adaptive_wrap_mesh(source_mesh, dest_mesh, tolerance, max_iterations):
    """Project source_mesh onto dest_mesh and adaptively refine coarse faces.

    After each projection pass, every triangle face is tested: the terrain Z at
    each edge midpoint is compared to the face's interpolated Z at that point.
    If the deviation on any edge exceeds `tolerance`, the face is split into 4
    sub-triangles (midpoint subdivision). New vertices are projected immediately.
    The loop repeats until no face fails the tolerance test or `max_iterations`
    is reached. Faces over flat terrain are never split.

    Returns: (refined_mesh, total_vertices_projected, miss_count)
    """
    project_z = _make_projector(dest_mesh)

    work = source_mesh.DuplicateMesh()
    work.Faces.ConvertQuadsToTriangles()

    # Work in plain Python lists — faster than repeated RhinoCommon calls
    verts = [[float(v.X), float(v.Y), float(v.Z)] for v in work.Vertices]
    faces = []
    for fi in range(work.Faces.Count):
        f = work.Faces[fi]
        faces.append((f.A, f.B, f.C))

    # Project initial vertices
    miss_count = 0
    for i, v in enumerate(verts):
        z = project_z(v[0], v[1])
        if z is not None:
            verts[i][2] = z
        else:
            miss_count += 1

    # edge_vert: (min_vi, max_vi) → index of the midpoint vertex.
    # Persists across iterations so shared edges are never duplicated.
    edge_vert = {}

    for _ in range(max_iterations):
        new_faces = []
        any_split = False

        for face in faces:
            i0, i1, i2 = face
            v0, v1, v2  = verts[i0], verts[i1], verts[i2]

            # Midpoint coords + interpolated Z for each of the 3 edges
            edge_def = [
                (i0, i1, (v0[0]+v1[0])/2, (v0[1]+v1[1])/2, (v0[2]+v1[2])/2),
                (i1, i2, (v1[0]+v2[0])/2, (v1[1]+v2[1])/2, (v1[2]+v2[2])/2),
                (i2, i0, (v2[0]+v0[0])/2, (v2[1]+v0[1])/2, (v2[2]+v0[2])/2),
            ]

            # Evaluate terrain Z at each midpoint and check deviation
            needs_split = False
            mid_data    = []
            for ia, ib, mx, my, interp_z in edge_def:
                true_z = project_z(mx, my)
                if true_z is None:
                    true_z = interp_z          # outside destination — keep interpolated
                mid_data.append((ia, ib, mx, my, true_z))
                if abs(true_z - interp_z) > tolerance:
                    needs_split = True

            if not needs_split:
                new_faces.append(face)
                continue

            # Split: get or create the three midpoint vertices
            any_split = True
            mid_idxs  = []
            for ia, ib, mx, my, mz in mid_data:
                key = (min(ia, ib), max(ia, ib))
                if key not in edge_vert:
                    edge_vert[key] = len(verts)
                    verts.append([mx, my, mz])
                mid_idxs.append(edge_vert[key])

            m01, m12, m20 = mid_idxs
            new_faces += [
                (i0, m01, m20),
                (i1, m12, m01),
                (i2, m20, m12),
                (m01, m12, m20),   # centre triangle
            ]

        faces = new_faces
        if not any_split:
            break

    # Assemble final RhinoCommon Mesh
    result = rg.Mesh()
    for v in verts:
        result.Vertices.Add(v[0], v[1], v[2])
    for f in faces:
        result.Faces.AddFace(f[0], f[1], f[2])
    result.Normals.ComputeNormals()
    result.Compact()

    hit_count = len(verts) - miss_count
    return result, hit_count, miss_count


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

class SebucanForm(forms.Form):

    def __init__(self):
        super().__init__()
        self.Title     = "Sebucan — Wrap Mesh on Mesh"
        self.Resizable = False
        self.Padding   = drawing.Padding(10)
        self.MinimumSize = drawing.Size(320, 420)
        self.Owner     = Rhino.UI.RhinoEtoApp.MainWindow

        self.dest_id    = None
        self.source_ids = []

        self._build_ui()

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------

    def _build_ui(self):
        layout = forms.DynamicLayout()
        layout.DefaultSpacing = drawing.Size(5, 5)

        # Header
        title_lbl = forms.Label()
        title_lbl.Text = "Wrap Mesh on Mesh"
        title_lbl.Font = drawing.Font(drawing.SystemFont.Bold, 12)
        layout.AddRow(title_lbl)

        desc_lbl = forms.Label()
        desc_lbl.Text = "Projects source vertices onto the destination\nmesh along the Z axis."
        desc_lbl.TextColor = drawing.Colors.Gray
        layout.AddRow(desc_lbl)

        layout.AddRow(None)

        # Step 1 — Destination
        s1_lbl = forms.Label()
        s1_lbl.Text = "1 — Destination  (mesh, SubD, polysurface…)"
        s1_lbl.Font = drawing.Font(drawing.SystemFont.Bold)
        layout.AddRow(s1_lbl)

        self.dest_btn = forms.Button()
        self.dest_btn.Text = "Select Destination"
        self.dest_btn.Click += self.on_select_destination
        layout.AddRow(self.dest_btn)

        self.dest_info = forms.Label()
        self.dest_info.Text = "No destination selected."
        self.dest_info.TextColor = drawing.Colors.Gray
        layout.AddRow(self.dest_info)

        layout.AddRow(None)

        # Step 2 — Sources
        s2_lbl = forms.Label()
        s2_lbl.Text = "2 — Source Mesh(es)  (e.g., roads)"
        s2_lbl.Font = drawing.Font(drawing.SystemFont.Bold)
        layout.AddRow(s2_lbl)

        self.src_btn = forms.Button()
        self.src_btn.Text = "Select Source Mesh(es)"
        self.src_btn.Enabled = False
        self.src_btn.Click += self.on_select_sources
        layout.AddRow(self.src_btn)

        self.src_info = forms.Label()
        self.src_info.Text = "No sources selected."
        self.src_info.TextColor = drawing.Colors.Gray
        layout.AddRow(self.src_info)

        layout.AddRow(None)

        # Options
        opts_lbl = forms.Label()
        opts_lbl.Text = "Options"
        opts_lbl.Font = drawing.Font(drawing.SystemFont.Bold)
        layout.AddRow(opts_lbl)

        self.replace_check = forms.CheckBox()
        self.replace_check.Text = "Replace source mesh(es) with result"
        self.replace_check.Checked = False
        layout.AddRow(self.replace_check)

        self.adaptive_check = forms.CheckBox()
        self.adaptive_check.Text = "Adaptive refinement"
        self.adaptive_check.Checked = False
        self.adaptive_check.CheckedChanged += self.on_adaptive_changed
        layout.AddRow(self.adaptive_check)

        adapt_hint = forms.Label()
        adapt_hint.Text = "Splits faces where terrain Z deviation exceeds tolerance."
        adapt_hint.TextColor = drawing.Colors.Gray
        layout.AddRow(adapt_hint)

        tol_row = forms.StackLayout()
        tol_row.Orientation = forms.Orientation.Horizontal
        tol_row.Spacing = 8
        tol_lbl = forms.Label()
        tol_lbl.Text = "Tolerance:"
        tol_lbl.Width = 80
        self.tol_input = forms.TextBox()
        self.tol_input.Text = "0.1"
        self.tol_input.Width = 55
        self.tol_input.Enabled = False
        tol_units = forms.Label()
        tol_units.Text = "(Z units)"
        tol_units.TextColor = drawing.Colors.Gray
        tol_row.Items.Add(forms.StackLayoutItem(tol_lbl))
        tol_row.Items.Add(forms.StackLayoutItem(self.tol_input))
        tol_row.Items.Add(forms.StackLayoutItem(tol_units))
        layout.AddRow(tol_row)

        iter_row = forms.StackLayout()
        iter_row.Orientation = forms.Orientation.Horizontal
        iter_row.Spacing = 8
        iter_lbl = forms.Label()
        iter_lbl.Text = "Max. passes:"
        iter_lbl.Width = 80
        self.iter_drop = forms.DropDown()
        for v in _ITER_VALUES:
            self.iter_drop.Items.Add(str(v))
        self.iter_drop.SelectedIndex = 2  # default: 3 passes
        self.iter_drop.Enabled = False
        iter_row.Items.Add(forms.StackLayoutItem(iter_lbl))
        iter_row.Items.Add(forms.StackLayoutItem(self.iter_drop))
        layout.AddRow(iter_row)

        layout.AddRow(None)

        # Wrap button
        self.wrap_btn = forms.Button()
        self.wrap_btn.Text = "Wrap!"
        self.wrap_btn.Enabled = False
        self.wrap_btn.Click += self.on_wrap
        layout.AddRow(self.wrap_btn)

        layout.AddRow(None)

        # Status
        self.status_lbl = forms.Label()
        self.status_lbl.Text = "Ready — select a destination mesh to begin."
        self.status_lbl.TextColor = drawing.Colors.Gray
        layout.AddRow(self.status_lbl)

        # Close
        close_btn = forms.Button()
        close_btn.Text = "Close"
        close_btn.Click += lambda s, e: self.Close()
        layout.AddRow(close_btn)

        self.Content = layout

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _set_status(self, text, color=None):
        self.status_lbl.Text  = text
        self.status_lbl.TextColor = color if color is not None else drawing.Colors.Gray

    def _update_wrap_btn(self):
        self.wrap_btn.Enabled = (self.dest_id is not None) and (len(self.source_ids) > 0)

    def on_adaptive_changed(self, *_):
        enabled = self.adaptive_check.Checked
        self.tol_input.Enabled = enabled
        self.iter_drop.Enabled = enabled

    # ------------------------------------------------------------------
    # Selection handlers
    # ------------------------------------------------------------------

    def on_select_destination(self, sender, e):
        self._set_status("Select destination geometry in the viewport…")
        obj_id = rs.GetObject("Select destination (mesh, SubD, polysurface)", _DEST_FILTER, preselect=True)

        if obj_id is None:
            self._set_status("Destination selection cancelled.")
            return

        self.dest_id = obj_id
        name = rs.ObjectName(obj_id) or "(unnamed)"
        type_label = _obj_type_label(obj_id)
        self.dest_info.Text = f"\u2713 {name}  [{type_label}]"
        self.dest_info.TextColor = drawing.Colors.Green
        self.src_btn.Enabled = True
        self._set_status("Destination set. Now select source mesh(es).")
        self._update_wrap_btn()

    def on_select_sources(self, sender, e):
        self._set_status("Select source mesh(es) in the viewport…")
        obj_ids = rs.GetObjects("Select source mesh(es)", rs.filter.mesh, preselect=True)

        if not obj_ids:
            self._set_status("Source selection cancelled.")
            return

        self.source_ids = list(obj_ids)
        n = len(self.source_ids)
        self.src_info.Text = f"\u2713 {n} mesh(es) selected"
        self.src_info.TextColor = drawing.Colors.Green
        self._set_status(f"{n} source mesh(es) ready. Click Wrap! to process.")
        self._update_wrap_btn()

    # ------------------------------------------------------------------
    # Wrap
    # ------------------------------------------------------------------

    def on_wrap(self, sender, e):
        self.wrap_btn.Enabled = False
        self._set_status("Converting destination to mesh…")

        dest_mesh = coerce_to_destination_mesh(self.dest_id)
        if dest_mesh is None:
            self._set_status("Error: could not convert destination to mesh.", drawing.Colors.Red)
            self.wrap_btn.Enabled = True
            return

        replace       = self.replace_check.Checked
        use_adaptive  = self.adaptive_check.Checked
        total_hits    = 0
        total_miss    = 0
        total_added   = 0
        results_added = 0
        n_src         = len(self.source_ids)

        if use_adaptive:
            try:
                tol = float(self.tol_input.Text)
            except ValueError:
                tol = 0.1
            max_iter = _ITER_VALUES[self.iter_drop.SelectedIndex]

        for idx, src_id in enumerate(self.source_ids):
            label = "adaptive" if use_adaptive else "projecting"
            self._set_status(f"Wrapping mesh {idx + 1} of {n_src} ({label})…")

            src_mesh = rs.coercemesh(src_id)
            if src_mesh is None:
                continue

            obj_layer  = rs.ObjectLayer(src_id)
            obj_name   = rs.ObjectName(src_id)
            orig_count = src_mesh.Vertices.Count

            if use_adaptive:
                wrapped, hits, misses = adaptive_wrap_mesh(src_mesh, dest_mesh, tol, max_iter)
                total_added += wrapped.Vertices.Count - orig_count
            else:
                wrapped, hits, misses = wrap_mesh_on_mesh(src_mesh, dest_mesh)

            total_hits += hits
            total_miss += misses

            if replace:
                rs.DeleteObject(src_id)

            new_id = sc.doc.Objects.AddMesh(wrapped)
            if new_id != System.Guid.Empty:
                rs.ObjectLayer(new_id, obj_layer)
                if obj_name:
                    rs.ObjectName(new_id, obj_name)
                results_added += 1

        sc.doc.Views.Redraw()

        msg = f"Done! {results_added} mesh(es) wrapped ({total_hits} vertices projected)."
        if total_added > 0:
            msg += f" {total_added} vertices added by refinement."
        if total_miss > 0:
            msg += f" {total_miss} outside destination — original Z kept."
        self._set_status(msg, drawing.Colors.Green)
        self.wrap_btn.Enabled = True


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    form = SebucanForm()
    form.Show()


if __name__ == "__main__":
    main()
