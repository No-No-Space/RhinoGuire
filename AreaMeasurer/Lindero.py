#! python3
# r: openpyxl
# -*- coding: utf-8 -*-
# __title__ = "Lindero"
# __doc__ = """Version = 0.2
# Date    = 2026-02-20
# Author: Aquelon - aquelon@pm.me
# _____________________________________________________________________
# Description:
# Footprint area calculator for Rhino objects.
# "Footprint" = the plan area (XY projection), NOT the sum of all surfaces.
# Modeless window — Rhino stays accessible while it is open.
# _____________________________________________________________________
# Three scenarios:
#   S1 — Selected Objects:
#        Individual footprint per object, no overlap handling.
#        Results: area per object + total.
#
#   S2 — By Layer:
#        All objects on a layer. Overlapping footprints are merged with a
#        Boolean Union to avoid double-counting.
#        Results: area per object key + combined layer total.
#
#   S3 — By Layer Hierarchy:
#        Parent layer contains sublayers (each sublayer = one level/floor).
#        Objects carry two user text keys: an object key (individual name)
#        and a group key (department / class).
#        Overlaps removed within each sublayer.
#        Results: per object, per group subtotal, per sublayer total, grand total.
# _____________________________________________________________________
# How-to:
# -> Run the script; the calculator window opens (modeless)
# -> S1: Select objects in Rhino while the form is open, then click Calculate
# -> S2: Pick a Layer and an Object Key, click Calculate
# -> S3: Pick a Parent Layer, Object Key, and Group Key, click Calculate
# -> Click "Refresh Model" if you add new layers or user text keys
# _____________________________________________________________________
# Last update:
# - [20.02.2026] - 0.2 Export results to Excel (Objects + Summary sheets)
# - [20.02.2026] - 0.1 Initial release
# _____________________________________________________________________
# To-Do:
# - Compare calculated areas against Excel-defined area goals
# _____________________________________________________________________


import rhinoscriptsyntax as rs
import Rhino
import Rhino.Geometry as rg
import scriptcontext as sc
import Eto.Drawing as drawing
import Eto.Forms as forms
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ══════════════════════════════════════════════════════════════════════════════
# Model helpers
# ══════════════════════════════════════════════════════════════════════════════

def get_all_user_text_keys():
    """Return sorted list of unique user text keys found on any object in the model."""
    all_objects = rs.AllObjects()
    if not all_objects:
        return []
    keys = set()
    for guid in all_objects:
        obj_keys = rs.GetUserText(guid)  # returns list of key strings
        if obj_keys:
            for k in obj_keys:
                keys.add(k)
    return sorted(keys)


def all_layer_names():
    """Return sorted list of all layer full path names in the model."""
    return sorted(rs.LayerNames() or [])


def short_name(full_path):
    """Return the last segment of a layer full path (last part after '::')."""
    return full_path.split("::")[-1] if "::" in full_path else full_path


def get_layer_objects(layer_name):
    """Return GUIDs of objects directly on the given layer."""
    return rs.ObjectsByLayer(layer_name) or []


def get_child_layers(parent_layer_name):
    """
    Return full path names of direct child layers of the given layer.
    Uses string matching on layer names to avoid RhinoCommon attribute differences.
    """
    prefix = parent_layer_name + "::"
    return sorted(
        name for name in (rs.LayerNames() or [])
        if name.startswith(prefix) and "::" not in name[len(prefix):]
    )


def unit_label():
    """Return the squared-unit label matching the current Rhino model unit system."""
    mapping = {
        "Millimeters": "mm²", "Centimeters": "cm²", "Meters": "m²",
        "Kilometers": "km²", "Feet": "ft²", "Inches": "in²",
    }
    return mapping.get(sc.doc.ModelUnitSystem.ToString(), "units²")


# ══════════════════════════════════════════════════════════════════════════════
# Footprint geometry
# ══════════════════════════════════════════════════════════════════════════════

def _bbox_footprint(obj_guid):
    """Bounding-box fallback: XY rectangle from the object's bounding box."""
    rhobj = sc.doc.Objects.FindId(obj_guid)
    if rhobj is None:
        return None
    bbox = rhobj.GetBoundingBox(True)
    if not bbox.IsValid:
        return None
    pts = [
        rg.Point3d(bbox.Min.X, bbox.Min.Y, 0),
        rg.Point3d(bbox.Max.X, bbox.Min.Y, 0),
        rg.Point3d(bbox.Max.X, bbox.Max.Y, 0),
        rg.Point3d(bbox.Min.X, bbox.Max.Y, 0),
        rg.Point3d(bbox.Min.X, bbox.Min.Y, 0),
    ]
    return rg.PolylineCurve(pts)


def _brep_footprint_curves(brep):
    """
    Find the bottom horizontal face(s) of a Brep and return their outer border
    curves projected to Z=0.
    Falls back to the XY bounding box if no horizontal faces are found.
    """
    tol = sc.doc.ModelAbsoluteTolerance
    proj = rg.Transform.PlanarProjection(rg.Plane.WorldXY)

    # Gather all horizontal faces (|normalZ| > 0.9) with their centroid Z
    horiz = []
    for face in brep.Faces:
        u = (face.Domain(0).Min + face.Domain(0).Max) * 0.5
        v = (face.Domain(1).Min + face.Domain(1).Max) * 0.5
        normal = face.NormalAt(u, v)
        if abs(normal.Z) > 0.9:
            pt = face.PointAt(u, v)
            horiz.append((face, pt.Z))

    if not horiz:
        # No horizontal faces → bounding-box fallback
        bbox = brep.GetBoundingBox(True)
        if not bbox.IsValid:
            return []
        pts = [
            rg.Point3d(bbox.Min.X, bbox.Min.Y, 0),
            rg.Point3d(bbox.Max.X, bbox.Min.Y, 0),
            rg.Point3d(bbox.Max.X, bbox.Max.Y, 0),
            rg.Point3d(bbox.Min.X, bbox.Max.Y, 0),
            rg.Point3d(bbox.Min.X, bbox.Min.Y, 0),
        ]
        return [rg.PolylineCurve(pts)]

    # Keep only the bottom-most horizontal face(s)
    min_z = min(z for _, z in horiz)
    bottom = [f for f, z in horiz if abs(z - min_z) <= tol * 100]

    curves = []
    for face in bottom:
        if face.OuterLoop is None:
            continue
        border = face.OuterLoop.To3dCurve()
        if border is None:
            continue
        dup = border.DuplicateCurve()
        dup.Transform(proj)
        if dup.IsClosed:
            curves.append(dup)
    return curves


def get_footprint_curves(obj_guid):
    """
    Return a list of planar closed curves at Z=0 representing the footprint
    of the object. Returns an empty list if the footprint cannot be determined.
    """
    rhobj = sc.doc.Objects.FindId(obj_guid)
    if rhobj is None:
        return []
    geom = rhobj.Geometry

    # Extrusion and Surface → convert to Brep for uniform handling
    if isinstance(geom, (rg.Extrusion, rg.Surface)):
        brep = geom.ToBrep()
        if brep:
            return _brep_footprint_curves(brep)

    if isinstance(geom, rg.Brep):
        return _brep_footprint_curves(geom)

    # Closed planar curve (e.g., a flat hatch boundary used as a space outline)
    if isinstance(geom, rg.Curve) and geom.IsClosed and geom.IsPlanar():
        proj = rg.Transform.PlanarProjection(rg.Plane.WorldXY)
        dup = geom.DuplicateCurve()
        dup.Transform(proj)
        return [dup]

    # Fallback: bounding box
    fallback = _bbox_footprint(obj_guid)
    return [fallback] if fallback else []


def curve_area(curve):
    """Return the area enclosed by a planar closed curve, or 0 on failure."""
    try:
        amp = rg.AreaMassProperties.Compute(curve)
        return amp.Area if amp else 0.0
    except Exception:
        return 0.0


def get_footprint_area(obj_guid):
    """Footprint area of a single object (sum of bottom face areas)."""
    return sum(curve_area(c) for c in get_footprint_curves(obj_guid))


def combined_area(obj_guids):
    """
    Total footprint area for a list of objects after Boolean Union on their
    projected outlines (removes overlapping footprints).
    Returns (area: float, union_succeeded: bool).
    """
    all_curves = []
    for guid in obj_guids:
        all_curves.extend(get_footprint_curves(guid))

    if not all_curves:
        return 0.0, False
    if len(all_curves) == 1:
        return curve_area(all_curves[0]), True

    tol = sc.doc.ModelAbsoluteTolerance
    try:
        unioned = rg.Curve.CreateBooleanUnion(all_curves, tol)
    except Exception:
        unioned = None

    if unioned and len(unioned) > 0:
        return sum(curve_area(c) for c in unioned), True

    # Boolean union failed — return plain sum (may double-count overlaps)
    return sum(curve_area(c) for c in all_curves), False


# ══════════════════════════════════════════════════════════════════════════════
# Calculation engines
# ══════════════════════════════════════════════════════════════════════════════

def _label(guid, key):
    """Object display label: user text key → Rhino object name → short GUID."""
    if key:
        val = rs.GetUserText(guid, key)
        if val:
            return val
    name = rs.ObjectName(guid)
    return name if name else str(guid)[:8] + "…"


def calc_s1(name_key):
    """
    Scenario 1 — Selected objects.
    Returns list of {name, area}.
    """
    guids = rs.SelectedObjects() or []
    return [{"guid": str(g), "name": _label(g, name_key), "area": get_footprint_area(g)} for g in guids]


def calc_s2(layer_name, obj_key):
    """
    Scenario 2 — All objects on one layer, footprints merged to avoid double-counting.
    Returns {objects: [{name, area}], total: float, union_ok: bool}.
    """
    guids = get_layer_objects(layer_name)
    per_obj = [{"guid": str(g), "name": _label(g, obj_key), "area": get_footprint_area(g)} for g in guids]
    total, union_ok = combined_area(guids)
    return {"objects": per_obj, "total": total, "union_ok": union_ok}


def calc_s3(parent_layer, obj_key, grp_key):
    """
    Scenario 3 — Sublayer hierarchy (each sublayer = one floor/level).
    Overlaps removed within each sublayer independently.
    Grand total = sum of per-sublayer totals (floors are additive).
    Returns {
        sublayers: {name: {objects, group_totals, total, union_ok}},
        overall_total: float
    }.
    """
    result = {"sublayers": {}, "overall_total": 0.0}
    for sl in get_child_layers(parent_layer):
        guids = get_layer_objects(sl)
        if not guids:
            result["sublayers"][sl] = {
                "objects": [], "group_totals": {}, "total": 0.0, "union_ok": True
            }
            continue

        objects = []
        groups = {}  # group_val → list[guid]
        for guid in guids:
            grp_val = (rs.GetUserText(guid, grp_key) if grp_key else None) or "—"
            objects.append({
                "guid": str(guid),
                "name": _label(guid, obj_key),
                "group": grp_val,
                "area": get_footprint_area(guid),
            })
            groups.setdefault(grp_val, []).append(guid)

        total, union_ok = combined_area(guids)
        group_totals = {gv: combined_area(gguids)[0] for gv, gguids in groups.items()}

        result["sublayers"][sl] = {
            "objects": objects,
            "group_totals": group_totals,
            "total": total,
            "union_ok": union_ok,
        }
        result["overall_total"] += total

    return result


# ══════════════════════════════════════════════════════════════════════════════
# Results formatting (monospace text for the TextArea)
# ══════════════════════════════════════════════════════════════════════════════

W = 62  # result panel width in characters


def _rule(char="═"):
    return char * W


def _row(left, right="", lw=38):
    return f"  {left:<{lw}}{right:>{W - lw - 2}}"


def _fmt(v):
    return f"{v:,.4f}"


def format_s1(results, unit):
    if not results:
        return "  No objects selected.\n"
    lines = [
        _rule(),
        f"  SCENARIO 1 — SELECTED OBJECTS   [{unit}]",
        _rule("─"),
        _row("Object", "Area"),
        _rule("─"),
    ]
    for r in results:
        lines.append(_row(r["name"][:36], _fmt(r["area"])))
    total = sum(r["area"] for r in results)
    lines += [_rule("─"), _row("TOTAL", _fmt(total)), _rule(), ""]
    return "\n".join(lines)


def format_s2(data, layer_name, obj_key, unit):
    if not data["objects"]:
        return f"  No objects found on layer '{layer_name}'.\n"
    raw_sum = sum(o["area"] for o in data["objects"])
    union_note = "" if data["union_ok"] else "  [union failed — sum shown]"
    lines = [
        _rule(),
        f"  SCENARIO 2 — BY LAYER   [{unit}]",
        f"  Layer : {layer_name}",
        f"  Key   : {obj_key or '(object name / GUID)'}",
        _rule("─"),
        _row("Object", "Area"),
        _rule("─"),
    ]
    for o in data["objects"]:
        lines.append(_row(o["name"][:36], _fmt(o["area"])))
    lines += [
        _rule("─"),
        _row("Sum (individual totals)", _fmt(raw_sum)),
        _row(f"TOTAL (footprint, overlaps merged){union_note}", _fmt(data["total"])),
        _rule(),
        "",
    ]
    return "\n".join(lines)


def format_s3(data, parent_layer, obj_key, grp_key, unit):
    if not data["sublayers"]:
        return f"  No sublayers found under '{parent_layer}'.\n"
    lines = [
        _rule(),
        f"  SCENARIO 3 — LAYER HIERARCHY   [{unit}]",
        f"  Parent     : {parent_layer}",
        f"  Object key : {obj_key or '(object name / GUID)'}",
        f"  Group key  : {grp_key or '—'}",
        _rule("═"),
        "",
    ]
    for sl, sl_data in data["sublayers"].items():
        sn = short_name(sl)
        union_note = "" if sl_data["union_ok"] else "  [union failed]"
        lines += [
            f"  ▸ {sn}   —   Total: {_fmt(sl_data['total'])} {unit}{union_note}",
            _rule("─"),
        ]

        if sl_data["group_totals"] and grp_key:
            lines.append(_row("Group (combined footprint)", "Area"))
            lines.append("  " + "·" * (W - 2))
            for gv, ga in sorted(sl_data["group_totals"].items()):
                lines.append(_row(f"  {gv}"[:36], _fmt(ga)))
            lines.append(_rule("─"))

        if sl_data["objects"]:
            if grp_key:
                n_w, g_w, a_w = 26, 20, W - 26 - 20 - 4
                lines.append(f"  {'Object':<{n_w}}{'Group':<{g_w}}{'Area':>{a_w}}")
                lines.append("  " + "·" * (W - 2))
                for o in sl_data["objects"]:
                    n = o["name"][:n_w - 1]
                    g = o["group"][:g_w - 1]
                    a = _fmt(o["area"])
                    lines.append(f"  {n:<{n_w}}{g:<{g_w}}{a:>{a_w}}")
            else:
                lines.append(_row("Object", "Area"))
                lines.append("  " + "·" * (W - 2))
                for o in sl_data["objects"]:
                    lines.append(_row(o["name"][:36], _fmt(o["area"])))
        lines.append("")

    lines += [
        _rule("═"),
        _row("OVERALL TOTAL (sum of all levels)", _fmt(data["overall_total"])),
        _rule(),
        "",
    ]
    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════

class LinderoForm(forms.Form):
    """Modeless footprint area calculator — stays open between calculations."""

    def __init__(self):
        super().__init__()
        self.Title = "Lindero — Footprint Area Calculator"
        self.Padding = drawing.Padding(10)
        self.Resizable = True
        self.MinimumSize = drawing.Size(640, 620)
        self.Owner = Rhino.UI.RhinoEtoApp.MainWindow

        self.available_keys = get_all_user_text_keys()
        self.available_layers = all_layer_names()
        self._export_data = None  # populated after each successful calculation

        self._build_ui()

    # ------------------------------------------------------------------
    # Layout builders
    # ------------------------------------------------------------------

    def _build_ui(self):
        outer = forms.StackLayout()
        outer.Orientation = forms.Orientation.Vertical
        outer.Spacing = 8
        outer.HorizontalContentAlignment = forms.HorizontalAlignment.Stretch

        # Tabs
        self.tabs = forms.TabControl()
        self.tabs.Pages.Add(self._tab_s1())
        self.tabs.Pages.Add(self._tab_s2())
        self.tabs.Pages.Add(self._tab_s3())
        outer.Items.Add(forms.StackLayoutItem(self.tabs))

        # Button row
        btn_row = forms.StackLayout()
        btn_row.Orientation = forms.Orientation.Horizontal
        btn_row.Spacing = 8

        calc_btn = forms.Button()
        calc_btn.Text = "Calculate"
        calc_btn.Click += self.on_calculate

        clear_btn = forms.Button()
        clear_btn.Text = "Clear"
        clear_btn.Click += self.on_clear

        refresh_btn = forms.Button()
        refresh_btn.Text = "Refresh Model"
        refresh_btn.Click += self.on_refresh_model

        export_btn = forms.Button()
        export_btn.Text = "Export to Excel"
        export_btn.Click += self.on_export

        btn_row.Items.Add(forms.StackLayoutItem(calc_btn))
        btn_row.Items.Add(forms.StackLayoutItem(clear_btn))
        btn_row.Items.Add(forms.StackLayoutItem(refresh_btn))
        btn_row.Items.Add(forms.StackLayoutItem(export_btn))
        outer.Items.Add(forms.StackLayoutItem(btn_row))

        # Results area (expands to fill remaining height)
        self.results_area = forms.TextArea()
        self.results_area.ReadOnly = True
        self.results_area.Font = drawing.Font("Courier New", 9)
        self.results_area.Text = (
            "Select a scenario tab, set the parameters, and click Calculate.\n"
        )
        outer.Items.Add(forms.StackLayoutItem(self.results_area, True))

        # Status label
        self.status_label = forms.Label()
        self.status_label.Text = f"Ready  —  {len(self.available_layers)} layer(s), {len(self.available_keys)} key(s) loaded."
        self.status_label.TextColor = drawing.Colors.Gray
        outer.Items.Add(forms.StackLayoutItem(self.status_label))

        self.Content = outer

    def _tab_s1(self):
        page = forms.TabPage()
        page.Text = "S1 — Selected Objects"

        layout = forms.DynamicLayout()
        layout.DefaultSpacing = drawing.Size(5, 6)
        layout.Padding = drawing.Padding(8)

        desc = forms.Label()
        desc.Text = "Individual footprint per selected object. No overlap handling."
        desc.TextColor = drawing.Colors.Gray
        layout.AddRow(desc)
        layout.AddRow(None)

        name_lbl = forms.Label()
        name_lbl.Text = "Name Key (optional):"

        self.name_key_combo = forms.ComboBox()
        self.name_key_combo.DataStore = self.available_keys
        self.name_key_combo.PlaceholderText = "Falls back to object name / GUID"
        self.name_key_combo.Width = 300

        layout.AddRow(name_lbl, self.name_key_combo)
        layout.AddRow(None)

        note = forms.Label()
        note.Text = "Select objects in Rhino (form stays open), then click Calculate."
        note.TextColor = drawing.Colors.Gray
        layout.AddRow(note)

        page.Content = layout
        return page

    def _tab_s2(self):
        page = forms.TabPage()
        page.Text = "S2 — By Layer"

        layout = forms.DynamicLayout()
        layout.DefaultSpacing = drawing.Size(5, 6)
        layout.Padding = drawing.Padding(8)

        desc = forms.Label()
        desc.Text = "All objects on a layer. Overlapping footprints merged (Boolean Union)."
        desc.TextColor = drawing.Colors.Gray
        layout.AddRow(desc)
        layout.AddRow(None)

        layer_lbl = forms.Label()
        layer_lbl.Text = "Layer:"

        self.layer_s2_dd = forms.DropDown()
        self._populate_layer_dd(self.layer_s2_dd)
        self.layer_s2_dd.Width = 300

        layout.AddRow(layer_lbl, self.layer_s2_dd)

        obj_key_lbl = forms.Label()
        obj_key_lbl.Text = "Object Key:"

        self.obj_key_s2 = forms.ComboBox()
        self.obj_key_s2.DataStore = self.available_keys
        self.obj_key_s2.PlaceholderText = "User text key for object labels"
        self.obj_key_s2.Width = 300

        layout.AddRow(obj_key_lbl, self.obj_key_s2)

        page.Content = layout
        return page

    def _tab_s3(self):
        page = forms.TabPage()
        page.Text = "S3 — Layer Hierarchy"

        layout = forms.DynamicLayout()
        layout.DefaultSpacing = drawing.Size(5, 6)
        layout.Padding = drawing.Padding(8)

        desc = forms.Label()
        desc.Text = "Parent layer + sublayers (each sublayer = one level / floor)."
        desc.TextColor = drawing.Colors.Gray
        layout.AddRow(desc)
        layout.AddRow(None)

        parent_lbl = forms.Label()
        parent_lbl.Text = "Parent Layer:"

        self.parent_layer_dd = forms.DropDown()
        self._populate_layer_dd(self.parent_layer_dd)
        self.parent_layer_dd.Width = 300

        layout.AddRow(parent_lbl, self.parent_layer_dd)
        layout.AddRow(None)

        obj_key_lbl = forms.Label()
        obj_key_lbl.Text = "Object Key:"

        self.obj_key_s3 = forms.ComboBox()
        self.obj_key_s3.DataStore = self.available_keys
        self.obj_key_s3.PlaceholderText = "Individual / small group name"
        self.obj_key_s3.Width = 300

        obj_key_hint = forms.Label()
        obj_key_hint.Text = "Individual or small group name (e.g. 'Room Name')"
        obj_key_hint.TextColor = drawing.Colors.Gray

        layout.AddRow(obj_key_lbl, self.obj_key_s3)
        layout.AddRow(None, obj_key_hint)
        layout.AddRow(None)

        grp_key_lbl = forms.Label()
        grp_key_lbl.Text = "Group Key:"

        self.grp_key_s3 = forms.ComboBox()
        self.grp_key_s3.DataStore = self.available_keys
        self.grp_key_s3.PlaceholderText = "Department / class (larger grouping)"
        self.grp_key_s3.Width = 300

        grp_key_hint = forms.Label()
        grp_key_hint.Text = "Larger class grouping (e.g. 'Department')"
        grp_key_hint.TextColor = drawing.Colors.Gray

        layout.AddRow(grp_key_lbl, self.grp_key_s3)
        layout.AddRow(None, grp_key_hint)

        page.Content = layout
        return page

    # ------------------------------------------------------------------
    # Layer dropdown helpers
    # ------------------------------------------------------------------

    def _populate_layer_dd(self, dd):
        dd.Items.Clear()
        for name in self.available_layers:
            dd.Items.Add(name)
        if self.available_layers:
            dd.SelectedIndex = 0

    def _selected_layer(self, dd):
        """Return the layer name corresponding to the current DropDown selection."""
        idx = dd.SelectedIndex
        if idx < 0 or idx >= len(self.available_layers):
            return None
        return self.available_layers[idx]

    def _restore_layer_dd(self, dd, prev_name):
        """Repopulate a layer DropDown, restoring the previous selection by name."""
        self._populate_layer_dd(dd)
        if prev_name and prev_name in self.available_layers:
            dd.SelectedIndex = self.available_layers.index(prev_name)

    # ------------------------------------------------------------------
    # Event handlers
    # ------------------------------------------------------------------

    def on_refresh_model(self, _sender, _e):
        prev_s2 = self._selected_layer(self.layer_s2_dd)
        prev_parent = self._selected_layer(self.parent_layer_dd)

        self.available_keys = get_all_user_text_keys()
        self.available_layers = all_layer_names()

        # Restore key combos (preserve typed text)
        for combo in [self.name_key_combo, self.obj_key_s2, self.obj_key_s3, self.grp_key_s3]:
            txt = combo.Text
            combo.DataStore = self.available_keys
            combo.Text = txt

        # Restore layer dropdowns
        self._restore_layer_dd(self.layer_s2_dd, prev_s2)
        self._restore_layer_dd(self.parent_layer_dd, prev_parent)

        self.status_label.Text = (
            f"Refreshed  —  {len(self.available_layers)} layer(s), {len(self.available_keys)} key(s)."
        )
        self.status_label.TextColor = drawing.Colors.Gray

    def on_calculate(self, _sender, _e):
        unit = unit_label()
        try:
            idx = self.tabs.SelectedIndex
            if idx == 0:
                self._run_s1(unit)
            elif idx == 1:
                self._run_s2(unit)
            else:
                self._run_s3(unit)
        except Exception as ex:
            self.status_label.Text = f"Error: {ex}"
            self.status_label.TextColor = drawing.Colors.Red

    def on_clear(self, _sender, _e):
        self.results_area.Text = ""
        self.status_label.Text = "Cleared."
        self.status_label.TextColor = drawing.Colors.Gray

    def on_export(self, _sender, _e):
        if self._export_data is None:
            self.status_label.Text = "Nothing to export — run Calculate first."
            self.status_label.TextColor = drawing.Colors.Orange
            return

        path = rs.SaveFileName(
            "Export Results to Excel",
            "Excel Files (*.xlsx)|*.xlsx||",
            None, "Lindero_Results", "xlsx"
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"

        try:
            export_to_excel(self._export_data, path)
            self.status_label.Text = f"Exported → {path}"
            self.status_label.TextColor = drawing.Colors.Green
        except Exception as ex:
            self.status_label.Text = f"Export failed: {ex}"
            self.status_label.TextColor = drawing.Colors.Red

    # ------------------------------------------------------------------
    # Per-scenario runners
    # ------------------------------------------------------------------

    def _run_s1(self, unit):
        name_key = self.name_key_combo.Text.strip()
        results = calc_s1(name_key)

        if not results:
            self.results_area.Text = "  No objects are currently selected in Rhino.\n"
            self.status_label.Text = "No selection."
            self.status_label.TextColor = drawing.Colors.Orange
            return

        self.results_area.Text = format_s1(results, unit)
        total = sum(r["area"] for r in results)
        self._export_data = {
            "scenario": 1, "unit": unit,
            "params": {"name_key": name_key},
            "objects": results,
        }
        self.status_label.Text = (
            f"S1  —  {len(results)} object(s)  |  Total: {total:,.4f} {unit}"
        )
        self.status_label.TextColor = drawing.Colors.Green

    def _run_s2(self, unit):
        layer_name = self._selected_layer(self.layer_s2_dd)
        if not layer_name:
            self.status_label.Text = "Please select a layer."
            self.status_label.TextColor = drawing.Colors.Red
            return

        obj_key = self.obj_key_s2.Text.strip()
        data = calc_s2(layer_name, obj_key)

        if not data["objects"]:
            self.results_area.Text = f"  No objects found on layer '{layer_name}'.\n"
            self.status_label.Text = f"Layer '{short_name(layer_name)}' is empty."
            self.status_label.TextColor = drawing.Colors.Orange
            return

        self.results_area.Text = format_s2(data, layer_name, obj_key, unit)
        self._export_data = {
            "scenario": 2, "unit": unit,
            "params": {"layer_name": layer_name, "obj_key": obj_key},
            **data,
        }
        self.status_label.Text = (
            f"S2  —  '{short_name(layer_name)}'  |  "
            f"{len(data['objects'])} object(s)  |  Total: {data['total']:,.4f} {unit}"
        )
        self.status_label.TextColor = drawing.Colors.Green

    def _run_s3(self, unit):
        parent = self._selected_layer(self.parent_layer_dd)
        if not parent:
            self.status_label.Text = "Please select a parent layer."
            self.status_label.TextColor = drawing.Colors.Red
            return

        obj_key = self.obj_key_s3.Text.strip()
        grp_key = self.grp_key_s3.Text.strip()
        data = calc_s3(parent, obj_key, grp_key)

        non_empty = [sl for sl, d in data["sublayers"].items() if d["objects"]]
        if not non_empty:
            self.results_area.Text = (
                f"  No objects found in any sublayer of '{parent}'.\n"
            )
            self.status_label.Text = "No objects found in sublayers."
            self.status_label.TextColor = drawing.Colors.Orange
            return

        self.results_area.Text = format_s3(data, parent, obj_key, grp_key, unit)
        self._export_data = {
            "scenario": 3, "unit": unit,
            "params": {"parent": parent, "obj_key": obj_key, "grp_key": grp_key},
            "sublayers": data["sublayers"],
            "overall_total": data["overall_total"],
        }
        self.status_label.Text = (
            f"S3  —  '{short_name(parent)}'  |  "
            f"{len(data['sublayers'])} sublayer(s)  |  "
            f"Overall total: {data['overall_total']:,.4f} {unit}"
        )
        self.status_label.TextColor = drawing.Colors.Green


# ══════════════════════════════════════════════════════════════════════════════
# Excel export
# ══════════════════════════════════════════════════════════════════════════════

# Styles -------------------------------------------------------------------

_HDR_FONT  = Font(bold=True, color="FFFFFF")
_HDR_FILL  = PatternFill(fill_type="solid", fgColor="2F5496")   # dark blue
_HDR_ALIGN = Alignment(horizontal="center", vertical="center")

_SEC_FONT  = Font(bold=True, color="1F3864")
_SEC_FILL  = PatternFill(fill_type="solid", fgColor="D9E1F2")   # light blue

_TOT_FONT  = Font(bold=True)
_TOT_FILL  = PatternFill(fill_type="solid", fgColor="F2F2F2")   # light grey

_AREA_FMT  = "#,##0.0000"


def _hdr(ws, row, cols):
    """Write a styled header row. Returns the row index + 1."""
    for ci, text in enumerate(cols, 1):
        c = ws.cell(row=row, column=ci, value=text)
        c.font  = _HDR_FONT
        c.fill  = _HDR_FILL
        c.alignment = _HDR_ALIGN
    return row + 1


def _sec(ws, row, text, n_cols=2):
    """Write a section-label row (sublayer / table title)."""
    c = ws.cell(row=row, column=1, value=text)
    c.font = _SEC_FONT
    c.fill = _SEC_FILL
    for ci in range(2, n_cols + 1):
        ws.cell(row=row, column=ci).fill = _SEC_FILL
    return row + 1


def _tot(ws, row, label, value):
    """Write a bold total row with area formatting."""
    lc = ws.cell(row=row, column=1, value=label)
    lc.font = _TOT_FONT
    lc.fill = _TOT_FILL
    vc = ws.cell(row=row, column=2, value=value)
    vc.font = _TOT_FONT
    vc.fill = _TOT_FILL
    vc.number_format = _AREA_FMT
    return row + 1


def _col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# Per-scenario writers -----------------------------------------------------

def _xl_s1(wb, data, unit):
    name_col = data["params"]["name_key"] or "Object Name"

    # Objects sheet
    ws = wb.create_sheet("Objects")
    _col_widths(ws, [38, 32, 20])
    row = _hdr(ws, 1, ["GUID", name_col, f"Footprint Area ({unit})"])
    for obj in data["objects"]:
        ws.cell(row, 1, obj["guid"])
        ws.cell(row, 2, obj["name"])
        ws.cell(row, 3, obj["area"]).number_format = _AREA_FMT
        row += 1

    # Summary sheet
    ws2 = wb.create_sheet("Summary")
    _col_widths(ws2, [36, 22])
    row = _hdr(ws2, 1, ["S1 — Selected Objects", f"[{unit}]"])
    pairs = [
        ("Name Key used",          data["params"]["name_key"] or "(object name / GUID)"),
        ("Object count",           len(data["objects"])),
        ("Total footprint area",   sum(o["area"] for o in data["objects"])),
    ]
    for label, value in pairs:
        lc = ws2.cell(row, 1, label)
        lc.font = Font(bold=True)
        vc = ws2.cell(row, 2, value)
        if isinstance(value, float):
            vc.number_format = _AREA_FMT
        row += 1


def _xl_s2(wb, data, unit):
    obj_col   = data["params"]["obj_key"] or "Object Name"
    layer     = data["params"]["layer_name"]
    raw_sum   = sum(o["area"] for o in data["objects"])

    # Objects sheet
    ws = wb.create_sheet("Objects")
    _col_widths(ws, [38, 30, 32, 20])
    row = _hdr(ws, 1, ["GUID", "Layer", obj_col, f"Footprint Area ({unit})"])
    for obj in data["objects"]:
        ws.cell(row, 1, obj["guid"])
        ws.cell(row, 2, layer)
        ws.cell(row, 3, obj["name"])
        ws.cell(row, 4, obj["area"]).number_format = _AREA_FMT
        row += 1

    # Summary sheet
    ws2 = wb.create_sheet("Summary")
    _col_widths(ws2, [38, 22])
    row = _hdr(ws2, 1, ["S2 — By Layer", f"[{unit}]"])
    kv_rows = [
        ("Layer",                              layer),
        ("Object Key used",                    data["params"]["obj_key"] or "(object name / GUID)"),
        ("Object count",                       len(data["objects"])),
        (f"Sum of individual areas ({unit})",  raw_sum),
        (f"Combined footprint ({unit})",       data["total"]),
        ("Boolean Union succeeded",            "Yes" if data["union_ok"] else "No — sum shown"),
    ]
    for label, value in kv_rows:
        lc = ws2.cell(row, 1, label)
        lc.font = Font(bold=True)
        vc = ws2.cell(row, 2, value)
        if isinstance(value, float):
            vc.number_format = _AREA_FMT
        row += 1


def _xl_s3(wb, data, unit):
    params    = data["params"]
    obj_col   = params["obj_key"] or "Object Name"
    grp_col   = params["grp_key"] or "Group"
    parent    = params["parent"]
    has_grp   = bool(params["grp_key"])

    # Objects sheet — one row per object, all levels flat
    ws = wb.create_sheet("Objects")
    _col_widths(ws, [38, 28, 28, 30, 28, 20])
    row = _hdr(ws, 1, ["GUID", "Parent Layer", "Level", obj_col, grp_col, f"Footprint Area ({unit})"])
    for sl, sl_data in data["sublayers"].items():
        level = short_name(sl)
        for obj in sl_data["objects"]:
            ws.cell(row, 1, obj["guid"])
            ws.cell(row, 2, parent)
            ws.cell(row, 3, level)
            ws.cell(row, 4, obj["name"])
            ws.cell(row, 5, obj["group"])
            ws.cell(row, 6, obj["area"]).number_format = _AREA_FMT
            row += 1

    # Summary sheet
    ws2 = wb.create_sheet("Summary")
    _col_widths(ws2, [30, 28, 22])
    row = _hdr(ws2, 1, ["S3 — Layer Hierarchy", "", f"[{unit}]"])

    # Parameters block
    for label, value in [
        ("Parent Layer",  parent),
        ("Object Key",    params["obj_key"] or "(object name / GUID)"),
        ("Group Key",     params["grp_key"] or "—"),
    ]:
        ws2.cell(row, 1, label).font = Font(bold=True)
        ws2.cell(row, 2, value)
        row += 1
    row += 1  # blank separator

    # By-level table
    row = _sec(ws2, row, "Combined Footprint by Level", n_cols=2)
    ws2.cell(row, 1, "Level").font      = Font(bold=True)
    ws2.cell(row, 2, f"Area ({unit})").font = Font(bold=True)
    row += 1
    for sl, sl_data in data["sublayers"].items():
        note = "" if sl_data["union_ok"] else " [union failed]"
        ws2.cell(row, 1, short_name(sl) + note)
        ws2.cell(row, 2, sl_data["total"]).number_format = _AREA_FMT
        row += 1
    row = _tot(ws2, row, "Grand Total", data["overall_total"])
    row += 1  # blank separator

    # By-group table (only if a group key was specified)
    if has_grp:
        row = _sec(ws2, row, f"Combined Footprint by Level and {grp_col}", n_cols=3)
        ws2.cell(row, 1, "Level").font        = Font(bold=True)
        ws2.cell(row, 2, grp_col).font        = Font(bold=True)
        ws2.cell(row, 3, f"Area ({unit})").font = Font(bold=True)
        row += 1
        for sl, sl_data in data["sublayers"].items():
            level = short_name(sl)
            for gv, ga in sorted(sl_data["group_totals"].items()):
                ws2.cell(row, 1, level)
                ws2.cell(row, 2, gv)
                ws2.cell(row, 3, ga).number_format = _AREA_FMT
                row += 1


def export_to_excel(data, filepath):
    """Write calculation results to an Excel workbook with Objects + Summary sheets."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # drop the default empty sheet

    s = data["scenario"]
    unit = data["unit"]

    if s == 1:
        _xl_s1(wb, data, unit)
    elif s == 2:
        _xl_s2(wb, data, unit)
    else:
        _xl_s3(wb, data, unit)

    wb.save(filepath)


# ══════════════════════════════════════════════════════════════════════════════
# Entry point
# ══════════════════════════════════════════════════════════════════════════════

def main():
    form = LinderoForm()
    form.Show()


if __name__ == "__main__":
    main()
