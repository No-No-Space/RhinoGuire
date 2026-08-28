#! python3
# -*- coding: utf-8 -*-
# __title__ = "Trocha"
# __doc__ = """Version = 0.1
# Date    = 2026-07-15
# Author: Aquelon - aquelon@pm.me
# _____________________________________________________________________
# Description:
# Drapes a solid road slab onto a terrain from a user-drawn centerline: a
# clean, crease-free top face with real downward thickness, buried deep
# enough into the terrain that it never "flies" over a dip. Non-destructive
# (the terrain is never modified) and create<->update linked to the
# centerline via RG_ROAD_* user strings, so re-running after you edit the
# curve (or the terrain moves) replaces the slab instead of duplicating it.
# Distinct from TerrainTools/WayGrader, which grades/modifies the terrain
# itself for earthwork design - Trocha only drapes a presentation object on
# top of it. See RoadTool/road_tool_plan.md for the full design spec.
# The window is modeless - Rhino stays accessible while it is open.
# _____________________________________________________________________
# How-to:
# -> Run in Rhino 8 (RunPythonScript). The Trocha panel opens.
# -> 1: Select the terrain (Mesh / Surface / Polysurface / SubD).
# -> 2: Pick a centerline curve. Picking an already-tagged curve loads its
#       stored width/thickness (update mode); an untagged curve starts fresh.
# -> Set width / thickness / sample spacing, click Generate.
# -> Adjust parameters and Regenerate on the same centerline as needed.
# -> Picking more than one curve at once (viewport multi-select, or window/
#    crossing select before clicking Pick/Re-pick) switches to batch mode:
#    Generate Selected drapes a new road on every picked curve using the
#    current Width/Thickness/Sample spacing fields; Update Selected re-drapes
#    just the tagged ones among them with each curve's own stored
#    width/thickness; Merge Boolean-unions just their built roads into one
#    solid (deliberately scoped to the pick, not the whole document -
#    unioning every road in a large file is slow even though most never
#    touch).
# -> Remove clears the current road; Update All re-drapes every tagged road
#    in the document (run after the terrain changes).
# _____________________________________________________________________
# Last update:
# - [15.07.2026] - 0.1 Initial release
# _____________________________________________________________________

import System
import Rhino
import Eto.Drawing as drawing
import Eto.Forms as forms
import rhinoscriptsyntax as rs
import scriptcontext as sc

import sys as _sys, os as _os
_rg_root = _os.path.normpath(_os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))
if _rg_root not in _sys.path:
    _sys.path.insert(0, _rg_root)

import importlib as _il
from ui import theme as _t; _il.reload(_t)
from TerrainTools._core import terrain as _terrain; _il.reload(_terrain)
from RoadTool._core import config as _config;       _il.reload(_config)
from RoadTool._core import geometry as _geometry;   _il.reload(_geometry)
from RoadTool._core import state as _state;         _il.reload(_state)
from RoadTool._core import junctions as _junctions; _il.reload(_junctions)


_TERRAIN_FILTER = rs.filter.mesh | rs.filter.surface | rs.filter.polysurface
try:
    _TERRAIN_FILTER |= rs.filter.subd
except AttributeError:
    pass


# ---------------------------------------------------------------------------
# Local doc helpers (this tool's own layer/AddBrep glue - Sebucan-style, kept
# in the tool file rather than a shared module since Trocha is the only
# consumer; see road_tool_plan.md S11).
# ---------------------------------------------------------------------------

def _ensure_layer(full_name):
    parts = full_name.split("::")
    path, parent = "", ""
    for part in parts:
        path = part if not path else (path + "::" + part)
        if not rs.IsLayer(path):
            rs.AddLayer(part, parent=parent if parent else None)
        parent = path
    return full_name


def _add_brep_to_layer(brep, layer_name, name=None):
    _ensure_layer(layer_name)
    obj_id = sc.doc.Objects.AddBrep(brep)
    if obj_id and obj_id != System.Guid.Empty:
        rs.ObjectLayer(obj_id, layer_name)
        if name:
            rs.ObjectName(obj_id, name)
    return obj_id


def _positive(text, label):
    try:
        v = float(text)
        if v <= 0:
            raise ValueError
        return v
    except ValueError:
        raise ValueError("%s must be a positive number." % label)


# ---------------------------------------------------------------------------
# Window
# ---------------------------------------------------------------------------

class TrochaForm(forms.Form):

    def __init__(self):
        super().__init__()
        self.Title = "Trocha — Road on Terrain"
        self.Resizable = True
        self.Padding = drawing.Padding(12)
        self.BackgroundColor = _t.BG
        self.MinimumSize = drawing.Size(410, 520)
        self.ClientSize = drawing.Size(440, 600)
        self.Owner = Rhino.UI.RhinoEtoApp.MainWindow

        self.terrain_id = None
        self.terrain = None
        self.center_id = None
        self.center_ids = []

        self._build_ui()

    # ------------------------------------------------------------------
    def _build_ui(self):
        L = forms.StackLayout()
        L.Orientation = forms.Orientation.Vertical
        L.HorizontalContentAlignment = forms.HorizontalAlignment.Stretch
        L.Spacing = 5

        def _add(ctrl):
            L.Items.Add(forms.StackLayoutItem(ctrl))

        def _gap():
            sp = forms.Panel()
            sp.Height = 6
            _add(sp)

        _add(_t.lbl("Road on Terrain", _t.F_HEAD, _t.TEXT))
        _add(_t.hint("Drape a solid road slab on the terrain from a centerline."))
        _gap()

        # 1 - Terrain
        _add(_t.lbl("1 — Terrain", _t.F_SANS_B, _t.TEXT))
        self.terrain_btn = _t.btn("Select Terrain")
        self.terrain_btn.Click += self.on_select_terrain
        _add(self.terrain_btn)
        self.terrain_info = _t.lbl("No terrain selected.", _t.F_SANS, _t.TEXT_MUTED)
        _add(self.terrain_info)
        _gap()

        # 2 - Centerline
        _add(_t.lbl("2 — Centerline", _t.F_SANS_B, _t.TEXT))
        self.center_btn = _t.btn("Pick / Re-pick Centerline")
        self.center_btn.Enabled = False
        self.center_btn.Click += self.on_select_center
        _add(self.center_btn)
        self.center_info = _t.lbl("No centerline selected.", _t.F_SANS, _t.TEXT_MUTED)
        _add(self.center_info)
        _gap()

        # 3 - Parameters
        _add(_t.lbl("3 — Parameters", _t.F_SANS_B, _t.TEXT))
        self.width_box = forms.TextBox(); self.width_box.Text = repr(_config.DEFAULT_WIDTH); self.width_box.Width = 70
        self.thick_box = forms.TextBox(); self.thick_box.Text = repr(_config.DEFAULT_THICKNESS); self.thick_box.Width = 70
        self.step_box = forms.TextBox(); self.step_box.Text = repr(_config.DEFAULT_SAMPLE_STEP); self.step_box.Width = 70
        _add(_t.lbl("Width:", _t.F_SANS, _t.TEXT))
        _add(self.width_box)
        _add(_t.lbl("Thickness:", _t.F_SANS, _t.TEXT))
        _add(self.thick_box)
        _add(_t.lbl("Sample spacing:", _t.F_SANS, _t.TEXT))
        _add(self.step_box)
        self.smooth_check = forms.CheckBox()
        self.smooth_check.Text = "Smooth centerline before draping"
        self.smooth_check.Checked = True
        _add(self.smooth_check)
        _gap()

        # Advanced (contact-rule overrides) - blank = auto from sample spacing.
        _add(_t.lbl("Advanced (blank = auto)", _t.F_SANS_B, _t.TEXT))
        self.fit_tol_box = forms.TextBox(); self.fit_tol_box.Width = 70
        self.top_rise_box = forms.TextBox(); self.top_rise_box.Width = 70
        self.margin_box = forms.TextBox(); self.margin_box.Width = 70
        _add(_t.lbl("Fit tolerance:", _t.F_SANS, _t.TEXT))
        _add(self.fit_tol_box)
        _add(_t.lbl("Top rise:", _t.F_SANS, _t.TEXT))
        _add(self.top_rise_box)
        _add(_t.lbl("Margin:", _t.F_SANS, _t.TEXT))
        _add(self.margin_box)
        _gap()

        self.gen_btn = _t.btn("Generate", _t.BTN_CALC)
        self.gen_btn.Enabled = False
        self.gen_btn.Click += self.on_generate
        _add(self.gen_btn)

        _gap()
        divider = forms.Panel()
        divider.Height = 1
        divider.BackgroundColor = _t.TEXT_MUTED
        _add(divider)
        _gap()

        _add(_t.lbl("4 — Other actions", _t.F_SANS_B, _t.TEXT))

        self.remove_btn = _t.btn("Remove")
        self.remove_btn.Enabled = False
        self.remove_btn.Click += self.on_remove
        self.generate_selected_btn = _t.btn("Generate Selected")
        self.generate_selected_btn.Enabled = False
        self.generate_selected_btn.Click += self.on_generate_selected
        self.update_selected_btn = _t.btn("Update Selected")
        self.update_selected_btn.Enabled = False
        self.update_selected_btn.Click += self.on_update_selected
        self.update_all_btn = _t.btn("Update All")
        self.update_all_btn.Click += self.on_update_all
        self.merge_btn = _t.btn("Merge")
        self.merge_btn.Enabled = False
        self.merge_btn.Click += self.on_merge

        actions_table = forms.TableLayout()
        actions_table.Spacing = drawing.Size(8, 6)
        for button, explain in (
            (self.remove_btn, "Delete the currently picked road and un-tag its centerline."),
            (self.generate_selected_btn, "Drape a new road on each centerline picked above, using "
                                          "the Width/Thickness/Sample spacing set above (tagged "
                                          "ones are replaced)."),
            (self.update_selected_btn, "Re-drape only the tagged centerlines picked above, each "
                                        "with its own stored width/thickness."),
            (self.update_all_btn, "Re-drape every tagged road in the whole document — "
                                   "use after the terrain changes."),
            (self.merge_btn, "Boolean-union only the roads of the centerlines picked above "
                              "into one solid (pick several first — merging the whole document "
                              "at once is slow)."),
        ):
            # Fixed widths (not auto-sized to the unwrapped text) so the label
            # wraps in place instead of forcing the whole window wider.
            button.Width = 120
            explain_lbl = _t.hint(explain)
            explain_lbl.Width = 250
            explain_lbl.Wrap = forms.WrapMode.Word
            actions_table.Rows.Add(forms.TableRow(
                forms.TableCell(button), forms.TableCell(explain_lbl, True)))
        _add(actions_table)

        _gap()
        self.status_lbl = _t.lbl("Ready — select a terrain to begin.", _t.F_SANS, _t.TEXT_MUTED)
        _add(self.status_lbl)

        close_btn = _t.btn("Close", _t.BTN_CLEAR)
        close_btn.Click += lambda s, e: self.Close()
        _add(close_btn)

        scroll = forms.Scrollable()
        try:
            scroll.Border = getattr(forms.BorderType, "None")
        except AttributeError:
            pass
        scroll.Content = L
        self.Content = scroll

    # ------------------------------------------------------------------
    def _status(self, text, state=None):
        self.status_lbl.Text = text
        self.status_lbl.TextColor = _t.status_color(state) if state else _t.TEXT_MUTED

    def _update_buttons(self):
        self.gen_btn.Enabled = (self.terrain_id is not None) and (self.center_id is not None)
        self.gen_btn.Text = "Regenerate" if (self.center_id and _state.is_road_centerline(self.center_id)) else "Generate"
        self.remove_btn.Enabled = bool(self.center_id and _state.is_road_centerline(self.center_id))
        self.generate_selected_btn.Enabled = bool(self.center_ids) and (self.terrain_id is not None)
        self.update_selected_btn.Enabled = bool(self.center_ids)
        self.merge_btn.Enabled = bool(self.center_ids)

    # ------------------------------------------------------------------
    def on_select_terrain(self, sender, e):
        self._status("Select the terrain in the viewport…")
        obj = rs.GetObject("Select terrain", _TERRAIN_FILTER, preselect=True)
        if obj is None:
            self._status("Terrain selection cancelled.", "warn")
            return
        try:
            self.terrain = _terrain.TerrainModel(obj)
        except Exception as ex:
            self._status("Could not read terrain: %s" % ex, "error")
            return
        self.terrain_id = obj
        name = rs.ObjectName(obj) or "(unnamed)"
        self.terrain_info.Text = "Terrain: %s  [%s]" % (name, _terrain.obj_type_label(obj))
        self.terrain_info.TextColor = _t.TEXT_OK
        self.center_btn.Enabled = True
        self._status("Terrain set. Now pick a centerline.", "info")
        self._update_buttons()

    def on_select_center(self, sender, e):
        self._status("Pick the centerline curve(s)…")
        cids = rs.GetObjects("Select centerline(s) (tagged = update, untagged = create)",
                              rs.filter.curve, preselect=True)
        if not cids:
            self._status("Centerline selection cancelled.", "warn")
            return

        if len(cids) > 1:
            self.center_id = None
            self.center_ids = list(cids)
            tagged_count = sum(1 for c in cids if _state.is_road_centerline(c))
            self.center_info.Text = "%d centerlines selected (%d tagged)." % (len(cids), tagged_count)
            self.center_info.TextColor = _t.TEXT_OK
            self._status("Multiple centerlines selected — Generate Selected builds a road on each, "
                         "Update Selected re-drapes the tagged ones with their own stored params.", "info")
            self._update_buttons()
            return

        self.center_ids = []
        cid = cids[0]
        curve = rs.coercecurve(cid)
        if curve is None:
            self._status("Could not read that curve.", "error")
            return

        self.center_id = cid
        if _state.is_road_centerline(cid):
            width, thick, _terrain_id = _state.read_params(cid)
            if width:
                self.width_box.Text = repr(width)
            if thick:
                self.thick_box.Text = repr(thick)
            self.center_info.Text = "Centerline: tagged road (update mode)"
            self.center_info.TextColor = _t.TEXT_OK
            self._status("Loaded stored width/thickness. Adjust and Regenerate as needed.", "info")
        else:
            # Re-link fallback (road_tool_plan.md S5, rev.2 #7): the picked
            # curve is untagged, but there may be an orphaned slab nearby
            # whose original centerline was replaced. Original width/
            # thickness died with that centerline, so re-linking just clears
            # the orphan and starts fresh rather than guessing lost params.
            orphan = _state.find_orphan_near(curve)
            if orphan:
                answer = rs.MessageBox(
                    "An orphaned road slab was found near this curve (its original "
                    "centerline is gone).\n\nDelete it and rebuild here?",
                    4 | 32, "Trocha")
                if answer == 6:
                    rs.DeleteObject(orphan)
                    self._status("Orphan slab cleared. Set width/thickness and Generate.", "info")
                else:
                    self._status("New centerline. Set width/thickness and Generate.", "info")
            else:
                self._status("New centerline. Set width/thickness and Generate.", "info")
            self.center_info.Text = "Centerline: untagged (create mode)"
            self.center_info.TextColor = _t.TEXT_OK
        self._update_buttons()

    # ------------------------------------------------------------------
    def _read_optional_positive(self, text, label):
        """Blank -> None (auto); otherwise must be a positive number."""
        text = (text or "").strip()
        if not text:
            return None
        return _positive(text, label)

    def _read_config(self):
        step = _positive(self.step_box.Text, "Sample spacing")
        return _config.TrochaConfig(
            tolerance=sc.doc.ModelAbsoluteTolerance,
            sample_step=step,
            fit_tol=self._read_optional_positive(self.fit_tol_box.Text, "Fit tolerance"),
            top_rise=self._read_optional_positive(self.top_rise_box.Text, "Top rise"),
            margin=self._read_optional_positive(self.margin_box.Text, "Margin"),
            smooth_center=self.smooth_check.Checked,
        )

    def on_generate(self, sender, e):
        if self.terrain_id is None or self.center_id is None:
            self._status("Select a terrain and a centerline first.", "warn")
            return
        try:
            width = _positive(self.width_box.Text, "Width")
            thickness = _positive(self.thick_box.Text, "Thickness")
            cfg = self._read_config()
        except ValueError as ex:
            self._status(str(ex), "error")
            return

        curve = rs.coercecurve(self.center_id)
        if curve is None:
            self._status("Could not read the centerline.", "error")
            return
        if not rs.IsObject(self.terrain_id):
            self._status("The selected terrain no longer exists — select it again.", "error")
            return

        self._status("Draping road on terrain…")
        result = _geometry.build_slab(curve, self.terrain, width, thickness, cfg)
        if result.brep is None:
            self._status(result.flags[-1] if result.flags else "Slab build failed.", "error")
            return

        record = sc.doc.BeginUndoRecord("Trocha: build road")
        try:
            old_child = _state.resolve_child(self.center_id)
            if old_child and rs.IsObject(old_child):
                rs.DeleteObject(old_child)
            slab_id = _add_brep_to_layer(result.brep, _config.DEFAULT_LAYER, name="Trocha road")
            _state.write_centerline_tags(self.center_id, width, thickness,
                                          str(self.terrain_id), str(slab_id))
            _state.write_slab_tag(slab_id, self.center_id)
        finally:
            sc.doc.EndUndoRecord(record)
        sc.doc.Views.Redraw()

        kind = "solid" if result.brep.IsSolid else "surface only — see warning"
        msg = "Road built (%s, thickness %.4g)." % (kind, result.thickness or thickness)
        if result.flags:
            msg += "  ⚠ " + result.flags[0]
            self._status(msg, "warn")
        else:
            self._status(msg, "ok")
        self.center_info.Text = "Centerline: tagged road (update mode)"
        self.center_info.TextColor = _t.TEXT_OK
        self._update_buttons()

    def on_remove(self, sender, e):
        if not self.center_id or not _state.is_road_centerline(self.center_id):
            self._status("Pick a tagged road centerline first.", "warn")
            return
        record = sc.doc.BeginUndoRecord("Trocha: remove road")
        try:
            _state.remove_road(self.center_id)
        finally:
            sc.doc.EndUndoRecord(record)
        sc.doc.Views.Redraw()
        self.center_info.Text = "Centerline: untagged (create mode)"
        self.center_info.TextColor = _t.TEXT_MUTED
        self._status("Road removed.", "ok")
        self._update_buttons()

    def on_generate_selected(self, sender, e):
        """Drape a fresh road on each picked centerline (road_tool_plan.md create model),
        using the current Width/Thickness/Sample spacing fields for all of them — unlike
        Update Selected, which re-uses each centerline's own already-stored params."""
        if not self.center_ids:
            self._status("Pick two or more centerlines first (Pick / Re-pick Centerline).", "warn")
            return
        if self.terrain_id is None:
            self._status("Select a terrain first.", "warn")
            return
        try:
            width = _positive(self.width_box.Text, "Width")
            thickness = _positive(self.thick_box.Text, "Thickness")
            cfg = self._read_config()
        except ValueError as ex:
            self._status(str(ex), "error")
            return
        if not rs.IsObject(self.terrain_id):
            self._status("The selected terrain no longer exists — select it again.", "error")
            return

        ids = [cid for cid in self.center_ids if rs.IsObject(cid)]
        if not ids:
            self._status("None of the picked centerlines still exist.", "warn")
            return

        record = sc.doc.BeginUndoRecord("Trocha: generate selected roads")
        built, skipped = 0, 0
        try:
            for cid in ids:
                curve = rs.coercecurve(cid)
                if curve is None:
                    skipped += 1
                    continue
                result = _geometry.build_slab(curve, self.terrain, width, thickness, cfg)
                if result.brep is None:
                    skipped += 1
                    continue
                old_child = _state.resolve_child(cid)
                if old_child and rs.IsObject(old_child):
                    rs.DeleteObject(old_child)
                slab_id = _add_brep_to_layer(result.brep, _config.DEFAULT_LAYER, name="Trocha road")
                _state.write_centerline_tags(cid, width, thickness, str(self.terrain_id), str(slab_id))
                _state.write_slab_tag(slab_id, cid)
                built += 1
        finally:
            sc.doc.EndUndoRecord(record)
        sc.doc.Views.Redraw()
        self._status("Generated %d road(s), skipped %d." % (built, skipped),
                      "ok" if built else "warn")
        self._update_buttons()

    def _regenerate_centerlines(self, centerlines):
        """Re-drape each of *centerlines* using its own stored width/thickness/
        terrain (road_tool_plan.md create<->update model). Returns (updated, skipped)."""
        updated, skipped = 0, 0
        for cid in centerlines:
            curve = rs.coercecurve(cid)
            width, thickness, terrain_id = _state.read_params(cid)
            if curve is None or width is None or thickness is None \
                    or not terrain_id or not rs.IsObject(terrain_id):
                skipped += 1
                continue
            try:
                terrain_model = _terrain.TerrainModel(terrain_id)
            except Exception:
                skipped += 1
                continue
            cfg = _config.TrochaConfig(tolerance=sc.doc.ModelAbsoluteTolerance)
            result = _geometry.build_slab(curve, terrain_model, width, thickness, cfg)
            if result.brep is None:
                skipped += 1
                continue
            old_child = _state.resolve_child(cid)
            if old_child and rs.IsObject(old_child):
                rs.DeleteObject(old_child)
            slab_id = _add_brep_to_layer(result.brep, _config.DEFAULT_LAYER, name="Trocha road")
            _state.write_centerline_tags(cid, width, thickness, str(terrain_id), str(slab_id))
            _state.write_slab_tag(slab_id, cid)
            updated += 1
        return updated, skipped

    def on_update_selected(self, sender, e):
        if not self.center_ids:
            self._status("Select two or more centerlines first (Pick / Re-pick Centerline).", "warn")
            return
        live = [cid for cid in self.center_ids if rs.IsObject(cid)]
        tagged = [cid for cid in live if _state.is_road_centerline(cid)]
        if not tagged:
            self._status("None of the selected centerlines are tagged roads.", "warn")
            return
        record = sc.doc.BeginUndoRecord("Trocha: update selected roads")
        try:
            updated, skipped = self._regenerate_centerlines(tagged)
        finally:
            sc.doc.EndUndoRecord(record)
        sc.doc.Views.Redraw()
        skipped += len(self.center_ids) - len(tagged)
        self._status("Updated %d road(s), skipped %d." % (updated, skipped),
                      "ok" if updated else "warn")

    def on_update_all(self, sender, e):
        centerlines = _state.all_tagged_centerlines()
        if not centerlines:
            self._status("No tagged roads found in the document.", "warn")
            return
        record = sc.doc.BeginUndoRecord("Trocha: update all roads")
        try:
            updated, skipped = self._regenerate_centerlines(centerlines)
        finally:
            sc.doc.EndUndoRecord(record)
        sc.doc.Views.Redraw()
        self._status("Updated %d road(s), skipped %d." % (updated, skipped),
                      "ok" if updated else "warn")

    def on_merge(self, sender, e):
        # Scoped to the picked centerlines (not the whole document): a Boolean
        # union over every tagged road in a large file is heavy even though
        # most of those roads never touch each other. Picking the handful
        # that actually need merging keeps CreateBooleanUnion's input small.
        if not self.center_ids:
            self._status("Pick two or more centerlines first (Pick / Re-pick Centerline), then Merge.", "warn")
            return
        tagged = [cid for cid in self.center_ids if rs.IsObject(cid) and _state.is_road_centerline(cid)]
        slab_ids = []
        for cid in tagged:
            sid = _state.resolve_child(cid)
            if sid and rs.IsObject(sid):
                slab_ids.append(sid)
        if len(slab_ids) < 2:
            self._status("Need at least two built roads among the picked centerlines to merge.", "warn")
            return
        breps = [b for b in (rs.coercebrep(sid) for sid in slab_ids) if b is not None]

        record = sc.doc.BeginUndoRecord("Trocha: merge selected roads")
        try:
            merged, ok = _junctions.merge_slabs(breps, sc.doc.ModelAbsoluteTolerance)
            if not ok:
                self._status("Boolean union failed — roads left unmerged.", "warn")
                return
            source_ids = set(str(cid) for cid in tagged)
            for oid in rs.AllObjects():
                if rs.GetUserText(oid, _config.TAG_MERGED) != "1":
                    continue
                prior_sources = (rs.GetUserText(oid, _config.TAG_MERGE_SOURCES) or "").split("|")
                if source_ids.intersection(prior_sources):
                    rs.DeleteObject(oid)
            sources_tag = "|".join(sorted(source_ids))
            for brep in merged:
                mid = _add_brep_to_layer(brep, _config.MERGED_LAYER, name="Trocha merged road")
                rs.SetUserText(mid, _config.TAG_MERGED, "1")
                rs.SetUserText(mid, _config.TAG_MERGE_SOURCES, sources_tag)
        finally:
            sc.doc.EndUndoRecord(record)
        sc.doc.Views.Redraw()
        self._status("Merged %d road(s) into %d solid(s)." % (len(slab_ids), len(merged)), "ok")


def main():
    TrochaForm().Show()


if __name__ == "__main__":
    main()
