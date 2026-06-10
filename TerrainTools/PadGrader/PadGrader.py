#! python3
# -*- coding: utf-8 -*-
# __title__ = "PadGrader"
# __doc__ = """Version = 0.1
# Date    = 2026-06-10
# Author: Aquelon - aquelon@pm.me
# _____________________________________________________________________
# Description:
# Places one or more building pads (closed planar boundaries at a target
# elevation) onto a terrain and grades cut/fill slopes around them down to the
# daylight line. Outputs a new graded mesh on a separate layer (the original
# terrain is never modified) and reports cut & fill volumes.
# Terrain may be a Mesh, SubD, Surface (incl. trimmed) or Polysurface.
# The window is modeless — Rhino stays accessible while it is open.
# _____________________________________________________________________
# How-to:
# -> Run in Rhino 8 (RunPythonScript). The PadGrader panel opens.
# -> 1: Select the terrain.  2: Select closed planar pad curve(s).
# -> 3: Choose the pad elevation (from terrain mean/min/max, or explicit Z).
# -> 4: Set the cut and fill grading slopes (H:V, %, or °).
# -> Set cell size / max reach / output layer, then click Generate.
# -> Use "Open in CutFillReport" to chart and export the result.
# _____________________________________________________________________
# Last update:
# - [10.06.2026] - 0.1 Initial release
# _____________________________________________________________________

import os
import Rhino
import Eto.Drawing as drawing
import Eto.Forms as forms
import rhinoscriptsyntax as rs
import scriptcontext as sc

import sys as _sys, os as _os
_rg_root = _os.path.normpath(_os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "..", ".."))
if _rg_root not in _sys.path:
    _sys.path.insert(0, _rg_root)

import importlib as _il
from ui import theme as _t; _il.reload(_t)
from TerrainTools._core import slope as _slope;       _il.reload(_slope)
from TerrainTools._core import terrain as _terrain;    _il.reload(_terrain)
from TerrainTools._core import grading as _grading;    _il.reload(_grading)
from TerrainTools._core import volumes as _volumes;    _il.reload(_volumes)
from TerrainTools._core import meshbuild as _meshbuild; _il.reload(_meshbuild)
from TerrainTools import _widgets as _w;               _il.reload(_w)


_TERRAIN_FILTER = rs.filter.mesh | rs.filter.surface | rs.filter.polysurface
try:
    _TERRAIN_FILTER |= rs.filter.subd
except AttributeError:
    pass

_PAD_Z_MODES = [
    ("From terrain (mean)", "mean"),
    ("From terrain (min)",  "min"),
    ("From terrain (max)",  "max"),
    ("Explicit Z",          "explicit"),
]


class PadGraderForm(forms.Form):

    def __init__(self):
        super().__init__()
        self.Title = "PadGrader — Building Pad Grading"
        self.Resizable = True
        self.Padding = drawing.Padding(12)
        self.BackgroundColor = _t.BG
        self.MinimumSize = drawing.Size(420, 560)
        self.ClientSize = drawing.Size(460, 660)
        self.Owner = Rhino.UI.RhinoEtoApp.MainWindow

        self.terrain_id = None
        self.terrain = None
        self.pad_ids = []
        self._cell_autofilled = False

        self._build_ui()

    # ------------------------------------------------------------------
    def _build_ui(self):
        L = forms.DynamicLayout()
        L.DefaultSpacing = drawing.Size(5, 5)

        L.AddRow(_t.lbl("Building Pad Grading", _t.F_HEAD, _t.TEXT))
        L.AddRow(_t.hint("Cut/fill a terrain around level pads, down to daylight."))
        L.AddRow(None)

        # 1 — Terrain
        L.AddRow(_t.lbl("1 — Terrain  (mesh, surface, polysurface…)", _t.F_SANS_B, _t.TEXT))
        self.terrain_btn = _t.btn("Select Terrain")
        self.terrain_btn.Click += self.on_select_terrain
        L.AddRow(self.terrain_btn)
        self.terrain_info = _t.lbl("No terrain selected.", _t.F_SANS, _t.TEXT_MUTED)
        L.AddRow(self.terrain_info)
        L.AddRow(None)

        # 2 — Pad boundaries
        L.AddRow(_t.lbl("2 — Pad boundary(ies)  (closed planar curves)", _t.F_SANS_B, _t.TEXT))
        self.pad_btn = _t.btn("Select Pad Curve(s)")
        self.pad_btn.Enabled = False
        self.pad_btn.Click += self.on_select_pads
        L.AddRow(self.pad_btn)
        self.pad_info = _t.lbl("No pad curves selected.", _t.F_SANS, _t.TEXT_MUTED)
        L.AddRow(self.pad_info)
        L.AddRow(None)

        # 3 — Pad elevation
        L.AddRow(_t.lbl("3 — Pad elevation", _t.F_SANS_B, _t.TEXT))
        self.elev_mode = forms.DropDown()
        for text, _key in _PAD_Z_MODES:
            self.elev_mode.Items.Add(text)
        self.elev_mode.SelectedIndex = 0
        self.elev_mode.SelectedIndexChanged += self.on_elev_mode_changed
        self.elev_z = forms.TextBox()
        self.elev_z.Text = "0.0"
        self.elev_z.Width = 70
        self.elev_z.Enabled = False
        L.AddRow(_w.labeled_row("Mode:", self.elev_mode, 70))
        L.AddRow(_w.labeled_row("Explicit Z:", self.elev_z, 70))
        L.AddRow(None)

        # 4 — Grading slopes
        L.AddRow(_t.lbl("4 — Grading slopes  (H:V is run:rise, e.g. 2:1)", _t.F_SANS_B, _t.TEXT))
        self.cut_slope = _w.SlopeInput("Cut:", "2", "ratio_hv")
        self.fill_slope = _w.SlopeInput("Fill:", "2", "ratio_hv")
        L.AddRow(self.cut_slope.control)
        L.AddRow(self.fill_slope.control)
        L.AddRow(None)

        # Options
        L.AddRow(_t.lbl("Options", _t.F_SANS_B, _t.TEXT))
        self.cell_box = forms.TextBox()
        self.cell_box.Text = "1.0"
        self.cell_box.Width = 70
        self.reach_box = forms.TextBox()
        self.reach_box.Text = "50"
        self.reach_box.Width = 70
        self.layer_box = forms.TextBox()
        self.layer_box.Text = "TerrainTools::Graded"
        self.layer_box.Width = 180
        L.AddRow(_w.labeled_row("Cell size:", self.cell_box, 90))
        L.AddRow(_w.labeled_row("Max reach:", self.reach_box, 90))
        L.AddRow(_w.labeled_row("Output layer:", self.layer_box, 90))
        L.AddRow(None)

        # Generate / Close
        self.gen_btn = _t.btn("Generate", _t.BTN_CALC)
        self.gen_btn.Enabled = False
        self.gen_btn.Click += self.on_generate
        L.AddRow(self.gen_btn)

        # Results panel
        self.results_lbl = _t.lbl("", _t.F_SANS, _t.TEXT)
        L.AddRow(self.results_lbl)

        self.report_btn = _t.btn("Open in CutFillReport")
        self.report_btn.Enabled = False
        self.report_btn.Click += self.on_open_report
        L.AddRow(self.report_btn)

        L.AddRow(None)
        self.status_lbl = _t.lbl("Ready — select a terrain to begin.", _t.F_SANS, _t.TEXT_MUTED)
        L.AddRow(self.status_lbl)

        close_btn = _t.btn("Close", _t.BTN_CLEAR)
        close_btn.Click += lambda s, e: self.Close()
        L.AddRow(close_btn)

        scroll = forms.Scrollable()
        scroll.Border = getattr(forms.BorderType, "None")
        scroll.Content = L
        self.Content = scroll

    # ------------------------------------------------------------------
    def _status(self, text, state=None):
        self.status_lbl.Text = text
        self.status_lbl.TextColor = _t.status_color(state) if state else _t.TEXT_MUTED

    def _update_gen(self):
        self.gen_btn.Enabled = (self.terrain_id is not None) and len(self.pad_ids) > 0

    def on_elev_mode_changed(self, *_):
        key = _PAD_Z_MODES[self.elev_mode.SelectedIndex][1]
        self.elev_z.Enabled = (key == "explicit")

    # ------------------------------------------------------------------
    def on_select_terrain(self, sender, e):
        self._status("Select the terrain in the viewport…")
        obj = rs.GetObject("Select terrain (mesh, surface, polysurface, SubD)",
                           _TERRAIN_FILTER, preselect=True)
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
        self.terrain_info.Text = "✓ %s  [%s]" % (name, _terrain.obj_type_label(obj))
        self.terrain_info.TextColor = _t.TEXT_OK
        self.pad_btn.Enabled = True
        if not self._cell_autofilled:
            self.cell_box.Text = "%.3g" % _w.default_cell_size(self.terrain)
            self._cell_autofilled = True
        self._status("Terrain set. Now select pad curve(s).", "info")
        self._update_gen()

    def on_select_pads(self, sender, e):
        self._status("Select closed planar pad curve(s)…")
        ids = rs.GetObjects("Select pad boundary curve(s)", rs.filter.curve, preselect=True)
        if not ids:
            self._status("Pad selection cancelled.", "warn")
            return
        bad = []
        good = []
        for cid in ids:
            crv = rs.coercecurve(cid)
            if crv is None:
                continue
            closed, planar = _terrain.is_closed_planar(crv)
            if closed and planar:
                good.append(cid)
            else:
                bad.append(cid)
        self.pad_ids = good
        if not good:
            self.pad_info.Text = "✗ No closed planar curves among the selection."
            self.pad_info.TextColor = _t.TEXT_ERROR
        else:
            msg = "✓ %d pad curve(s)" % len(good)
            if bad:
                msg += "  (%d skipped: open/non-planar)" % len(bad)
            self.pad_info.Text = msg
            self.pad_info.TextColor = _t.TEXT_WARN if bad else _t.TEXT_OK
        self._update_gen()
        self._status("Pad curves ready." if good else "Need at least one closed planar curve.",
                     "info" if good else "warn")

    # ------------------------------------------------------------------
    def _read_params(self):
        m_cut = self.cut_slope.gradient()
        m_fill = self.fill_slope.gradient()
        if m_cut is None or m_fill is None:
            raise ValueError("Cut/Fill slope value is invalid.")
        try:
            cell = float(self.cell_box.Text)
            if cell <= 0:
                raise ValueError
        except ValueError:
            raise ValueError("Cell size must be a positive number.")
        try:
            reach = float(self.reach_box.Text)
            if reach <= 0:
                raise ValueError
        except ValueError:
            raise ValueError("Max reach must be a positive number.")

        key = _PAD_Z_MODES[self.elev_mode.SelectedIndex][1]
        if key == "explicit":
            try:
                pad_mode = float(self.elev_z.Text)
            except ValueError:
                raise ValueError("Explicit Z must be a number.")
        else:
            pad_mode = key
        return m_cut, m_fill, cell, reach, pad_mode

    def on_generate(self, sender, e):
        self.gen_btn.Enabled = False
        try:
            m_cut, m_fill, cell, reach, pad_mode = self._read_params()
        except ValueError as ex:
            self._status(str(ex), "error")
            self.gen_btn.Enabled = True
            return

        self._status("Grading… (sampling terrain and building the heightfield)")
        curves = [rs.coercecurve(c) for c in self.pad_ids]
        try:
            grade = _grading.grade_pads(self.terrain, curves, pad_mode,
                                        m_cut, m_fill, cell, reach)
            mesh = _meshbuild.grid_to_mesh(grade, only_region=True)
            if mesh.Vertices.Count == 0:
                self._status("Grading produced no mesh — check slopes / reach / elevation.", "warn")
                self.gen_btn.Enabled = True
                return
            layer = self.layer_box.Text.strip() or "TerrainTools::Graded"
            mesh_id = _w.add_mesh_to_layer(mesh, layer, name="PadGraded")
            sc.doc.Views.Redraw()
            kpi = _volumes.cut_fill(grade)
        except Exception as ex:
            self._status("Grading failed: %s" % ex, "error")
            self.gen_btn.Enabled = True
            return

        unit = _terrain.model_unit_label()
        sc.sticky["terraintools_last_grade"] = {
            "grade": grade, "terrain_id": self.terrain_id, "mesh_id": mesh_id,
            "kind": "pad", "unit": unit, "kpi": kpi,
        }

        vu = ("%s³" % unit) if unit else "vol"
        self.results_lbl.Text = (
            "Cut %s · Fill %s · Net %s %s\n"
            "Balance (fill/cut): %s · %d cells graded"
            % (_fmt(kpi["cut_volume"]), _fmt(kpi["fill_volume"]),
               _fmt(kpi["net"]), vu, _bal(kpi["balance_ratio"]), kpi["n_cells"])
        )
        self.results_lbl.TextColor = _t.TEXT
        self.report_btn.Enabled = True

        msg = "Done — graded mesh added to '%s'." % layer
        if grade.flags:
            msg += "  ⚠ " + grade.flags[0]
        self._status(msg, "ok")
        self.gen_btn.Enabled = True

    def on_open_report(self, sender, e):
        path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                            "CutFillReport", "CutFillReport.py")
        if not os.path.exists(path):
            self._status("CutFillReport.py not found.", "error")
            return
        exec(open(path, encoding="utf-8").read(),
             {"__file__": path, "__name__": "__main__"})


def _fmt(v):
    return "{:,.2f}".format(v)


def _bal(b):
    return "∞" if b == float("inf") else "{:.2f}".format(b)


def main():
    PadGraderForm().Show()


if __name__ == "__main__":
    main()
