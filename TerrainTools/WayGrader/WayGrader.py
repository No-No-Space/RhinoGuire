#! python3
# -*- coding: utf-8 -*-
# __title__ = "WayGrader"
# __doc__ = """Version = 0.1
# Date    = 2026-06-10
# Author: Aquelon - aquelon@pm.me
# _____________________________________________________________________
# Description:
# Grades a way / path corridor onto a terrain from its centerline. A persistent
# window holds the corridor parameters (width, crossfall, cut/fill slopes); edit
# them and Regenerate without re-picking. The carriageway is draped on the
# terrain (v1) with a crown or single crossfall; cut/fill skirts run to daylight
# each side. Outputs a new graded mesh on a separate layer (terrain untouched).
# The window is modeless — Rhino stays accessible while it is open.
# _____________________________________________________________________
# How-to:
# -> Run in Rhino 8 (RunPythonScript). The WayGrader panel opens.
# -> 1: Select the terrain.  2: Pick the centerline polyline/curve.
# -> Set width, crossfall (crown/single), cut & fill slopes, station/cell/reach.
# -> Click Generate. Adjust parameters and Regenerate as needed.
# -> Use "Open in CutFillReport" for charts, a mass-haul curve and Excel export.
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

_CROSSFALL_MODES = [("Crown (falls both ways)", "crown"),
                    ("Single crossfall (L→R)", "single")]


class WayGraderForm(forms.Form):

    def __init__(self):
        super().__init__()
        self.Title = "WayGrader — Way / Path Corridor Grading"
        self.Resizable = True
        self.Padding = drawing.Padding(12)
        self.BackgroundColor = _t.BG
        self.MinimumSize = drawing.Size(430, 600)
        self.ClientSize = drawing.Size(470, 700)
        self.Owner = Rhino.UI.RhinoEtoApp.MainWindow

        self.terrain_id = None
        self.terrain = None
        self.center_id = None
        self.last_mesh_id = None
        self._cell_autofilled = False

        self._build_ui()

    # ------------------------------------------------------------------
    def _build_ui(self):
        L = forms.DynamicLayout()
        L.DefaultSpacing = drawing.Size(5, 5)

        L.AddRow(_t.lbl("Way / Path Corridor Grading", _t.F_HEAD, _t.TEXT))
        L.AddRow(_t.hint("Drape a corridor on the terrain, grade skirts to daylight."))
        L.AddRow(None)

        # 1 — Terrain
        L.AddRow(_t.lbl("1 — Terrain", _t.F_SANS_B, _t.TEXT))
        self.terrain_btn = _t.btn("Select Terrain")
        self.terrain_btn.Click += self.on_select_terrain
        L.AddRow(self.terrain_btn)
        self.terrain_info = _t.lbl("No terrain selected.", _t.F_SANS, _t.TEXT_MUTED)
        L.AddRow(self.terrain_info)
        L.AddRow(None)

        # 2 — Centerline
        L.AddRow(_t.lbl("2 — Centerline  (polyline / curve)", _t.F_SANS_B, _t.TEXT))
        self.center_btn = _t.btn("Pick / Re-pick Centerline")
        self.center_btn.Enabled = False
        self.center_btn.Click += self.on_select_center
        L.AddRow(self.center_btn)
        self.center_info = _t.lbl("No centerline selected.", _t.F_SANS, _t.TEXT_MUTED)
        L.AddRow(self.center_info)
        L.AddRow(None)

        # Parameters
        L.AddRow(_t.lbl("Corridor parameters", _t.F_SANS_B, _t.TEXT))
        self.width_box = forms.TextBox()
        self.width_box.Text = "4.0"
        self.width_box.Width = 70
        L.AddRow(_w.labeled_row("Width:", self.width_box, 95))

        self.crossfall_mode = forms.DropDown()
        for text, _k in _CROSSFALL_MODES:
            self.crossfall_mode.Items.Add(text)
        self.crossfall_mode.SelectedIndex = 0
        L.AddRow(_w.labeled_row("Crossfall:", self.crossfall_mode, 95))
        self.cross_slope = _w.SlopeInput("Cross slope:", "2", "percent", width_label=95)
        L.AddRow(self.cross_slope.control)

        self.cut_slope = _w.SlopeInput("Cut:", "2", "ratio_hv", width_label=95)
        self.fill_slope = _w.SlopeInput("Fill:", "2", "ratio_hv", width_label=95)
        L.AddRow(self.cut_slope.control)
        L.AddRow(self.fill_slope.control)
        L.AddRow(None)

        # Options
        L.AddRow(_t.lbl("Options", _t.F_SANS_B, _t.TEXT))
        self.station_box = forms.TextBox(); self.station_box.Text = "2.0"; self.station_box.Width = 70
        self.cell_box = forms.TextBox(); self.cell_box.Text = "1.0"; self.cell_box.Width = 70
        self.reach_box = forms.TextBox(); self.reach_box.Text = "30"; self.reach_box.Width = 70
        self.layer_box = forms.TextBox(); self.layer_box.Text = "TerrainTools::Graded"; self.layer_box.Width = 180
        L.AddRow(_w.labeled_row("Station spacing:", self.station_box, 110))
        L.AddRow(_w.labeled_row("Cell / cross step:", self.cell_box, 110))
        L.AddRow(_w.labeled_row("Max reach:", self.reach_box, 110))
        L.AddRow(_w.labeled_row("Output layer:", self.layer_box, 110))
        self.replace_check = forms.CheckBox()
        self.replace_check.Text = "Replace previous result on Regenerate"
        self.replace_check.Checked = True
        L.AddRow(self.replace_check)
        L.AddRow(None)

        # Generate
        self.gen_btn = _t.btn("Generate", _t.BTN_CALC)
        self.gen_btn.Enabled = False
        self.gen_btn.Click += self.on_generate
        L.AddRow(self.gen_btn)

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

    def _update_gen(self):
        self.gen_btn.Enabled = (self.terrain_id is not None) and (self.center_id is not None)
        self.gen_btn.Text = "Regenerate" if self.last_mesh_id else "Generate"

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
        self.terrain_info.Text = "✓ %s  [%s]" % (name, _terrain.obj_type_label(obj))
        self.terrain_info.TextColor = _t.TEXT_OK
        self.center_btn.Enabled = True
        if not self._cell_autofilled:
            self.cell_box.Text = "%.3g" % _w.default_cell_size(self.terrain)
            self._cell_autofilled = True
        self._status("Terrain set. Now pick the centerline.", "info")
        self._update_gen()

    def on_select_center(self, sender, e):
        self._status("Pick the centerline curve…")
        cid = rs.GetObject("Select centerline (polyline / curve)", rs.filter.curve, preselect=True)
        if cid is None:
            self._status("Centerline selection cancelled.", "warn")
            return
        crv = rs.coercecurve(cid)
        if crv is None:
            self._status("Could not read that curve.", "error")
            return
        self.center_id = cid
        length = crv.GetLength()
        self.center_info.Text = "✓ centerline  (length %.2f)" % length
        self.center_info.TextColor = _t.TEXT_OK
        self._status("Centerline set. Set parameters and Generate.", "info")
        self._update_gen()

    # ------------------------------------------------------------------
    def _read_params(self):
        def pos(text, label):
            try:
                v = float(text)
                if v <= 0:
                    raise ValueError
                return v
            except ValueError:
                raise ValueError("%s must be a positive number." % label)

        width = pos(self.width_box.Text, "Width")
        cell = pos(self.cell_box.Text, "Cell size")
        station = pos(self.station_box.Text, "Station spacing")
        reach = pos(self.reach_box.Text, "Max reach")
        m_cross = self.cross_slope.gradient()
        m_cut = self.cut_slope.gradient()
        m_fill = self.fill_slope.gradient()
        if m_cross is None or m_cut is None or m_fill is None:
            raise ValueError("A slope value is invalid.")
        crossfall = _CROSSFALL_MODES[self.crossfall_mode.SelectedIndex][1]
        params = _grading.WayParams(width=width, m_cross=m_cross, crossfall=crossfall,
                                    m_cut=m_cut, m_fill=m_fill, cell=cell,
                                    station=station, max_reach=reach)
        return params

    def on_generate(self, sender, e):
        self.gen_btn.Enabled = False
        try:
            params = self._read_params()
        except ValueError as ex:
            self._status(str(ex), "error")
            self.gen_btn.Enabled = True
            return

        self._status("Grading corridor…")
        centerline = rs.coercecurve(self.center_id)
        try:
            grade = _grading.grade_corridor(self.terrain, centerline, params)
            mesh = _meshbuild.grid_to_mesh(grade, only_region=True)
            if mesh.Vertices.Count == 0:
                self._status("Corridor produced no mesh — check width / reach / slopes.", "warn")
                self.gen_btn.Enabled = True
                return
            if self.replace_check.Checked and self.last_mesh_id:
                rs.DeleteObject(self.last_mesh_id)
                self.last_mesh_id = None
            layer = self.layer_box.Text.strip() or "TerrainTools::Graded"
            mesh_id = _w.add_mesh_to_layer(mesh, layer, name="WayGraded")
            self.last_mesh_id = mesh_id
            sc.doc.Views.Redraw()
            kpi = _volumes.cut_fill(grade)
            stations = _volumes.per_station(grade)
        except Exception as ex:
            self._status("Grading failed: %s" % ex, "error")
            self.gen_btn.Enabled = True
            return

        unit = _terrain.model_unit_label()
        sc.sticky["terraintools_last_grade"] = {
            "grade": grade, "terrain_id": self.terrain_id, "mesh_id": mesh_id,
            "kind": "corridor", "unit": unit, "kpi": kpi, "stations": stations,
        }

        vu = ("%s³" % unit) if unit else "vol"
        self.results_lbl.Text = (
            "Cut %s · Fill %s · Net %s %s\n"
            "Balance (fill/cut): %s · %d stations"
            % ("{:,.2f}".format(kpi["cut_volume"]), "{:,.2f}".format(kpi["fill_volume"]),
               "{:,.2f}".format(kpi["net"]), vu,
               ("∞" if kpi["balance_ratio"] == float("inf") else "{:.2f}".format(kpi["balance_ratio"])),
               len(stations))
        )
        self.results_lbl.TextColor = _t.TEXT
        self.report_btn.Enabled = True

        msg = "Done — corridor mesh added to '%s'." % layer
        if grade.flags:
            msg += "  ⚠ " + grade.flags[0]
        self._status(msg, "ok")
        self._update_gen()
        self.gen_btn.Enabled = True

    def on_open_report(self, sender, e):
        path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                            "CutFillReport", "CutFillReport.py")
        if not os.path.exists(path):
            self._status("CutFillReport.py not found.", "error")
            return
        exec(open(path, encoding="utf-8").read(),
             {"__file__": path, "__name__": "__main__"})


def main():
    WayGraderForm().Show()


if __name__ == "__main__":
    main()
