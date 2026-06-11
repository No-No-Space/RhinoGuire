#! python3
# -*- coding: utf-8 -*-
# r: openpyxl
# __title__ = "CutFillReport"
# __doc__ = """Version = 0.1
# Date    = 2026-06-10
# Author: Aquelon - aquelon@pm.me
# _____________________________________________________________________
# Description:
# Compares an original vs. a modified terrain (or reads the last PadGrader /
# WayGrader result), computes cut & fill volumes, shows KPIs and charts
# (cut-vs-fill, depth distribution, and a mass-haul curve for ways), builds a
# cut/fill-tinted map mesh with a legend, and exports the numbers to Excel and
# the charts to PNG.
# The window is modeless — Rhino stays accessible while it is open.
# _____________________________________________________________________
# How-to:
# -> Run in Rhino 8 (RunPythonScript). The CutFillReport panel opens.
# -> Source "From last grading" reads the most recent Pad/WayGrader result.
# -> Source "Two terrains" lets you pick Original + Modified (+ optional
#    analysis boundary) and a cell size, then Compute.
# -> Review KPIs/charts, "Show cut/fill map" to tint a mesh, then export.
# _____________________________________________________________________
# Last update:
# - [10.06.2026] - 0.1 Initial release
# _____________________________________________________________________

import os
import math
import Rhino
import Rhino.Geometry as rg
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
from TerrainTools._core import terrain as _terrain;    _il.reload(_terrain)
from TerrainTools._core import grading as _grading;    _il.reload(_grading)
from TerrainTools._core import volumes as _volumes;    _il.reload(_volumes)
from TerrainTools._core import meshbuild as _meshbuild; _il.reload(_meshbuild)
from TerrainTools._core import report as _report;      _il.reload(_report)
from TerrainTools import _widgets as _w;               _il.reload(_w)


_TERRAIN_FILTER = rs.filter.mesh | rs.filter.surface | rs.filter.polysurface
try:
    _TERRAIN_FILTER |= rs.filter.subd
except AttributeError:
    pass

_CUT_COLOR = _t.CHART_BAR        # teal-blue
_FILL_COLOR = _t.CHART_HIGH      # salmon
_NEUTRAL = drawing.Color.FromArgb(190, 184, 176)
_MAP_LAYER = "TerrainTools::CutFillMap"


# ===========================================================================
# Chart painting (shared by the on-screen Drawable and the PNG exporter)
# ===========================================================================

def paint_report(g, width, summary, stations, unit):
    """Paint the cut/fill dashboard onto Graphics *g*. Returns the height used."""
    pad = 12
    vu = ("%s³" % unit) if unit else ""
    y = pad

    g.DrawText(_t.F_SANS_B, _t.CHART_LABEL, drawing.PointF(pad, y),
               "Cut vs Fill %s" % vu)
    y += 20

    cut = summary["cut_volume"]
    fill = summary["fill_volume"]
    mx = max(cut, fill, 1e-9)
    bar_w = max(width - 2 * pad - 130, 40)
    bar_h = 16
    for label, val, color in (("Cut", cut, _CUT_COLOR), ("Fill", fill, _FILL_COLOR)):
        g.DrawText(_t.F_SANS_S, _t.CHART_LABEL, drawing.PointF(pad, y + 1), label)
        w = bar_w * (val / mx)
        g.FillRectangle(color, drawing.RectangleF(pad + 36, y, float(w), bar_h))
        g.DrawText(_t.F_SANS_S, _t.CHART_LABEL,
                   drawing.PointF(pad + 36 + bar_w + 6, y + 1), "{:,.1f}".format(val))
        y += bar_h + 6

    bal = summary["balance_ratio"]
    bal_s = "∞" if bal == float("inf") else "{:.2f}".format(bal)
    g.DrawText(_t.F_SANS_S, _t.CHART_DELTA, drawing.PointF(pad, y),
               "Net: {:,.1f}    Balance (fill/cut): {}".format(summary["net"], bal_s))
    y += 26

    # Depth distribution histogram
    g.DrawText(_t.F_SANS_B, _t.CHART_LABEL, drawing.PointF(pad, y),
               "Depth distribution (%s)" % (unit or "Δ"))
    y += 20
    hist = summary.get("depth_histogram") or []
    hh = 70
    base = y + hh
    if hist:
        maxc = max([c for _, _, c in hist] or [1]) or 1
        n = len(hist)
        bw = (width - 2 * pad) / float(max(n, 1))
        for k, (lo, hi, c) in enumerate(hist):
            h = hh * (c / float(maxc))
            if hi <= 0:
                color = _CUT_COLOR
            elif lo >= 0:
                color = _FILL_COLOR
            else:
                color = _NEUTRAL
            x = pad + k * bw
            g.FillRectangle(color, drawing.RectangleF(float(x + 1), float(base - h),
                                                      float(bw - 2), float(h)))
        g.DrawLine(drawing.Pen(_t.CHART_TOL, 1.0),
                   drawing.PointF(pad, base), drawing.PointF(width - pad, base))
        g.DrawText(_t.F_SANS_S, _t.TEXT_MUTED, drawing.PointF(pad, base + 2),
                   "{:.2f}".format(summary["min_delta"]))
        g.DrawText(_t.F_SANS_S, _t.TEXT_MUTED, drawing.PointF(width - pad - 44, base + 2),
                   "{:.2f}".format(summary["max_delta"]))
    y = base + 24

    # Mass-haul (cumulative net) for ways
    if stations and len(stations) >= 2:
        g.DrawText(_t.F_SANS_B, _t.CHART_LABEL, drawing.PointF(pad, y),
                   "Mass haul — cumulative net %s" % vu)
        y += 20
        ch = 80
        xs = [s["station"] for s in stations]
        ys = [s["cum_net"] for s in stations]
        x_min, x_max = min(xs), max(xs)
        y_min, y_max = min(ys + [0.0]), max(ys + [0.0])
        x_span = (x_max - x_min) or 1.0
        y_span = (y_max - y_min) or 1.0
        plot_w = width - 2 * pad
        zero_y = y + ch * (y_max / y_span)
        g.DrawLine(drawing.Pen(_t.CHART_TOL, 1.0),
                   drawing.PointF(pad, zero_y), drawing.PointF(width - pad, zero_y))
        pen = drawing.Pen(_CUT_COLOR, 2.0)
        prev = None
        for sx, sy in zip(xs, ys):
            px = pad + plot_w * ((sx - x_min) / x_span)
            py = y + ch * ((y_max - sy) / y_span)
            if prev is not None:
                g.DrawLine(pen, drawing.PointF(prev[0], prev[1]), drawing.PointF(px, py))
            prev = (px, py)
        y += ch + 8

    return int(y + pad)


# ===========================================================================
# Two-terrain comparison -> GradeResult
# ===========================================================================

def grade_from_two(orig, modi, cell, boundary):
    """Sample original (terrain) and modified (design) onto a shared grid."""
    if boundary is not None:
        x0, y0, nx, ny, cell = _terrain.make_grid_bounds(boundary, 0.0, cell)
    else:
        bo, bm = orig.bbox, modi.bbox
        minx = max(bo.Min.X, bm.Min.X)
        miny = max(bo.Min.Y, bm.Min.Y)
        maxx = min(bo.Max.X, bm.Max.X)
        maxy = min(bo.Max.Y, bm.Max.Y)
        if maxx <= minx or maxy <= miny:
            raise ValueError("The two terrains do not overlap in plan.")
        x0, y0 = minx, miny
        nx = max(int(math.ceil((maxx - minx) / cell)) + 1, 2)
        ny = max(int(math.ceil((maxy - miny) / cell)) + 1, 2)

    res = _grading.GradeResult(x0, y0, cell, nx, ny, kind="compare")
    tol = 1e-4
    for j in range(ny):
        y = y0 + j * cell
        for i in range(nx):
            x = x0 + i * cell
            zt = orig.project_z(x, y)
            zd = modi.project_z(x, y)
            res.z_terrain[j][i] = zt
            res.z_design[j][i] = zd
            if boundary is not None:
                if not _terrain.point_in_curve(boundary, rg.Point3d(x, y, 0.0), 1e-6):
                    continue
            if zt is not None and zd is not None:
                if boundary is not None or abs(zd - zt) > tol:
                    res.region_mask[j][i] = True
    return res


# ===========================================================================
# UI
# ===========================================================================

class CutFillReportForm(forms.Form):

    def __init__(self):
        super().__init__()
        self.Title = "CutFillReport — Compare · Quantify · Export"
        self.Resizable = True
        self.Padding = drawing.Padding(12)
        self.BackgroundColor = _t.BG
        self.MinimumSize = drawing.Size(470, 640)
        self.ClientSize = drawing.Size(520, 760)
        self.Owner = Rhino.UI.RhinoEtoApp.MainWindow

        self.orig_id = None
        self.modi_id = None
        self.boundary_id = None
        self.grade = None
        self.summary = None
        self.stations = []
        self.unit = ""
        self._tint_scale = None

        self._build_ui()
        self._refresh_source_mode()

    # ------------------------------------------------------------------
    def _build_ui(self):
        L = forms.DynamicLayout()
        L.DefaultSpacing = drawing.Size(5, 5)

        L.AddRow(_t.lbl("Cut / Fill Report", _t.F_HEAD, _t.TEXT))
        L.AddRow(_t.hint("Compare original vs modified terrain; chart and export."))
        L.AddRow(None)

        # Source
        L.AddRow(_t.lbl("Source", _t.F_SANS_B, _t.TEXT))
        self.source_mode = forms.DropDown()
        self.source_mode.Items.Add("From last grading (Pad/WayGrader)")
        self.source_mode.Items.Add("Two terrains (Original vs Modified)")
        self.source_mode.SelectedIndex = 0
        self.source_mode.SelectedIndexChanged += lambda s, e: self._refresh_source_mode()
        L.AddRow(self.source_mode)

        self.sticky_info = _t.lbl("", _t.F_SANS_S, _t.TEXT_MUTED)
        L.AddRow(self.sticky_info)

        # Two-terrain controls
        self.orig_btn = _t.btn("Select Original Terrain")
        self.orig_btn.Click += self.on_select_orig
        self.orig_info = _t.lbl("—", _t.F_SANS_S, _t.TEXT_MUTED)
        self.modi_btn = _t.btn("Select Modified Terrain")
        self.modi_btn.Click += self.on_select_modi
        self.modi_info = _t.lbl("—", _t.F_SANS_S, _t.TEXT_MUTED)
        self.bound_btn = _t.btn("Analysis Boundary (optional)")
        self.bound_btn.Click += self.on_select_boundary
        self.bound_info = _t.lbl("—", _t.F_SANS_S, _t.TEXT_MUTED)
        self.cell_box = forms.TextBox(); self.cell_box.Text = "1.0"; self.cell_box.Width = 70
        self.cell_row = _w.labeled_row("Cell size:", self.cell_box, 70)

        for ctrl in (self.orig_btn, self.orig_info, self.modi_btn, self.modi_info,
                     self.bound_btn, self.bound_info, self.cell_row):
            L.AddRow(ctrl)
        L.AddRow(None)

        # Compute
        self.compute_btn = _t.btn("Compute", _t.BTN_CALC)
        self.compute_btn.Click += self.on_compute
        L.AddRow(self.compute_btn)
        L.AddRow(None)

        # KPI panel
        L.AddRow(_t.lbl("Results", _t.F_SANS_B, _t.TEXT))
        self.kpi_lbl = _t.lbl("Run Compute to see KPIs.", _t.F_SANS, _t.TEXT_MUTED)
        L.AddRow(self.kpi_lbl)
        L.AddRow(None)

        # Charts
        self.chart = forms.Drawable()
        self.chart.Size = drawing.Size(480, 280)
        self.chart.Paint += self._on_paint
        chart_scroll = forms.Scrollable()
        chart_scroll.ExpandContentWidth = True
        chart_scroll.ExpandContentHeight = False
        chart_scroll.Size = drawing.Size(490, 300)
        chart_scroll.Content = self.chart
        L.AddRow(chart_scroll)
        L.AddRow(None)

        # Map + legend
        self.map_btn = _t.btn("Show cut/fill map (tinted mesh)")
        self.map_btn.Enabled = False
        self.map_btn.Click += self.on_show_map
        L.AddRow(self.map_btn)
        self.legend_panel = forms.Panel()
        L.AddRow(self.legend_panel)
        L.AddRow(None)

        # Exports
        L.AddRow(_t.lbl("Export", _t.F_SANS_B, _t.TEXT))
        self.xlsx_btn = _t.btn("Export to Excel (.xlsx)")
        self.xlsx_btn.Enabled = False
        self.xlsx_btn.Click += self.on_export_xlsx
        self.png_btn = _t.btn("Export charts as PNG")
        self.png_btn.Enabled = False
        self.png_btn.Click += self.on_export_png
        self.cap_btn = _t.btn("Capture viewport as PNG…")
        self.cap_btn.Click += self.on_capture_viewport
        L.AddRow(self.xlsx_btn)
        L.AddRow(self.png_btn)
        L.AddRow(self.cap_btn)
        L.AddRow(None)

        self.status_lbl = _t.lbl("Ready.", _t.F_SANS, _t.TEXT_MUTED)
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

    def _refresh_source_mode(self):
        two = (self.source_mode.SelectedIndex == 1)
        for c in (self.orig_btn, self.orig_info, self.modi_btn, self.modi_info,
                  self.bound_btn, self.bound_info, self.cell_row):
            c.Visible = two
        if not two:
            data = sc.sticky.get("terraintools_last_grade")
            if data:
                self.sticky_info.Text = "✓ %s result available (unit: %s)" % (
                    data.get("kind", "?"), data.get("unit", "") or "—")
                self.sticky_info.TextColor = _t.TEXT_OK
            else:
                self.sticky_info.Text = "No grading found yet — run PadGrader or WayGrader."
                self.sticky_info.TextColor = _t.TEXT_WARN
        else:
            self.sticky_info.Text = ""

    # ------------------------------------------------------------------
    def on_select_orig(self, sender, e):
        oid = rs.GetObject("Select ORIGINAL terrain", _TERRAIN_FILTER, preselect=True)
        if oid:
            self.orig_id = oid
            self.orig_info.Text = "✓ %s" % (rs.ObjectName(oid) or "(unnamed)")
            self.orig_info.TextColor = _t.TEXT_OK

    def on_select_modi(self, sender, e):
        mid = rs.GetObject("Select MODIFIED terrain", _TERRAIN_FILTER, preselect=True)
        if mid:
            self.modi_id = mid
            self.modi_info.Text = "✓ %s" % (rs.ObjectName(mid) or "(unnamed)")
            self.modi_info.TextColor = _t.TEXT_OK

    def on_select_boundary(self, sender, e):
        bid = rs.GetObject("Select analysis boundary (closed curve)", rs.filter.curve, preselect=True)
        if bid:
            crv = rs.coercecurve(bid)
            closed, planar = _terrain.is_closed_planar(crv) if crv else (False, False)
            if closed and planar:
                self.boundary_id = bid
                self.bound_info.Text = "✓ boundary set"
                self.bound_info.TextColor = _t.TEXT_OK
            else:
                self.bound_info.Text = "✗ boundary must be closed & planar"
                self.bound_info.TextColor = _t.TEXT_ERROR

    # ------------------------------------------------------------------
    def on_compute(self, sender, e):
        self._status("Computing…")
        try:
            if self.source_mode.SelectedIndex == 0:
                data = sc.sticky.get("terraintools_last_grade")
                if not data:
                    self._status("No grading found — run PadGrader/WayGrader, or use Two terrains.",
                                 "error")
                    return
                self.grade = data["grade"]
                self.unit = data.get("unit", "")
                self.stations = data.get("stations") or _volumes.per_station(self.grade)
            else:
                if not (self.orig_id and self.modi_id):
                    self._status("Pick both Original and Modified terrains.", "warn")
                    return
                try:
                    cell = float(self.cell_box.Text)
                    if cell <= 0:
                        raise ValueError
                except ValueError:
                    self._status("Cell size must be a positive number.", "error")
                    return
                orig = _terrain.TerrainModel(self.orig_id)
                modi = _terrain.TerrainModel(self.modi_id)
                boundary = rs.coercecurve(self.boundary_id) if self.boundary_id else None
                self.grade = grade_from_two(orig, modi, cell, boundary)
                self.unit = _terrain.model_unit_label()
                self.stations = []

            self.summary = _volumes.cut_fill(self.grade)
        except Exception as ex:
            self._status("Compute failed: %s" % ex, "error")
            return

        self._update_kpis()
        self._resize_chart()
        self.chart.Invalidate()
        self.map_btn.Enabled = True
        self.xlsx_btn.Enabled = True
        self.png_btn.Enabled = True
        msg = "Computed."
        if self.grade.flags:
            msg += "  ⚠ " + self.grade.flags[0]
        self._status(msg, "ok")

    def _update_kpis(self):
        s = self.summary
        u = self.unit
        vu = ("%s³" % u) if u else "vol"
        au = ("%s²" % u) if u else "area"
        bal = "∞" if s["balance_ratio"] == float("inf") else "{:.3f}".format(s["balance_ratio"])
        self.kpi_lbl.Text = (
            "Cut:  {cut:,.2f} {vu}\n"
            "Fill: {fill:,.2f} {vu}\n"
            "Net:  {net:,.2f} {vu}\n"
            "Cut area: {ca:,.1f} {au}   Fill area: {fa:,.1f} {au}\n"
            "Balance (fill/cut): {bal}   Cell: {cs:g} {u}   Grid: {nx}×{ny}"
        ).format(cut=s["cut_volume"], fill=s["fill_volume"], net=s["net"], vu=vu,
                 ca=s["cut_area"], fa=s["fill_area"], au=au, bal=bal,
                 cs=s["cell_size"], u=u, nx=s["grid_nx"], ny=s["grid_ny"])
        self.kpi_lbl.TextColor = _t.TEXT

    def _resize_chart(self):
        h = 250
        if self.stations and len(self.stations) >= 2:
            h += 110
        self.chart.Size = drawing.Size(self.chart.Width if self.chart.Width > 0 else 480, h)

    def _on_paint(self, sender, e):
        g = e.Graphics
        w = self.chart.Width or 480
        g.FillRectangle(drawing.Colors.White, drawing.RectangleF(0, 0, float(w), float(self.chart.Height)))
        if self.summary is None:
            g.DrawText(_t.F_SANS, _t.TEXT_MUTED, drawing.PointF(12, 12), "No data yet — run Compute.")
            return
        paint_report(g, w, self.summary, self.stations, self.unit)

    # ------------------------------------------------------------------
    def on_show_map(self, sender, e):
        if self.grade is None:
            return
        try:
            mesh = _meshbuild.grid_to_mesh(self.grade, only_region=True)
            if mesh.Vertices.Count == 0:
                self._status("Nothing to tint (empty graded region).", "warn")
                return
            self._tint_scale = _meshbuild.tint_by_delta(mesh, self.grade, _meshbuild.default_ramp)
            _w.add_mesh_to_layer(mesh, _MAP_LAYER, name="CutFillMap")
            sc.doc.Views.Redraw()
            self._build_legend(self._tint_scale)
            self._status("Cut/fill map added to '%s'." % _MAP_LAYER, "ok")
        except Exception as ex:
            self._status("Map failed: %s" % ex, "error")

    def _build_legend(self, scale):
        u = self.unit
        lay = forms.DynamicLayout()
        lay.DefaultSpacing = drawing.Size(6, 3)
        lay.Padding = drawing.Padding(4)
        lay.AddRow(_t.lbl("Legend — cut/fill depth (%s)" % (u or "Δ"), _t.F_SANS_B, _t.TEXT))
        stops = [(-1.0, "cut −{:.2f}".format(scale)),
                 (-0.5, ""), (0.0, "0"), (0.5, ""),
                 (1.0, "fill +{:.2f}".format(scale))]
        for t, label in stops:
            sw = forms.Panel()
            sw.Size = drawing.Size(28, 16)
            c = _meshbuild.default_ramp(t)
            sw.BackgroundColor = drawing.Color.FromArgb(c.R, c.G, c.B)
            row = _w.labeled_row("", sw, 0)
            if label:
                row.Items.Add(forms.StackLayoutItem(_t.lbl(label, _t.F_SANS_S, _t.TEXT_MUTED)))
            lay.AddRow(row)
        self.legend_panel.Content = lay

    # ------------------------------------------------------------------
    def on_export_xlsx(self, sender, e):
        if self.summary is None:
            return
        path = rs.SaveFileName("Export cut/fill report",
                               "Excel Files (*.xlsx)|*.xlsx||",
                               _t.prefs_get("cutfill_export_xlsx"),
                               "CutFillReport", "xlsx")
        if not path:
            return
        _t.prefs_set("cutfill_export_xlsx", path)
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        try:
            per_station = self.stations or None
            meta = {"unit": self.unit, "Source": self.grade.kind}
            _report.write_xlsx(path, self.summary, per_station=per_station, meta=meta)
            self._status("Exported → %s" % path, "ok")
        except Exception as ex:
            self._status("Excel export failed: %s" % ex, "error")

    def on_export_png(self, sender, e):
        if self.summary is None:
            return
        path = rs.SaveFileName("Export charts as PNG",
                               "PNG Files (*.png)|*.png||",
                               _t.prefs_get("cutfill_export_png"),
                               "CutFillReport_Charts", "png")
        if not path:
            return
        _t.prefs_set("cutfill_export_png", path)
        if not path.lower().endswith(".png"):
            path += ".png"
        try:
            width = 640
            # measure height with a throwaway bitmap
            tmp = drawing.Bitmap(width, 10, drawing.PixelFormat.Format32bppRgba)
            tg = drawing.Graphics(tmp)
            height = paint_report(tg, width, self.summary, self.stations, self.unit)
            tg.Dispose(); tmp.Dispose()

            bmp = drawing.Bitmap(width, max(height, 20), drawing.PixelFormat.Format32bppRgba)
            g = drawing.Graphics(bmp)
            try:
                g.FillRectangle(drawing.Colors.White,
                                drawing.RectangleF(0, 0, float(width), float(height)))
                paint_report(g, width, self.summary, self.stations, self.unit)
            finally:
                g.Dispose()
            bmp.Save(path)
            bmp.Dispose()
            self._status("Charts saved → %s" % path, "ok")
        except Exception as ex:
            self._status("PNG export failed: %s" % ex, "error")

    def on_capture_viewport(self, sender, e):
        self._status("Opening Rhino's viewport capture…")
        Rhino.RhinoApp.RunScript("ViewCaptureToFile", False)
        self._status("Viewport capture dialog closed.", "info")


def main():
    CutFillReportForm().Show()


if __name__ == "__main__":
    main()
