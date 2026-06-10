#! python3
# -*- coding: utf-8 -*-
"""Shared Eto widgets + Rhino document helpers for the TerrainTools front-ends.

This module *does* import Eto (unlike _core); it is tool-side glue used by
PadGrader / WayGrader / CutFillReport so their slope inputs, layer handling and
status helpers stay consistent and DRY.
"""

import Eto.Drawing as drawing
import Eto.Forms as forms
import rhinoscriptsyntax as rs
import scriptcontext as sc
import System

import sys as _sys, os as _os
_rg_root = _os.path.normpath(_os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))
if _rg_root not in _sys.path:
    _sys.path.insert(0, _rg_root)
from ui import theme as _t
from TerrainTools._core import slope as _slope


# Stable order for the unit dropdown <-> slope.UNITS mapping.
_UNIT_KEYS = ("ratio_hv", "percent", "degrees")
_UNIT_TEXT = {"ratio_hv": "H:V", "percent": "%", "degrees": "°"}


class SlopeInput(object):
    """A horizontal slope row: label · textbox · unit dropdown · live equivalents.

    Use :meth:`gradient` to read the canonical gradient ``m = V/H`` (or None if
    the field is invalid). Add :attr:`control` to a layout.
    """

    def __init__(self, label, default_value="2", default_unit="ratio_hv", width_label=70):
        self._lbl = _t.lbl(label, _t.F_SANS, _t.TEXT)
        self._lbl.Width = width_label

        self._box = forms.TextBox()
        self._box.Text = str(default_value)
        self._box.Width = 60

        self._unit = forms.DropDown()
        for k in _UNIT_KEYS:
            self._unit.Items.Add(_UNIT_TEXT[k])
        self._unit.SelectedIndex = _UNIT_KEYS.index(default_unit)
        self._unit.Width = 60

        self._equiv = _t.lbl("", _t.F_SANS_S, _t.TEXT_MUTED)

        self._box.TextChanged += self._refresh
        self._unit.SelectedIndexChanged += self._refresh

        row = forms.StackLayout()
        row.Orientation = forms.Orientation.Horizontal
        row.Spacing = 6
        row.VerticalContentAlignment = forms.VerticalAlignment.Center
        for w in (self._lbl, self._box, self._unit, self._equiv):
            row.Items.Add(forms.StackLayoutItem(w))
        self.control = row
        self._refresh(None, None)

    def unit_key(self):
        return _UNIT_KEYS[self._unit.SelectedIndex]

    def gradient(self):
        """Canonical gradient m (>0), or None if the entry is invalid."""
        try:
            return _slope.to_gradient(self._box.Text.strip(), self.unit_key())
        except (ValueError, Exception):
            return None

    def _refresh(self, _s, _e):
        m = self.gradient()
        if m is None:
            self._equiv.Text = "—"
            self._equiv.TextColor = _t.TEXT_ERROR
        else:
            self._equiv.Text = "= " + _slope.describe(m)
            self._equiv.TextColor = _t.TEXT_MUTED


def labeled_row(label, control, width_label=70):
    """label · control on one horizontal line."""
    lbl = _t.lbl(label, _t.F_SANS, _t.TEXT)
    lbl.Width = width_label
    row = forms.StackLayout()
    row.Orientation = forms.Orientation.Horizontal
    row.Spacing = 6
    row.VerticalContentAlignment = forms.VerticalAlignment.Center
    row.Items.Add(forms.StackLayoutItem(lbl))
    row.Items.Add(forms.StackLayoutItem(control))
    return row


# ---------------------------------------------------------------------------
# Document helpers
# ---------------------------------------------------------------------------

def ensure_layer(full_name):
    """Ensure a (possibly nested, '::'-separated) layer exists; return its name."""
    parts = full_name.split("::")
    path = ""
    parent = ""
    for part in parts:
        path = part if not path else (path + "::" + part)
        if not rs.IsLayer(path):
            rs.AddLayer(part, parent=parent if parent else None)
        parent = path
    return full_name


def add_mesh_to_layer(mesh, layer_name, name=None):
    """Add a mesh to the document on *layer_name*; return its object id."""
    ensure_layer(layer_name)
    obj_id = sc.doc.Objects.AddMesh(mesh)
    if obj_id != System.Guid.Empty:
        rs.ObjectLayer(obj_id, layer_name)
        if name:
            rs.ObjectName(obj_id, name)
    return obj_id


def default_cell_size(terrain_model, divisions=200, min_cell=0.25):
    """A sensible default grid cell ~ bbox diagonal / divisions, clamped."""
    bb = terrain_model.bbox
    dx = bb.Max.X - bb.Min.X
    dy = bb.Max.Y - bb.Min.Y
    diag = (dx * dx + dy * dy) ** 0.5
    cell = diag / float(divisions) if diag > 0 else min_cell
    return max(cell, min_cell)
