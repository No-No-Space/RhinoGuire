#! python3
# -*- coding: utf-8 -*-
# __title__ = "Pregonero"
# __doc__ = """Version = 0.1
# Date    = 2026-06-30
# Author: Aquelon - aquelon@pm.me
# _____________________________________________________________________
# Description:
# Tag objects with leaders whose text is built from a user-defined
# template. Placeholders {Key} are inserted as LIVE Rhino text fields
# (%<UserText("guid","Key")>%) so every leader reflects the tagged
# object's own user text and updates automatically when it changes.
# Works like Revit's tag-by-category tool: pick a template object,
# write a template referencing its keys, then click objects to tag.
# Missing keys are created on the tagged object with the value "TBD".
# Modeless window — Rhino stays accessible while it is open.
# _____________________________________________________________________
# How-to:
# -> Run in Rhino 8 (RunPythonScript). The Pregonero window opens.
# -> Click "Pick template object" and select an object carrying the
#    user-text keys you want to tag with. Its keys are listed.
# -> Build the leader template in the text box. Insert a key with the
#    "Insert {key}" button or type {KeyName} yourself. Plain text and
#    line breaks are kept verbatim, e.g.:  Room: {RoomName} / {Level}
# -> Pick a Dimension Style (controls font + arrow) and, optionally, a
#    text-height override.
# -> Click "Start Tagging". For each object: first click ON the object
#    (the leader arrow lands at that point), then click where the text
#    should go. Repeat for as many objects as you want.
# -> Press Enter (or Esc) in the viewport to finish and return here.
# _____________________________________________________________________
# Notes:
# - Leader text uses live text fields, so editing an object's user text
#   (e.g. replacing TBD with a real value) updates every tag that
#   references it.
# - Each tagging session is a single Undo step.
# - The %<UserText("guid","Key")>% field syntax is documented at
#   docs.mcneel.com/rhino/8/help/en-us/information/text_fields.htm
# _____________________________________________________________________


import re
import System
from System.Collections.Generic import List
import rhinoscriptsyntax as rs
import Rhino
import Rhino.Geometry as rg
import scriptcontext as sc
import Eto.Drawing as drawing
import Eto.Forms as forms

import sys as _sys, os as _os
_rg_root = _os.path.normpath(_os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))
if _rg_root not in _sys.path:
    _sys.path.insert(0, _rg_root)
from ui import theme as _t
import importlib as _importlib; _importlib.reload(_t)


# Matches a {Key} placeholder. Key may contain spaces, but not braces.
_PLACEHOLDER_RE = re.compile(r"\{([^{}]+)\}")

# "No arrowhead" enum value. `None` is a reserved word in Python, so the enum
# member must be read with getattr.
try:
    _ARROW_NONE = getattr(Rhino.DocObjects.DimensionStyle.ArrowType, "None")
except Exception:
    _ARROW_NONE = None


# ══════════════════════════════════════════════════════════════════════════════
# Model helpers
# ══════════════════════════════════════════════════════════════════════════════

def get_object_user_text(obj_id):
    """Return {key: value} of one object's attribute user text (insertion order)."""
    obj = sc.doc.Objects.FindId(obj_id)
    pairs = {}
    if obj and obj.Attributes:
        us = obj.Attributes.GetUserStrings()
        if us:
            for k in us.AllKeys:
                pairs[k] = obj.Attributes.GetUserString(k)
    return pairs


def parse_template_keys(template):
    """Ordered, de-duplicated list of {Key} placeholders in the template."""
    seen = []
    for m in _PLACEHOLDER_RE.finditer(template or ""):
        key = m.group(1).strip()
        if key and key not in seen:
            seen.append(key)
    return seen


def ensure_keys(obj_id, keys):
    """Create any key the object is missing, with value 'TBD'.

    A live UserText field shows '####' if the referenced key does not exist,
    so every templated key must be present before the leader is created.
    Returns the list of keys that were created.
    """
    created = []
    present = set(rs.GetUserText(obj_id) or [])
    for key in keys:
        if key not in present:
            rs.SetUserText(obj_id, key, "TBD")
            created.append(key)
    return created


def build_leader_text(template, obj_id):
    """Replace every {Key} with a live UserText field bound to this object."""
    guid = str(obj_id)

    def repl(m):
        key = m.group(1).strip()
        return '%<UserText("{0}","{1}")>%'.format(guid, key)

    return _PLACEHOLDER_RE.sub(repl, template or "")


def _bbox_center(obj_id):
    """Fallback arrow point: bounding-box centre of the object."""
    obj = sc.doc.Objects.FindId(obj_id)
    if not obj:
        return None
    bb = obj.Geometry.GetBoundingBox(True)
    return bb.Center if bb.IsValid else None


# ══════════════════════════════════════════════════════════════════════════════
# Leader creation
# ══════════════════════════════════════════════════════════════════════════════

def _plane_through(point):
    """Copy of the active viewport CPlane, moved to pass through *point*."""
    vp = sc.doc.Views.ActiveView.ActiveViewport
    plane = rg.Plane(vp.ConstructionPlane())
    if point and point.IsValid:
        plane.Origin = point
    return plane


def _to_2d(plane, pt3d):
    rc, u, v = plane.ClosestParameter(pt3d)
    return rg.Point2d(u, v) if rc else None


def place_leader(template, obj_id, tip3d, text3d, style_id, text_height, show_arrow):
    """Create one leader. Returns its GUID, or None on failure."""
    plane = _plane_through(text3d)
    p_tip = _to_2d(plane, tip3d)
    p_txt = _to_2d(plane, text3d)
    if p_tip is None or p_txt is None:
        return None

    # Guarantee two distinct points so the leader is valid.
    if p_tip.DistanceTo(p_txt) < 1e-6:
        sep = text_height if text_height else 1.0
        p_txt = rg.Point2d(p_txt.X + sep, p_txt.Y + sep)

    text = build_leader_text(template, obj_id)
    pts = List[rg.Point2d]()
    pts.Add(p_tip)
    pts.Add(p_txt)
    gid = sc.doc.Objects.AddLeader(text, plane, pts)
    if gid == System.Guid.Empty:
        return None

    # Apply the chosen dimension style as the parent, then per-object overrides.
    obj = sc.doc.Objects.FindId(gid)
    if obj:
        leader = obj.Geometry
        try:
            if style_id is not None:
                leader.DimensionStyleId = style_id
        except Exception:
            pass
        # Per-object overrides: set directly on the geometry, then commit.
        # (Rhino 7/8 supports setting these on the Leader instance; the older
        # SetOverrideDimStyle route does not register an arrow change to None.)
        try:
            if text_height:
                leader.TextHeight = text_height
        except Exception:
            pass
        try:
            if (not show_arrow) and _ARROW_NONE is not None:
                leader.LeaderArrowType = _ARROW_NONE
        except Exception:
            pass
        try:
            obj.CommitChanges()
        except Exception:
            pass
    return gid


def place_text(template, obj_id, text3d, style_id, text_height):
    """Create a floating text object (no leader line). Returns GUID, or None."""
    plane = _plane_through(text3d)
    content = build_leader_text(template, obj_id)

    ds = None
    try:
        if style_id is not None:
            ds = sc.doc.DimStyles.FindId(style_id)
    except Exception:
        ds = None
    if ds is None:
        try:
            ds = sc.doc.DimStyles.Current
        except Exception:
            ds = None

    gid = System.Guid.Empty
    te = None
    try:
        if ds is not None:
            te = rg.TextEntity.Create(content, plane, ds, False, 0.0, 0.0)
    except Exception:
        te = None

    if te is not None:
        try:
            if text_height:
                te.TextHeight = text_height
        except Exception:
            pass
        try:
            gid = sc.doc.Objects.AddText(te)
        except Exception:
            gid = System.Guid.Empty

    if gid == System.Guid.Empty:
        # Fallback: simple string-based text object.
        try:
            h = text_height if text_height else (ds.TextHeight if ds else 1.0)
        except Exception:
            h = 1.0
        try:
            gid = sc.doc.Objects.AddText(content, plane, h, "Arial", False, False)
        except Exception:
            gid = System.Guid.Empty

    return gid if gid != System.Guid.Empty else None


def run_tagging(template, style_id, text_height, show_arrow, text_only, status_cb=None):
    """Two-click tagging loop. Click an object, then the text point. Enter ends.

    Returns the number of objects tagged.
    """
    keys = parse_template_keys(template)
    count = 0
    undo = sc.doc.BeginUndoRecord("Pregonero tagging")
    try:
        while True:
            go = Rhino.Input.Custom.GetObject()
            go.SetCommandPrompt("Select object to tag  (press Enter to finish)")
            go.AcceptNothing(True)
            go.EnablePreSelect(False, True)
            go.SubObjectSelect = False
            res = go.Get()
            if res != Rhino.Input.GetResult.Object:
                break  # Enter (Nothing) or Esc (Cancel) → finish

            oref = go.Object(0)
            obj_id = oref.ObjectId
            tip = oref.SelectionPoint()
            if tip is None or not tip.IsValid:
                tip = _bbox_center(obj_id)

            gp = Rhino.Input.Custom.GetPoint()
            gp.SetCommandPrompt("Pick text location  (Esc to skip this object)")
            if (not text_only) and tip and tip.IsValid:
                gp.DrawLineFromPoint(tip, True)
            if gp.Get() != Rhino.Input.GetResult.Point:
                continue  # skip this object but keep tagging others
            text_pt = gp.Point()

            ensure_keys(obj_id, keys)
            if text_only:
                gid = place_text(template, obj_id, text_pt, style_id, text_height)
            else:
                gid = place_leader(template, obj_id, tip, text_pt, style_id, text_height, show_arrow)
            if gid:
                count += 1
            sc.doc.Views.Redraw()
            if status_cb:
                status_cb(count)
    finally:
        sc.doc.EndUndoRecord(undo)
        sc.doc.Objects.UnselectAll()
        sc.doc.Views.Redraw()
    return count


def pick_single_object(prompt):
    """Pick one object and return its GUID, or None if cancelled."""
    go = Rhino.Input.Custom.GetObject()
    go.SetCommandPrompt(prompt)
    go.EnablePreSelect(True, True)
    go.SubObjectSelect = False
    if go.Get() != Rhino.Input.GetResult.Object:
        return None
    obj_id = go.Object(0).ObjectId
    sc.doc.Objects.UnselectAll()
    sc.doc.Views.Redraw()
    return obj_id


# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════

class PregoneroForm(forms.Form):
    """Modeless tagging tool — stays open between tagging sessions."""

    def __init__(self):
        super().__init__()
        self.Title = "Pregonero — Object Tagger"
        self.Padding = drawing.Padding(10)
        self.Resizable = True
        self.MinimumSize = drawing.Size(440, 540)
        self.ClientSize = drawing.Size(520, 660)
        self.BackgroundColor = _t.BG
        self.Owner = Rhino.UI.RhinoEtoApp.MainWindow

        self.template_obj_id = None
        self.available_keys = []      # key names from the template object
        self.key_values = {}          # {key: value}
        self.dim_styles = []          # [(name, id), ...]
        self._tagging = False

        self._build_ui()
        self._load_dim_styles()

    # ------------------------------------------------------------------ build
    def _build_ui(self):
        layout = forms.DynamicLayout()
        layout.DefaultSpacing = drawing.Size(6, 6)

        header = forms.Label()
        header.Text = "Tag objects with live, template-based leaders"
        header.Font = _t.F_HEAD
        layout.AddRow(header)

        desc = _t.hint("Leaders use live UserText fields; missing keys are created as \"TBD\".")
        layout.AddRow(desc)
        layout.AddRow(None)

        # 1 · Template object ------------------------------------------------
        layout.AddRow(_t.section_header("1 · Template object"))

        pick_btn = _t.btn("Pick template object", _t.BTN_DEFAULT)
        pick_btn.Click += self.on_pick_template
        self.obj_label = _t.hint("No template object picked yet.")
        pick_row = forms.StackLayout()
        pick_row.Orientation = forms.Orientation.Horizontal
        pick_row.Spacing = 8
        pick_row.VerticalContentAlignment = forms.VerticalAlignment.Center
        pick_row.Items.Add(forms.StackLayoutItem(pick_btn))
        pick_row.Items.Add(forms.StackLayoutItem(self.obj_label, True))
        layout.AddRow(pick_row)

        self.keys_list = forms.ListBox()
        self.keys_list.Height = 110
        layout.AddRow(self.keys_list)

        ins_btn = _t.btn("Insert {key}", _t.BTN_DEFAULT)
        ins_btn.Click += self.on_insert_key
        ins_all_btn = _t.btn("Insert all keys", _t.BTN_DEFAULT)
        ins_all_btn.Click += self.on_insert_all
        ins_row = forms.StackLayout()
        ins_row.Orientation = forms.Orientation.Horizontal
        ins_row.Spacing = 8
        ins_row.Items.Add(forms.StackLayoutItem(ins_btn))
        ins_row.Items.Add(forms.StackLayoutItem(ins_all_btn))
        layout.AddRow(ins_row)
        layout.AddRow(None)

        # 2 · Leader template ------------------------------------------------
        layout.AddRow(_t.section_header("2 · Leader template"))
        layout.AddRow(_t.hint("Type text and insert keys as {KeyName}. Line breaks are kept."))
        self.template_box = forms.TextArea()
        self.template_box.Height = 84
        self.template_box.Font = _t.F_MONO or _t.F_SANS
        layout.AddRow(self.template_box)
        layout.AddRow(None)

        # 3 · Leader style ---------------------------------------------------
        layout.AddRow(_t.section_header("3 · Leader style"))

        self.style_combo = forms.DropDown()
        layout.AddRow(self._labeled_row("Dimension style:", self.style_combo, expand=True))

        self.height_box = forms.TextBox()
        self.height_box.PlaceholderText = "(use style height)"
        self.height_box.Width = 120
        layout.AddRow(self._labeled_row("Text height:", self.height_box, expand=False))

        self.arrow_check = forms.CheckBox()
        self.arrow_check.Text = "Show leader arrowhead"
        self.arrow_check.Checked = True
        layout.AddRow(self.arrow_check)

        self.textonly_check = forms.CheckBox()
        self.textonly_check.Text = "No leader line (text only)"
        self.textonly_check.Checked = False
        self.textonly_check.CheckedChanged += self._on_textonly_changed
        layout.AddRow(self.textonly_check)
        layout.AddRow(None)

        # Buttons ------------------------------------------------------------
        start_btn = _t.btn("Start Tagging", _t.BTN_CALC)
        start_btn.Click += self.on_start
        refresh_btn = _t.btn("Refresh", _t.BTN_DEFAULT)
        refresh_btn.Click += self.on_refresh
        close_btn = _t.btn("Close", _t.BTN_CLEAR)
        close_btn.Click += self.on_close

        btn_row = forms.StackLayout()
        btn_row.Orientation = forms.Orientation.Horizontal
        btn_row.Spacing = 8
        btn_row.Items.Add(forms.StackLayoutItem(start_btn))
        btn_row.Items.Add(forms.StackLayoutItem(refresh_btn))
        btn_row.Items.Add(forms.StackLayoutItem(close_btn))
        layout.AddRow(btn_row)

        self.status_label = forms.Label()
        self.status_label.Text = "Ready."
        self.status_label.TextColor = _t.TEXT_MUTED
        layout.AddRow(self.status_label)
        layout.AddRow(None)

        scroll = forms.Scrollable()
        scroll.ExpandContentWidth = True
        scroll.ExpandContentHeight = False
        scroll.Content = layout
        self.Content = scroll

    # ------------------------------------------------------------- dim styles
    def _load_dim_styles(self):
        self.dim_styles = []
        names = []
        cur_id = None
        try:
            cur = sc.doc.DimStyles.Current
            cur_id = cur.Id if cur else None
        except Exception:
            cur_id = None
        for ds in sc.doc.DimStyles:
            try:
                self.dim_styles.append((ds.Name, ds.Id))
                names.append(ds.Name)
            except Exception:
                pass
        self.style_combo.DataStore = names
        if names:
            sel = 0
            for i, (_n, sid) in enumerate(self.dim_styles):
                if cur_id is not None and sid == cur_id:
                    sel = i
                    break
            self.style_combo.SelectedIndex = sel

    def _selected_style_id(self):
        i = self.style_combo.SelectedIndex
        if i is not None and 0 <= i < len(self.dim_styles):
            return self.dim_styles[i][1]
        return None

    # ---------------------------------------------------------------- helpers
    def _labeled_row(self, label_text, control, expand=True, label_w=108):
        """A horizontal label + control row that spans the form width."""
        row = forms.StackLayout()
        row.Orientation = forms.Orientation.Horizontal
        row.Spacing = 8
        row.VerticalContentAlignment = forms.VerticalAlignment.Center
        lab = forms.Label()
        lab.Text = label_text
        lab.Width = label_w
        row.Items.Add(forms.StackLayoutItem(lab))
        row.Items.Add(forms.StackLayoutItem(control, expand))
        return row

    def _set_status(self, text, state="info"):
        self.status_label.Text = text
        self.status_label.TextColor = _t.status_color(state)

    def _refresh_keys_list(self):
        self.keys_list.DataStore = [
            "{0}  =  {1}".format(k, self.key_values.get(k, "")) for k in self.available_keys
        ]

    def _insert_text(self, ins):
        try:
            ci = self.template_box.CaretIndex
            t = self.template_box.Text or ""
            self.template_box.Text = t[:ci] + ins + t[ci:]
            self.template_box.CaretIndex = ci + len(ins)
        except Exception:
            self.template_box.Text = (self.template_box.Text or "") + ins

    # --------------------------------------------------------------- handlers
    def on_pick_template(self, _s, _e):
        if self._tagging:
            return
        obj_id = pick_single_object("Select template object (carries the keys to tag)")
        if not obj_id:
            self._set_status("No object picked.", "info")
            return
        self.template_obj_id = obj_id
        self.key_values = get_object_user_text(obj_id)
        self.available_keys = list(self.key_values.keys())
        self._refresh_keys_list()
        n = len(self.available_keys)
        self.obj_label.Text = "Picked {0}…  —  {1} key(s)".format(str(obj_id)[:8], n)
        if n == 0:
            self._set_status(
                "That object has no user-text keys. You can still type {Key} names manually.",
                "warn",
            )
        else:
            self._set_status("Loaded {0} key(s). Insert them into the template.".format(n), "ok")

    def on_insert_key(self, _s, _e):
        i = self.keys_list.SelectedIndex
        if i is None or i < 0 or i >= len(self.available_keys):
            self._set_status("Select a key in the list first.", "info")
            return
        self._insert_text("{" + self.available_keys[i] + "}")

    def on_insert_all(self, _s, _e):
        if not self.available_keys:
            self._set_status("No keys to insert — pick a template object first.", "info")
            return
        existing = self.template_box.Text or ""
        if existing and not existing.endswith("\n"):
            existing += "\n"
        self.template_box.Text = existing + "\n".join(
            "{" + k + "}" for k in self.available_keys
        )

    def on_refresh(self, _s, _e):
        self._load_dim_styles()
        if self.template_obj_id:
            self.key_values = get_object_user_text(self.template_obj_id)
            self.available_keys = list(self.key_values.keys())
            self._refresh_keys_list()
        self._set_status("Refreshed styles and keys.", "info")

    def on_start(self, _s, _e):
        if self._tagging:
            return
        template = (self.template_box.Text or "").strip()
        if not template:
            self._set_status("Write a leader template first.", "warn")
            return

        height = None
        ht = (self.height_box.Text or "").strip()
        if ht:
            try:
                height = float(ht)
                if height <= 0:
                    height = None
            except ValueError:
                self._set_status("Text height must be a number.", "error")
                return

        style_id = self._selected_style_id()
        show_arrow = bool(self.arrow_check.Checked)
        text_only = bool(self.textonly_check.Checked)
        self._tagging = True
        self._set_status("Tagging… click objects in the viewport; Enter to finish.", "info")

        def cb(c):
            self._set_status("Tagged {0} object(s)…".format(c), "info")

        try:
            count = run_tagging(template, style_id, height, show_arrow, text_only, cb)
            state = "ok" if count else "warn"
            self._set_status("Done — {0} object(s) tagged.".format(count), state)
        except Exception as ex:
            self._set_status("Error: {0}".format(ex), "error")
        finally:
            self._tagging = False

    def _on_textonly_changed(self, _s, _e):
        # The arrowhead option is irrelevant when there is no leader line.
        self.arrow_check.Enabled = not bool(self.textonly_check.Checked)

    def on_close(self, _s, _e):
        self.Close()


def main():
    form = PregoneroForm()
    form.Show()


if __name__ == "__main__":
    main()
