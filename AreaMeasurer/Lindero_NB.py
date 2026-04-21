#! python3
# -*- coding: utf-8 -*-
# __title__ = "Lindero NB"
# __doc__ = """
# Neo-brutalist UI example for Rhino 8 / CPython 3
# Mirrors the Lindero layout with a soft neo-brutalist design system:
#   - Warm-cream background, hard-bordered panels
#   - Soft-pastel accent colors (yellow header, green total, blue/red buttons)
#   - Monospace font for all data, tabular alignment
#   - High data legibility, no gradients or blur
# """

import Rhino
import Eto.Drawing as drawing
import Eto.Forms as forms

# ── Palette ───────────────────────────────────────────────────────────────────

BG          = drawing.Color.FromArgb(248, 244, 235)   # warm cream
HEADER      = drawing.Color.FromArgb(255, 240, 120)   # soft yellow
ROW_ALT     = drawing.Color.FromArgb(241, 237, 227)   # subtle stripe
TOTAL_BG    = drawing.Color.FromArgb(195, 231, 192)   # soft green
BTN_CALC    = drawing.Color.FromArgb(168, 210, 255)   # soft blue
BTN_CLEAR   = drawing.Color.FromArgb(255, 188, 188)   # soft red/pink
BTN_DEFAULT = drawing.Color.FromArgb(218, 213, 203)   # warm gray
TAB_ACTIVE  = drawing.Color.FromArgb(255, 240, 120)   # same as header
TAB_INACTIVE= drawing.Color.FromArgb(232, 228, 219)   # muted warm
BORDER      = drawing.Color.FromArgb(18,  18,  18)    # near-black
TEXT        = drawing.Color.FromArgb(18,  18,  18)
TEXT_MUTED  = drawing.Color.FromArgb(98,  93,  85)
WHITE       = drawing.Color.FromArgb(255, 255, 255)

# ── Typography ────────────────────────────────────────────────────────────────

F_MONO   = drawing.Font("Courier New", 9.0)
F_MONO_B = drawing.Font("Courier New", 9.5, drawing.FontStyle.Bold)
F_SANS   = drawing.Font("Segoe UI",    9.0)
F_SANS_B = drawing.Font("Segoe UI",    9.0, drawing.FontStyle.Bold)
F_SANS_S = drawing.Font("Segoe UI",    8.0)

# ── Sample data ───────────────────────────────────────────────────────────────

SAMPLE = [
    ("c1150600…",  0.0000),
    ("3f29f874…",  2700.0000),
    ("8b5f3ea3…",  2700.0000),
    ("9a42088d…",  2700.0000),
    ("5aeffe7b…",  2700.0000),
    ("7efa09f6…",  2700.0000),
    ("4ba59e3a…",  2700.0000),
    ("325b0e21…",  2700.0000),
    ("210bcba7…",  2700.0000),
    ("2e1045b8…",  2700.0000),
    ("9a497ef8…",  2700.0000),
    ("4dc2715d…",  2700.0000),
    ("5cb24c5b…",  2700.0000),
    ("65ca72d2…",  2700.0000),
    ("f5c646e5…",  2700.0000),
    ("0a046cbe…",  2700.0000),
    ("6fe2a6d9…",  2700.0000),
    ("2c56ad7f…",  2700.0000),
    ("112d1a21…",  2700.0000),
    ("a2a74f4c…",  2700.0000),
    ("7994c42d…",  2700.0000),
    ("c6dcdc81…",  2700.0000),
    ("7864509f…",  2700.0025),
    ("d4821218…", 13159.2295),
    ("9e9b7409…",  2700.0025),
    ("663a789d…",  2700.0025),
    ("f739e94c…",  2700.0025),
]


# ── Small builder helpers ─────────────────────────────────────────────────────

def _lbl(text, font=None, color=None, align=None):
    w = forms.Label()
    w.Text = text
    if font:  w.Font          = font
    if color: w.TextColor     = color
    if align: w.TextAlignment = align
    return w


def _btn(text, bg):
    w = forms.Button()
    w.Text            = text
    w.Font            = F_SANS_B
    w.BackgroundColor = bg
    return w


def _pad(content, bg, h=8, v=6):
    p = forms.Panel()
    p.BackgroundColor = bg
    p.Padding         = drawing.Padding(h, v)
    p.Content         = content
    return p


# ── Form ──────────────────────────────────────────────────────────────────────

class LinderoNBForm(forms.Form):

    def __init__(self):
        super().__init__()
        self.Title           = "Lindero — Footprint Area Calculator"
        self.MinimumSize     = drawing.Size(700, 680)
        self.Resizable       = True
        self.BackgroundColor = BG
        self.Owner           = Rhino.UI.RhinoEtoApp.MainWindow

        self._build()
        self._populate(SAMPLE)

    # ── Top navigation ────────────────────────────────────────────────────────

    def _top_tabs(self):
        row = forms.StackLayout()
        row.Orientation = forms.Orientation.Horizontal
        row.Spacing     = 1

        for text in ["S4 — Custom Aggregation", "R1 — Room Analysis",
                     "R2 — Group Analysis", "Settings"]:
            b = _btn(text, TAB_INACTIVE)
            b.Font = F_SANS
            row.Items.Add(forms.StackLayoutItem(b, True))

        return _pad(row, BORDER, 1, 1)

    def _sub_tabs(self):
        row = forms.StackLayout()
        row.Orientation = forms.Orientation.Horizontal
        row.Spacing     = 1

        for text, active in [("S1 — Selected Objects", True),
                               ("S2 — By Layer",         False),
                               ("S3 — Layer Hierarchy",  False)]:
            b = _btn(text, TAB_ACTIVE if active else TAB_INACTIVE)
            b.Font = F_SANS_B if active else F_SANS
            row.Items.Add(forms.StackLayoutItem(b, True))

        return _pad(row, BORDER, 1, 1)

    # ── Info rows ─────────────────────────────────────────────────────────────

    def _desc_row(self):
        return _pad(
            _lbl("Individual footprint per selected object.  No overlap handling.",
                 F_SANS, TEXT_MUTED),
            BG, 10, 6
        )

    def _name_key_row(self):
        combo = forms.DropDown()
        combo.Font = F_MONO
        combo.Items.Add("")

        tbl = forms.TableLayout()
        tbl.Spacing = drawing.Size(8, 0)
        tbl.Rows.Add(forms.TableRow(
            forms.TableCell(_lbl("Name Key (optional):", F_SANS_B, TEXT)),
            forms.TableCell(combo, True),
        ))
        return _pad(tbl, BG, 10, 5)

    def _instr_row(self):
        return _pad(
            _lbl("Select objects in Rhino (form stays open), then click Calculate.",
                 F_SANS_S, TEXT_MUTED),
            BG, 10, 3
        )

    # ── Data table ────────────────────────────────────────────────────────────

    @staticmethod
    def _data_row(name, area, alt=False):
        """One Object/Area row as a TableLayout with two cells."""
        bg = ROW_ALT if alt else WHITE
        tbl = forms.TableLayout()
        tbl.Spacing = drawing.Size(0, 0)
        name_lbl = _pad(_lbl(name, F_MONO, TEXT), bg, 10, 4)
        area_lbl = _pad(
            _lbl(f"{area:,.4f}", F_MONO, TEXT, forms.TextAlignment.Right),
            bg, 10, 4
        )
        tbl.Rows.Add(forms.TableRow(
            forms.TableCell(name_lbl),
            forms.TableCell(area_lbl),
        ))
        return tbl

    def _grid_section(self):
        # Scenario header bar
        hdr_panel = _pad(
            _lbl("SCENARIO 1 — SELECTED OBJECTS    [m²]", F_MONO_B, TEXT),
            HEADER, 10, 7
        )

        # Column-header row (fixed, above the scrollable)
        col_hdr = forms.TableLayout()
        col_hdr.Spacing = drawing.Size(0, 0)
        col_hdr.BackgroundColor = BTN_DEFAULT
        col_hdr.Rows.Add(forms.TableRow(
            forms.TableCell(_pad(_lbl("Object", F_SANS_B, TEXT),                BTN_DEFAULT, 10, 5)),
            forms.TableCell(_pad(_lbl("Area",   F_SANS_B, TEXT,
                                      forms.TextAlignment.Right), BTN_DEFAULT, 10, 5)),
        ))

        # Scrollable area — rows added by _populate()
        self._rows_stack = forms.StackLayout()
        self._rows_stack.Orientation = forms.Orientation.Vertical
        self._rows_stack.Spacing     = 0
        self._rows_stack.BackgroundColor = WHITE

        scroll = forms.Scrollable()
        scroll.Content         = self._rows_stack
        scroll.BackgroundColor = WHITE
        scroll.Border          = forms.BorderType.Line

        # TOTAL row (fixed, below the scrollable)
        self._lbl_total_val = forms.Label()
        self._lbl_total_val.Text          = "0.0000"
        self._lbl_total_val.Font          = F_MONO_B
        self._lbl_total_val.TextColor     = TEXT
        self._lbl_total_val.TextAlignment = forms.TextAlignment.Right

        total_tbl = forms.TableLayout()
        total_tbl.Spacing = drawing.Size(0, 0)
        total_tbl.Rows.Add(forms.TableRow(
            forms.TableCell(_pad(_lbl("TOTAL", F_MONO_B, TEXT), TOTAL_BG, 10, 6)),
            forms.TableCell(_pad(self._lbl_total_val,           TOTAL_BG, 10, 6)),
        ))

        # Assemble
        wrapper = forms.TableLayout()
        wrapper.Spacing = drawing.Size(0, 0)
        wrapper.Rows.Add(self._trow(hdr_panel))
        wrapper.Rows.Add(self._trow(col_hdr))
        wrapper.Rows.Add(self._trow(scroll, scale=True))
        wrapper.Rows.Add(self._trow(total_tbl))
        return wrapper

    # ── Action buttons ────────────────────────────────────────────────────────

    def _buttons_row(self):
        row = forms.StackLayout()
        row.Orientation = forms.Orientation.Horizontal
        row.Spacing     = 5

        for text, bg in [
            ("Calculate",             BTN_CALC),
            ("Clear",                 BTN_CLEAR),
            ("Refresh Model",         BTN_DEFAULT),
            ("Export to Excel",       BTN_DEFAULT),
            ("Write Area to Objects", BTN_DEFAULT),
            ("Export Chart as PNG",   BTN_DEFAULT),
        ]:
            row.Items.Add(forms.StackLayoutItem(_btn(text, bg)))

        return _pad(row, BG, 8, 7)

    # ── Status bar ────────────────────────────────────────────────────────────

    def _status_bar(self):
        self._status = _lbl("S1 — 0 object(s) | Total: 0.0000 m²", F_SANS, TEXT)
        return _pad(self._status, BTN_DEFAULT, 10, 5)

    # ── Root layout ───────────────────────────────────────────────────────────

    @staticmethod
    def _trow(ctrl, scale=False):
        row = forms.TableRow(forms.TableCell(ctrl))
        row.ScaleHeight = scale
        return row

    def _build(self):
        root = forms.TableLayout()
        root.Padding = drawing.Padding(0)
        root.Spacing = drawing.Size(0, 0)

        root.Rows.Add(self._trow(self._top_tabs()))
        root.Rows.Add(self._trow(self._sub_tabs()))
        root.Rows.Add(self._trow(self._desc_row()))
        root.Rows.Add(self._trow(self._name_key_row()))
        root.Rows.Add(self._trow(self._instr_row()))
        root.Rows.Add(self._trow(self._grid_section(), scale=True))
        root.Rows.Add(self._trow(self._buttons_row()))
        root.Rows.Add(self._trow(self._status_bar()))

        self.Content = root

    # ── Data ──────────────────────────────────────────────────────────────────

    def _populate(self, data):
        for i, (name, area) in enumerate(data):
            row = self._data_row(name, area, alt=i % 2 == 0)
            self._rows_stack.Items.Add(forms.StackLayoutItem(row))
        total = sum(a for _, a in data)
        self._lbl_total_val.Text = f"{total:,.4f}"
        self._status.Text = f"S1 — {len(data)} object(s) | Total: {total:,.4f} m²"

    def OnClosed(self, e):
        self.Dispose()
        super().OnClosed(e)


# ── Entry point ───────────────────────────────────────────────────────────────

form = LinderoNBForm()
form.Show()
