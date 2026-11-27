# -*- coding: utf-8 -*-
"""Golden-value generator for RhinoGuire.Core.Tests (WP-1B C# port).

Runs the real Python functions of the RhinoGuire tools and writes Golden.g.cs
next to this file. Rerun after changing any ported Python behaviour:

    python src/RhinoGuire.Core.Tests/Golden/generate_golden.py

Lindero's _gather_config/_apply_config are extracted from Lindero.py with ast
and run against minimal fake Eto controls.
"""
import ast
import json
import math
import os
import re
import sys
import types
import unicodedata

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = sys.argv[1] if len(sys.argv) > 1 else os.path.normpath(os.path.join(HERE, "..", "..", ".."))
SNIPPETS = {}
sys.path.insert(0, REPO)

from TerrainTools._core import slope, volumes          # noqa: E402
from AreaMeasurer import _paths as _lp                 # noqa: E402


def cs(s):
    """C# regular string literal."""
    if s is None:
        return "null"
    out = ['"']
    for ch in s:
        o = ord(ch)
        if ch == "\\":
            out.append("\\\\")
        elif ch == '"':
            out.append('\\"')
        elif ch == "\n":
            out.append("\\n")
        elif ch == "\r":
            out.append("\\r")
        elif ch == "\t":
            out.append("\\t")
        elif 0x20 <= o < 0x7F:
            out.append(ch)
        elif o <= 0xFFFF:
            out.append("\\u%04x" % o)
        else:
            out.append("\\U%08x" % o)
    out.append('"')
    return "".join(out)


def csd(x):
    if isinstance(x, float):
        if math.isnan(x):
            return "double.NaN"
        if math.isinf(x):
            return "double.PositiveInfinity" if x > 0 else "double.NegativeInfinity"
        r = repr(x)
        return r if ("." in r or "e" in r) else r + ".0"
    return repr(float(x))


def write(name, lines):
    SNIPPETS[name] = list(lines)


# ── formatting ─────────────────────────────────────────────────────────────

values = [0.0, -0.0, 0.125, 0.375, 2.675, 1.005, 2.5, 0.5, 1.5, -0.001, -2.5,
          1234567.891, -1234567.891, 999.995, 1e20, 1e-7, 123456789012.345,
          5e-324, 1.7976931348623157e308, 0.1, 100.0, 12.345, 1000.0, 999999.995,
          float("nan"), float("inf"), float("-inf"), 2640.0, 1200.0 / 2.2, 1e16]
lines = []
for v in values:
    for d in (0, 1, 2, 3, 4):
        lines.append("{ %s, %d, false, %s }," % (csd(v), d, cs(("{:.%df}" % d).format(v))))
        lines.append("{ %s, %d, true, %s }," % (csd(v), d, cs(("{:,.%df}" % d).format(v))))
write("fixed.txt", lines)

gvals = [0.25, 1234.5, 0.0001234, 0.00001234, 1e16, 123.0, 999.5, 9.995, 0.999,
         1.0, 100.0, 2.5e-5, 0.0, -0.0, -3.14159, 12345678.0, 0.5, 1.25, 99950.0,
         float("nan"), float("inf"), float("-inf"), 5e-324, 1e100, 0.000099999]
lines = []
for v in gvals:
    for p in (0, 1, 3, 6):
        lines.append("{ %s, %d, %s }," % (csd(v), p, cs(("%%.%dg" % p) % v)))
write("general.txt", lines)

rvals = [0.1, 0.1 + 0.2, 1e16, 1e15, 123456789012345680.0, 1e-4, 1e-5, 2.2, 10.0,
         5e-324, 1.7976931348623157e308, -0.0, 0.0, 0.5, 100.0, 1.5e300, 12345.678,
         1e22, 1.0 / 3.0, 2640.0, 1200.0 / 2.2, 0.0001, 0.00012345, 9999999999999998.0,
         float("nan"), float("inf"), float("-inf"), -1.5e-7, 4.0, 0.25, 2.0]
lines = ["{ %s, %s }," % (csd(v), cs(repr(v))) for v in rvals]
write("repr.txt", lines)

pstrs = ["1", " 2.5 ", "1_000", "1__000", "_1", "1_", "1_000.000_1", "1e5", "1E-3",
         ".5", "5.", ".", "", "  ", "+1.5", "-0", "inf", "-Infinity", "INF", "nan",
         "+nan", "infinit", "1e", "e5", "0x10", "1,5", "1.200,5", "١٢",
         " 1.5 ", "\x1c3\x1f", "1e400", "-1e400", "1 2", "١.٥", "12abc",
         "1.5e+2", "1_0e1_0", "1._5", "1_.5", "abc"]
lines = []
for s in pstrs:
    try:
        v = float(s)
        lines.append("{ %s, true, %s }," % (cs(s), csd(v)))
    except ValueError:
        lines.append("{ %s, false, 0.0 }," % cs(s))
write("pyfloat.txt", lines)

# ── slope ──────────────────────────────────────────────────────────────────

lines = []
for m in (0.5, 1.0, 0.1, 0.25, 2.0, 1.0 / 3.0, 0.0333, 1e-6, 1e6, 0.0, -1.0,
          float("nan"), float("inf")):
    for u in ("ratio_hv", "percent", "degrees", "bogus"):
        lines.append("{ %s, %s, %s }," % (csd(m), cs(u), cs(slope.format_slope(m, u))))
write("slope_format.txt", lines)
lines = []
for m in (0.5, 1.0, 0.1, 1.0 / 3.0, 1e-6, 0.0, -2.0, float("nan"), float("inf")):
    lines.append("{ %s, %s }," % (csd(m), cs(slope.describe(m))))
write("slope_describe.txt", lines)

lines = []
for text, unit in [("2:1", "ratio_hv"), (" 3 : 2 ", "ratio_hv"), ("2", "ratio_hv"),
                   ("1.5", "ratio_hv"), ("0:1", "ratio_hv"), ("2:0", "ratio_hv"),
                   ("-2:1", "ratio_hv"), ("2:1:3", "ratio_hv"), (":1", "ratio_hv"),
                   ("abc", "ratio_hv"), ("", "ratio_hv"), ("1_0:1", "ratio_hv"),
                   ("50", "percent"), ("12.5", "percent"), ("0", "percent"),
                   ("-5", "percent"), ("x", "percent"), ("2:1", "percent"),
                   ("45", "degrees"), ("30", "degrees"), ("90", "degrees"),
                   ("89.9", "degrees"), ("0", "degrees"), ("nan", "percent"),
                   ("inf", "percent"), ("inf", "ratio_hv"), ("1e-3", "degrees")]:
    try:
        m = slope.to_gradient(text, unit)
        lines.append("{ %s, %s, true, %s }," % (cs(text), cs(unit), csd(m)))
    except ValueError:
        lines.append("{ %s, %s, false, 0.0 }," % (cs(text), cs(unit)))
write("slope_text.txt", lines)

# ── volumes: histogram on a mixed grid ─────────────────────────────────────

class G(object):
    def __init__(self, nx, ny, cell):
        self.nx, self.ny, self.cell = nx, ny, cell
        self.z_design = [[None] * nx for _ in range(ny)]
        self.z_terrain = [[None] * nx for _ in range(ny)]
        self.region_mask = [[False] * nx for _ in range(ny)]
        self.stations = []

g = G(7, 5, 0.5)
for j in range(5):
    for i in range(7):
        g.z_terrain[j][i] = math.sin(i * 0.7) + 0.3 * j
        g.z_design[j][i] = 0.8 if (i + j) % 3 else None if i == 3 else 0.2
        g.region_mask[j][i] = not (i == 0 and j == 4)
k = volumes.cut_fill(g, tol=1e-4, n_bins=5)
lines = ["// cut_fill(grade, tol=1e-4, n_bins=5) on VolumesTests.MixedGrid()"]
for key, prop in (("cut_volume", "CutVolume"), ("fill_volume", "FillVolume"), ("net", "Net"),
                  ("cut_area", "CutArea"), ("fill_area", "FillArea"), ("region_area", "RegionArea"),
                  ("balance_ratio", "BalanceRatio"), ("cell_area", "CellArea"),
                  ("min_delta", "MinDelta"), ("max_delta", "MaxDelta")):
    lines.append("%s = %s," % (prop, csd(k[key])))
lines.append("NCells = %d," % k["n_cells"])
lines.append("Bins = new (double Lo, double Hi, int Count)[]")
lines.append("{")
for lo, hi, c in k["depth_histogram"]:
    lines.append("    (%s, %s, %d)," % (csd(lo), csd(hi), c))
lines.append("},")
write("volumes_mixed.txt", lines)

# ── paths extras ───────────────────────────────────────────────────────────

lines = []
for n in ["Ebene_١٢", "B", "b", "Ä", "a", "Ebene_99999999999999999999999",
          "x-5y", "--3", "E_10_-2"]:
    k = _lp.level_sort_key(n)
    lines.append("{ %s, %d, %s, %s }," % (cs(n), k[0], cs(str(k[1])), cs(k[2])))
write("levelkeys.txt", lines)

# ── Baquiano / Pregonero (extracted from the tool files) ───────────────────

def extract(path, names, extra_ns=None):
    src = open(path, encoding="utf-8").read()
    tree = ast.parse(src)
    ns = dict(extra_ns or {})
    keep = []
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.ClassDef)) and node.name in names:
            keep.append(node)
        elif isinstance(node, ast.Assign) and any(
                isinstance(t, ast.Name) and t.id in names for t in node.targets):
            keep.append(node)
    mod = ast.Module(body=keep, type_ignores=[])
    exec(compile(mod, path, "exec"), ns)
    return ns

bq = extract(os.path.join(REPO, "SearchData", "Baquiano.py"), {"SearchCondition"})
SC = bq["SearchCondition"]
mtypes = ["contains", "equals", "starts_with", "ends_with", "not_contains",
          "not_equals", "not_starts_with", "not_ends_with"]
pairs = [("Pflege", "pflege"), ("Pflege Station A", "STATION"), ("Pflege", "Pfl"),
         ("Pflege", "ege"), ("Pflege", "xyz"), ("", "a"), ("Straße", "STRASSE"),
         ("ÄRZTE", "ärz"), (None, "a"), ("abc", "")]
lines = []
for ov, sv in pairs:
    for i, mt in enumerate(mtypes):
        lines.append("{ %s, %s, %d, %s }," % (cs(ov), cs(sv), i,
                                              "true" if SC("k", sv, mt).matches_value(ov) else "false"))
write("baquiano.txt", lines)

pg = extract(os.path.join(REPO, "Tagging", "Pregonero.py"),
             {"_PLACEHOLDER_RE", "parse_template_keys", "build_leader_text"}, {"re": re})
templates = ["Room: {RoomName} / {Level}", "{A}{B}{A}", "{ Spaced Key }", "{}", "{ }",
             "no keys", "{{Nested}}", "{a{b}c}", "{Multi\nLine}", "", None,
             "{Ä}·{😀}", "x {Key} y {Key } z", "{\"q\"}"]
lines = []
for t in templates:
    keys = pg["parse_template_keys"](t)
    text = pg["build_leader_text"](t, "0b1c2d3e-0000-4000-8000-123456789abc")
    lines.append("{ %s, new string[] { %s }, %s }," % (cs(t), ", ".join(cs(k) for k in keys), cs(text)))
write("pregonero.txt", lines)

# ── Lindero settings (fake Eto) ────────────────────────────────────────────

class Event(object):
    def __init__(self):
        self.handlers = []

    def __iadd__(self, h):
        self.handlers.append(h)
        return self

    def fire(self):
        for h in list(self.handlers):
            h(None, None)


class Items(list):
    def Add(self, x):
        self.append(x)

    @property
    def Count(self):
        return len(self)

    def RemoveAt(self, i):
        del self[i]


class Ctrl(object):
    def __init__(self):
        self.Visible = True
        self.Width = 0


class TextBox(Ctrl):
    def __init__(self):
        super().__init__()
        self._text = ""
        self.TextChanged = Event()
        self.PlaceholderText = ""

    @property
    def Text(self):
        return self._text

    @Text.setter
    def Text(self, v):
        if v is not None and not isinstance(v, str):
            raise TypeError("Text: expected str, got %s" % type(v).__name__)
        self._text = "" if v is None else v
        self.TextChanged.fire()


ComboBox = TextBox


class DropDown(Ctrl):
    def __init__(self):
        super().__init__()
        self.Items = Items()
        self._idx = -1
        self.SelectedIndexChanged = Event()

    @property
    def SelectedIndex(self):
        return self._idx

    @SelectedIndex.setter
    def SelectedIndex(self, v):
        self._idx = v
        self.SelectedIndexChanged.fire()


class NumericStepper(Ctrl):
    def __init__(self):
        super().__init__()
        self.MinValue = -1e300
        self.MaxValue = 1e300
        self.DecimalPlaces = 0
        self.Increment = 1
        self._v = 0.0

    @property
    def Value(self):
        return self._v

    @Value.setter
    def Value(self, v):
        v = float(v)
        self._v = max(self.MinValue, min(self.MaxValue, v))


class Button(Ctrl):
    def __init__(self):
        super().__init__()
        self.Click = Event()
        self.Text = ""


class StackLayout(Ctrl):
    def __init__(self):
        super().__init__()
        self.Items = Items()
        self.Orientation = None
        self.Spacing = 0


class StackLayoutItem(object):
    def __init__(self, control, expand=False):
        self.Control = control


fake_forms = types.SimpleNamespace(
    TextBox=TextBox, ComboBox=ComboBox, DropDown=DropDown, NumericStepper=NumericStepper,
    Button=Button, StackLayout=StackLayout, StackLayoutItem=StackLayoutItem,
    Orientation=types.SimpleNamespace(Horizontal="h", Vertical="v"))
fake_t = types.SimpleNamespace(bind_key_search=lambda cb, keys: (lambda k: None))

lindero_path = os.path.join(REPO, "AreaMeasurer", "Lindero.py")
lsrc = open(lindero_path, encoding="utf-8").read()
ltree = ast.parse(lsrc)
module_ns = {"forms": fake_forms, "_t": fake_t, "_lp": _lp, "unicodedata": unicodedata,
             "json": json}
mod_keep = []
for node in ltree.body:
    if isinstance(node, ast.Assign) and any(
            isinstance(t, ast.Name) and t.id in ("DIM_LEVEL", "DIM_CATEGORY", "DIM_LAYER",
                                                 "DIM_USERTEXT", "_INVISIBLE_CHARS")
            for t in node.targets):
        mod_keep.append(node)
    if isinstance(node, ast.FunctionDef) and node.name in ("_clean_label", "_dim_label"):
        mod_keep.append(node)
exec(compile(ast.Module(body=mod_keep, type_ignores=[]), lindero_path, "exec"), module_ns)

method_names = {"_gather_config", "_apply_config", "_s4_add_dim_row", "_s4_clear_dim_rows",
                "_s4_remove_dim_row", "_s4_dims", "_apply_r_dim", "_targets_dict",
                "_add_target_row", "_clear_target_rows", "_remove_target_row",
                "_gross_net_factor", "_ignore_prefix", "_selected_layer",
                "_refresh_targets_layout"}
cls = [n for n in ltree.body if isinstance(n, ast.ClassDef) and n.name == "LinderoForm"][0]
meths = [n for n in cls.body if isinstance(n, ast.FunctionDef) and n.name in method_names]
assert len(meths) == len(method_names), [m.name for m in meths]
fake_cls = ast.ClassDef(name="FakeLindero", bases=[], keywords=[], body=meths,
                        decorator_list=[])
fake_cls = ast.fix_missing_locations(ast.copy_location(fake_cls, cls))
exec(compile(ast.Module(body=[fake_cls], type_ignores=[]), lindero_path, "exec"), module_ns)
FakeLindero = module_ns["FakeLindero"]


def stepper(lo, hi, v):
    s = NumericStepper()
    s.MinValue, s.MaxValue = lo, hi
    s.Value = v
    return s


def new_form(layers):
    f = FakeLindero()
    f.available_layers = list(layers)
    f.available_keys = []
    f._settings_scroll = None
    f.tolerance_stepper = stepper(0.0, 50.0, 10.0)
    f.decimal_stepper = stepper(0, 4, 2)
    f.z_height_tol_stepper = stepper(0.0, 100.0, 0.5)
    f.factor_stepper = stepper(0.01, 100.0, 2.20)
    f.ignore_prefix_tb = TextBox()
    f.ignore_prefix_tb.Text = _lp.DEFAULT_IGNORE_PREFIX
    f.parent_layer_dd = ComboBox()
    f.parent_layer_s4_dd = ComboBox()
    f._ks_s4 = {}
    f._s4_keys_layout = StackLayout()
    f._s4_dim_rows = []
    f._s4_add_dim_row(module_ns["DIM_LAYER"], 1)
    f._s4_add_dim_row(module_ns["DIM_LAYER"], 2)
    for n, idx in (("r1", 0), ("r2", 1)):
        dd = DropDown()
        for t in ("Category (layer)", "Level (layer)", "UserText key"):
            dd.Items.Add(t)
        dd.SelectedIndex = idx
        cb = ComboBox()
        cb.Visible = False
        md = DropDown()
        md.Items.Add("Brutto (BGF)")
        md.Items.Add("Netto (NF)")
        md.SelectedIndex = 0
        setattr(f, n + "_dim_dd", dd)
        setattr(f, n + "_key_cb", cb)
        setattr(f, n + "_metric_dd", md)
    f._target_rows = []
    f._targets_layout = StackLayout()
    return f


P = "07-PROPOSAL-Phase_03"
LAYERS = [P, P + "::Ebene_06", "99-DOCUMENTATION", "Grün::Ebene_01"]

README_V2 = """{
  "config_version": 2,
  "tolerance_percent": 10.0,
  "decimal_places": 2,
  "z_height_tol": 0.5,
  "gross_net_factor": 2.2,
  "ignore_prefix": "_",
  "s3_parent_layer": "07-PROPOSAL-Phase_03",
  "s4_parent_layer": "07-PROPOSAL-Phase_03",
  "s4_dimensions": [["layer", 1], ["layer", 2]],
  "r1_dimension": ["category", null],
  "r1_metric": "brutto",
  "r2_dimension": ["level", null],
  "r2_metric": "netto",
  "targets": {
    "Pflege": {"brutto": 2640.0, "netto": 1200.0},
    "Diagnostik und Therapie": {"brutto": 5280.0, "netto": 2400.0}
  }
}"""

V1 = json.dumps({
    "room_target_key": "Raum", "group_target_key": "", "tolerance_percent": 7.5,
    "s4_parent_layer": P, "s4_key_sequence": ["Abteilung", "", "Raumtyp", "Gebäude"],
    "r1r2_source": 0, "r1_level_index": 1, "r2_level_index": 2}, indent=2)

cases = [
    ("defaults", None, LAYERS),
    ("readme_v2", README_V2, LAYERS),
    ("readme_v2_layers_missing", README_V2, ["Other"]),
    ("v1_file", V1, LAYERS),
    ("v1_empty_keys", '{"s4_key_sequence": ["", ""], "tolerance_percent": "12.5"}', LAYERS),
    ("v1_no_version_minimal", "{}", LAYERS),
    ("clamping", '{"config_version": 2, "tolerance_percent": 75, "decimal_places": 7, '
                 '"z_height_tol": -1, "gross_net_factor": 0, "ignore_prefix": "  xx  "}', LAYERS),
    ("decimals_truncate", '{"config_version": 2, "decimal_places": 2.9, '
                          '"targets": {"A": {"brutto": 1000}}}', LAYERS),
    ("targets_variety", json.dumps({
        "config_version": 2, "gross_net_factor": 2.0, "decimal_places": 3,
        "targets": {
            "b": {"brutto": 100.5},
            "A": {"netto": 40},
            "a": 900,
            "Ärzte": {"brutto": "1500,5", "netto": None},
            "Nulls": {"brutto": None, "netto": None},
            "Bad": {"brutto": "1.200,5"},
            "Str": "abc",
            "​Tech Raum‍ ": {"brutto": 10.0, "netto": 4.0},
            "Tech Raum": {"brutto": 20.0},
            "ﬁle": {"netto": 1.5},
            "Bool": {"brutto": True},
            "List": {"brutto": [1, "x"]},
            "Neg": {"brutto": -0.0},
            "Big": {"brutto": 1e16},
            "Tiny": {"netto": 1e-7},
        }}, ensure_ascii=False), LAYERS),
    ("nan_target", '{"config_version": 2, "targets": {"N": {"brutto": NaN}, '
                   '"I": {"netto": Infinity}}}', LAYERS),
    ("s4_dims_variety", json.dumps({
        "config_version": 2,
        "s4_dimensions": [["layer", 3], ["usertext", "Abteilung"], ["level"], [],
                          ["layer", "4"], ["layer", 12], ["layer", 0], ["usertext", 5],
                          ["layer", 2.9], ["usertext", " Key "], ["usertext", ""],
                          ["category", "x"], "layer", ["layer", True]]}), LAYERS),
    ("s4_dims_empty_list_keeps_rows", '{"config_version": 2, "s4_dimensions": []}', LAYERS),
    ("r_dims", json.dumps({
        "config_version": 2, "r1_dimension": ["usertext", " Raumtyp "],
        "r1_metric": "netto", "r2_dimension": ["usertext"], "r2_metric": "NETTO"}), LAYERS),
    ("r_dims_odd", json.dumps({
        "config_version": 2, "r1_dimension": "category", "r2_dimension": ["level", "x"]}),
     LAYERS),
    ("r_dims_number_arg", json.dumps({
        "config_version": 2, "r1_dimension": ["usertext", 0], "r2_dimension": ["usertext", 2.5]}),
     LAYERS),
    ("version_string", '{"config_version": "2", "s4_dimensions": [["usertext", "K"]]}', LAYERS),
    ("version_float", '{"config_version": 2.9, "s4_dimensions": [["usertext", "K"]]}', LAYERS),
    ("version_3", '{"config_version": 3, "s4_dimensions": [["usertext", "K"]]}', LAYERS),
    ("version_0_v1_path", '{"config_version": 0, "s4_key_sequence": ["Z"]}', LAYERS),
    ("unicode_strings", json.dumps({
        "config_version": 2, "ignore_prefix": "§", "s3_parent_layer": "Grün::Ebene_01",
        "s4_dimensions": [["usertext", "Gebäude \U0001f600"], ["usertext", "Tab\there"]],
        "targets": {"Technische Gebäudeausrüstung": {"brutto": 12.0},
                    "Quote\"Back\\slash": {"netto": 1.0}}}, ensure_ascii=True), LAYERS),
    ("ignore_prefix_null", '{"config_version": 2, "ignore_prefix": null}', LAYERS),
    ("duplicate_keys", '{"config_version": 2, "tolerance_percent": 5, "tolerance_percent": 6}',
     LAYERS),
    ("err_tolerance_str", '{"config_version": 2, "tolerance_percent": "abc"}', LAYERS),
    ("err_decimal_null", '{"config_version": 2, "tolerance_percent": 20, "decimal_places": null}',
     LAYERS),
    ("err_dims_number", '{"config_version": 2, "tolerance_percent": 20, '
                        '"s4_dimensions": [5]}', LAYERS),
    ("err_layer_depth_text", '{"config_version": 2, "s4_dimensions": [["usertext", "A"], '
                             '["layer", "abc"]]}', LAYERS),
    ("err_targets_list", '{"config_version": 2, "z_height_tol": 3, "targets": ["a"]}', LAYERS),
    ("err_rdim_dict", '{"config_version": 2, "r1_dimension": {"kind": "level"}}', LAYERS),
    ("err_ignore_prefix_number", '{"config_version": 2, "tolerance_percent": 30, '
                                 '"ignore_prefix": 5}', LAYERS),
    ("err_version_text", '{"config_version": "two"}', LAYERS),
    ("err_not_object", '[1, 2]', LAYERS),
    ("err_bom", '﻿{}', LAYERS),
    ("err_trailing", '{} x', LAYERS),
]

lines = []
for name, src, layers in cases:
    f = new_form(layers)
    version = "null"
    error = "false"
    if src is not None:
        try:
            cfg = json.loads(src)
            version = str(f._apply_config(cfg))
        except Exception as ex:  # noqa: BLE001
            error = "true"
    gathered = f._gather_config()
    doc = json.dumps(gathered, ensure_ascii=False)
    fil = json.dumps(gathered, indent=2, ensure_ascii=False)
    # Round trip: the doc string must reproduce itself.
    f2 = new_form(layers)
    f2._apply_config(json.loads(doc))
    doc2 = json.dumps(f2._gather_config(), ensure_ascii=False)
    if doc2 != doc:
        print("NOT IDEMPOTENT:", name)
        print("  1:", doc)
        print("  2:", doc2)
    lines.append("new LinderoCase(%s, %s, new[] { %s }, %s, %s, %s, %s, %s)," % (
        cs(name), cs(src), ", ".join(cs(l) for l in layers), version, error, cs(doc), cs(fil),
        cs(doc2)))
write("lindero_cases.txt", lines)

# clean_label / dim_label
lines = []
for t in ["", "  Pflege ", "​A‌‍B﻿", "ﬁ", "Ａｂｃ", "x\x1c", " y　",
          "Café"]:
    lines.append("{ %s, %s }," % (cs(t), cs(module_ns["_clean_label"](t))))
write("clean_label.txt", lines)


# ── assemble Golden.g.cs ───────────────────────────────────────────────────

MEMBERS = [
    ("public static TheoryData<double, int, bool, string> Fixed => new()", "fixed.txt"),
    ("public static TheoryData<double, int, string> General => new()", "general.txt"),
    ("public static TheoryData<double, string> Repr => new()", "repr.txt"),
    ("public static TheoryData<string, bool, double> PyFloatParse => new()", "pyfloat.txt"),
    ("public static TheoryData<double, string, string> SlopeFormat => new()", "slope_format.txt"),
    ("public static TheoryData<double, string> SlopeDescribe => new()", "slope_describe.txt"),
    ("public static TheoryData<string, string, bool, double> SlopeText => new()", "slope_text.txt"),
    ("public static readonly MixedGridExpected VolumesMixed = new()", "volumes_mixed.txt"),
    ("public static TheoryData<string, int, string, string> LevelKeys => new()", "levelkeys.txt"),
    ("public static TheoryData<string?, string, int, bool> Baquiano => new()", "baquiano.txt"),
    ("public static TheoryData<string?, string[], string> Pregonero => new()", "pregonero.txt"),
    ("public static readonly LinderoCase[] LinderoCases =", "lindero_cases.txt"),
    ("public static TheoryData<string, string> CleanLabel => new()", "clean_label.txt"),
]

out = [
    "// <auto-generated>",
    "// Generated by Golden/generate_golden.py from the Python tools. Do not edit;",
    "// rerun: python src/RhinoGuire.Core.Tests/Golden/generate_golden.py",
    "// </auto-generated>",
    "#nullable enable",
    "namespace RhinoGuire.Core.Tests.Golden;",
    "",
    "internal static partial class GoldenData",
    "{",
]
for header, name in MEMBERS:
    out.append("    " + header)
    out.append("    {")
    out.extend("        " + line for line in SNIPPETS[name])
    out.append("    };")
    out.append("")
out.append("}")
with open(os.path.join(HERE, "Golden.g.cs"), "w", encoding="utf-8", newline="\n") as f:
    f.write("\n".join(out) + "\n")
print("wrote", os.path.join(HERE, "Golden.g.cs"))
