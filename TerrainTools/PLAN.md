# TerrainTools — Implementation Plan

Hand this document to Claude Code (VS Code) to build the terrain-grading suite.
Read [`DECISIONS.md`](DECISIONS.md) first for the rationale; this file is the
*how*. Build in the phase order at the end. Keep both docs in sync as you learn.

- **Runtime:** Rhino 8, CPython 3. Header on every entry script: `#! python3`.
- **No new dependencies.** Only `openpyxl` (declared via `# r: openpyxl`).
- **Reuse, don't reinvent:** the projector and mesh coercion already exist in
  `MeshTools/WrapeMeshOnMesh/Sebucan.py`; the UI system in `ui/theme.py`; Excel +
  Drawable→PNG patterns in `AreaMeasurer/Lindero.py`; viewport PNG + vertex-color
  legend in `DataVisualization/Chivito.py`.

---

## 1. Goals (what the user must be able to do)

1. **Building pads.** Pick one or more closed planar curves, set a pad elevation,
   set cut & fill grading slopes; generate the graded terrain (pad platform +
   slope skirts to daylight) as a new mesh, keeping the original intact, and see
   the cut/fill totals.
2. **Ways.** Draw a centerline polyline; in a **persistent window** edit width,
   cross slope (crossfall), and cut/fill grading slopes; generate a graded
   corridor mesh; re-edit parameters and regenerate without reopening.
3. **Compare & quantify.** Compare original vs. modified terrain, compute cut &
   fill volumes, show the result graphically (KPIs, charts, a cut/fill-tinted
   mesh + legend), and export the numbers to **Excel** and the charts/legend to
   **PNG**.

---

## 2. Conventions to follow (copy these patterns)

**Script header** (mirror `Sebucan.py` lines 1–30):
```python
#! python3
# -*- coding: utf-8 -*-
# __title__ = "PadGrader"
# __doc__ = """Version = 0.1
# Date    = YYYY-MM-DD
# Author: Aquelon - aquelon@pm.me
# ... description / how-to / changelog ...
```

**Path bootstrap to reach `ui/` and `_core/`** (mirror `Sebucan.py` lines 41–48):
```python
import sys as _sys, os as _os
_rg_root = _os.path.normpath(_os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "..", ".."))
if _rg_root not in _sys.path:
    _sys.path.insert(0, _rg_root)
from ui import theme as _t
from TerrainTools import _core  # or: from TerrainTools._core import terrain, grading, ...
```
(Depth `".." , ".."` assumes tools live in `TerrainTools/<Tool>/`. Adjust the
number of `".."` to land on the repo root, as Sebucan does from its 2-deep path.)

**Modeless window** (mirror `SebucanForm`):
```python
class PadGraderForm(forms.Form):
    def __init__(self):
        super().__init__()
        self.Title = "PadGrader — Building Pad Grading"
        self.Resizable = True
        self.Padding = drawing.Padding(12)
        self.BackgroundColor = _t.BG
        self.Owner = Rhino.UI.RhinoEtoApp.MainWindow
        ...
def main():
    PadGraderForm().Show()      # .Show() = modeless; never .ShowModal()
if __name__ == "__main__":
    main()
```

**UI atoms** from `ui/theme.py`: `_t.lbl`, `_t.btn`, `_t.hint`,
`_t.section_header`, colors (`_t.BG`, `_t.TEXT`, `_t.BTN_CALC`, `_t.BTN_CLEAR`,
chart colors `_t.CHART_*`), fonts (`_t.F_HEAD`, `_t.F_SANS`, `_t.F_SANS_B`),
status colors `_t.status_color("ok"|"warn"|"error"|"info")`, and folder memory
`_t.prefs_get(key, fallback)` / `_t.prefs_set(key, path)`.

**Selection** via `rhinoscriptsyntax`: `rs.GetObject(s)` with filters
(`rs.filter.curve`, `.mesh`, `.surface`, `.polysurface`, `.subd`), `preselect=True`.

**Adding geometry:** `sc.doc.Objects.AddMesh(mesh)`; set layer with
`rs.ObjectLayer(id, layer)`; redraw with `sc.doc.Views.Redraw()`.

---

## 3. Target file structure

```text
TerrainTools/
├── README.md                 # done
├── PLAN.md                   # this file
├── DECISIONS.md              # done
├── _core/
│   ├── __init__.py
│   ├── slope.py              # unit conversions (ratio/%/deg ↔ canonical m)
│   ├── terrain.py            # TerrainModel: coerce→mesh, project_z, bbox, grid
│   ├── grading.py            # pad + corridor design-heightfield builders
│   ├── volumes.py            # cut/fill from two Z-grids; KPIs; per-station
│   ├── meshbuild.py          # grid/section → Rhino Mesh; vertex-color tinting
│   └── report.py             # openpyxl writer + Drawable→Bitmap PNG helpers
├── PadGrader/
│   ├── PadGrader.py
│   └── README.md
├── WayGrader/
│   ├── WayGrader.py
│   └── README.md
└── CutFillReport/
    ├── CutFillReport.py
    └── README.md
```

> Optional non-breaking refactor: move `coerce_to_mesh` and `_make_projector`
> out of `Sebucan.py` into `_core/terrain.py`, and have Sebucan import them.
> Only do this if it stays behaviourally identical and Sebucan still runs.

---

## 4. Shared engine — `grading_core` API spec

All functions are pure Python + RhinoCommon; **no Eto imports in `_core`** (keeps
it testable and reusable). Coordinates are world XY; elevation is world Z.

### 4.1 `slope.py`
```python
def to_gradient(value: float, unit: str) -> float:
    """unit ∈ {'ratio_hv','percent','degrees'} → canonical m = V/H (>0)."""

def from_gradient(m: float, unit: str) -> float:
    """canonical m → display value in the requested unit."""

def format_slope(m: float, unit: str) -> str:
    """Human label, e.g. '2:1', '50%', '26.6°'."""
```
- `ratio_hv`: input is H per 1 V (e.g. `2.0` means 2:1) → `m = 1/value`.
  Accept either a float `H` or a string `"H:V"`; document clearly.
- `percent`: `m = value/100`. `degrees`: `m = tan(radians(value))`.
- Guard `value <= 0`; treat `m → 0` as flat (no skirt, never daylights → capped).

### 4.2 `terrain.py`
```python
class TerrainModel:
    def __init__(self, obj_id_or_geometry): ...      # coerce Surface/Mesh/SubD/Polysurface → Mesh
    @property
    def mesh(self) -> Rhino.Geometry.Mesh: ...
    @property
    def bbox(self) -> Rhino.Geometry.BoundingBox: ...
    def project_z(self, x: float, y: float):         # vertical raycast; cached; None if outside
        ...
    def sample_grid(self, x0, y0, nx, ny, cell):     # → 2D list of Z (or None) using project_z
        ...
```
- Reuse `coerce_to_mesh` and `_make_projector` from `Sebucan.py` verbatim
  (they handle Mesh/SubD/BRep, dual up/down ray, and caching).
- `make_grid_bounds(curve_or_points, reach) -> (x0, y0, nx, ny, cell)`: free
  function that returns grid extents = feature bbox expanded by `reach`,
  snapped to `cell`. `reach` defaults to an auto estimate (see grading) capped by
  a user max.
- `point_in_curve(curve, pt, tol) -> bool`: wrap
  `curve.Contains(pt, Rhino.Geometry.Plane.WorldXY, tol)` and test for
  `PointContainment.Inside`. Curve must be closed & planar; validate and warn.
- `dist_and_edge_z(curve, pt)`: use `curve.ClosestPoint(pt, out t)` →
  `cp = curve.PointAt(t)`; return `(hypot(pt.X-cp.X, pt.Y-cp.Y), cp.Z)`.

### 4.3 `grading.py`
```python
def grade_pad(terrain: TerrainModel, boundary_curve, pad_z, m_cut, m_fill,
              cell, max_reach) -> GradeResult:
    """Build a design-Z grid for one pad and its slope skirts (see D5)."""

def grade_pads(terrain, boundary_curves, pad_z_mode, m_cut, m_fill, cell, max_reach):
    """Multiple pads in one grid; combine by taking, per cell, the design Z that
    represents the most material moved toward the pads (document the merge rule:
    nearest-feature wins; overlapping skirts resolved by min |Δ|)."""

def grade_corridor(terrain, centerline, params) -> GradeResult:
    """Build a design-Z grid for a way corridor (see §6)."""
```
- `pad_z_mode`: explicit float, or `'mean'|'min'|'max'` of terrain under the pad.
- **Per-node design Z (the core rule, D5):** for a node at distance `d` from the
  feature edge with edge elevation `z_edge` and terrain `z_t = project_z(node)`:
  - inside feature → `z_edge` (pad platform / carriageway Z at that station)
  - `z_t > z_edge` → cut → `min(z_t, z_edge + d*m_cut)`
  - `z_t < z_edge` → fill → `max(z_t, z_edge − d*m_fill)`
  - else → `z_t`
  - if `z_t is None` (outside terrain) → design = `z_edge ± d*m` uncapped, flag.
- **Auto `max_reach` estimate:** `reach ≈ |Δz_max| / min(m_cut, m_fill)` where
  `Δz_max` is the max terrain-vs-pad elevation difference within the feature bbox;
  cap by the user's "Max reach" field. Flag nodes still on a slope at the grid
  edge (slope never daylighted within reach).
- **`GradeResult`** (plain dataclass-like dict/obj, no Eto):
  ```text
  GradeResult:
    x0, y0, cell, nx, ny
    z_design : 2D list[float|None]
    z_terrain: 2D list[float|None]
    region_mask: 2D list[bool]     # True where node is inside feature or skirt
    flags: list[str]               # e.g. 'capped at max_reach', 'off-terrain'
  ```

### 4.4 `volumes.py`
```python
def cut_fill(grade: GradeResult, tol=1e-4) -> dict:
    """Grid-prism method (D6). Returns:
       { 'cut_volume', 'fill_volume', 'net', 'cut_area', 'fill_area',
         'balance_ratio', 'cell_area', 'min_delta', 'max_delta',
         'depth_histogram': [(bin_lo, bin_hi, count), ...] }"""

def per_station(centerline, grade, station_spacing) -> list[dict]:
    """For WayGrader: per-station cut/fill area + running (mass-haul) volume."""
```
- `Δ = z_design − z_terrain` per cell where both defined and `region_mask`.
- `fill += max(Δ,0)*A`, `cut += max(−Δ,0)*A`, `A = cell²`.
- `balance_ratio = fill_volume / cut_volume` (guard divide-by-zero).
- Histogram bins over `[min_delta, max_delta]` for the depth-distribution chart.

### 4.5 `meshbuild.py`
```python
def grid_to_mesh(grade: GradeResult, only_region=True) -> Rhino.Geometry.Mesh:
    """Quad grid → triangulated Mesh from z_design. Skip cells with None or
       (only_region and not region_mask)."""

def tint_by_delta(mesh, grade, ramp) -> None:
    """Assign mesh.VertexColors from per-vertex Δ: blue=cut, red=fill, neutral≈0.
       'ramp' returns an Eto/System color for a normalized value; keep the color
       function injectable so _core stays Eto-free (pass System.Drawing colors)."""
```
- Build vertices row-major; two triangles per cell; `ComputeNormals()`; `Compact()`.
- Provide vertex-Δ by sampling `z_design − z_terrain` at each vertex.

### 4.6 `report.py`
```python
def write_xlsx(path, summary: dict, per_cell=None, per_station=None) -> None:
    """openpyxl workbook: 'Summary' KPIs sheet; optional 'Per Station' and
       'Per Cell' sheets. Mirror Lindero's styling (Font/PatternFill/Alignment,
       get_column_letter, autosize)."""

def chart_to_png(draw_fn, width, height, path) -> None:
    """Render an Eto Drawable paint routine to a drawing.Bitmap and save PNG
       (mirror Lindero.py ~line 975 'bmp = drawing.Bitmap(...)')."""
```
> `report.py` is the one `_core` module allowed to touch Eto.Drawing for the
> Bitmap export, OR keep Bitmap export inside each tool and have `report.py` hold
> only the openpyxl writer. Prefer the latter to keep `_core` Eto-free — decide
> when implementing and note it here.

---

## 5. Tool — PadGrader (building pads)

**Window layout (Eto DynamicLayout, like SebucanForm):**
1. Header + one-line hint.
2. **1 — Terrain:** "Select Terrain" button (`rs.GetObject`, filter mesh |
   surface | polysurface | subd); show name + type (reuse `_obj_type_label`).
3. **2 — Pad boundary(ies):** "Select Pad Curve(s)" (`rs.GetObjects`, filter
   curve, must be closed planar). Show count; warn on open/non-planar.
4. **3 — Pad elevation:** radio/dropdown {Explicit Z (textbox), From terrain:
   mean | min | max}.
5. **4 — Grading slopes:** two slope rows (Cut, Fill), each a textbox + a unit
   dropdown {H:V, %, °} bound through `slope.to_gradient`. Live label showing the
   other two unit equivalents (nice-to-have).
6. **Options:** Resolution / cell size (textbox, model units; default ≈ terrain
   bbox diagonal / 200, clamped to a sane min); Max reach (textbox); output layer
   name (default `TerrainTools::Graded`).
7. **Generate** button (`_t.BTN_CALC`), **Close** (`_t.BTN_CLEAR`).
8. Status label (reuse `_set_status` pattern).

**On Generate:**
- Build `TerrainModel(terrain_id)`.
- `grade = grade_pads(...)`; `mesh = grid_to_mesh(grade)`.
- Add mesh to output layer; carry over nothing destructive; `Views.Redraw()`.
- `kpi = cut_fill(grade)`; show "Cut X / Fill Y / Net Z (model³)" in status and a
  small results panel. Offer "Open in CutFillReport" (stash last `grade`/ids in
  `sc.sticky['terraintools_last_grade']` so CutFillReport can pick it up).

**Edge cases to handle & flag:** open/non-planar boundary; pad fully above or
fully below terrain (all-fill / all-cut); overlapping pads; slope never daylights
(capped — report which); boundary partly off the terrain mesh.

---

## 6. Tool — WayGrader (ways / paths) — persistent parameter window

**Centerline:** user picks a polyline/curve (`rs.GetObject`, filter curve). The
way's plan geometry is the centerline; grading is built from cross-sections
sampled along it.

**Persistent window** stays open; parameters editable; **Generate / Regenerate**
rebuilds from current values without re-picking (store the picked curve id; allow
"Re-pick centerline"). Parameters:

- **Width** (full width; internally half-width L/R — expose separate L/R as
  optional advanced fields).
- **Cross slope (crossfall):** one of {crown (both sides fall from center),
  single crossfall L→R}, value via the slope unit toggle. Edge Z =
  `centerline_z ∓ (width/2)·m_cross`.
- **Centerline profile (v1):** "Drape on terrain" = `centerline_z =
  project_z(station)`. (Backlog: constant design grade, max longitudinal slope,
  vertical curves.)
- **Cut slope** and **Fill slope** (slope unit toggles) for the skirts each side.
- **Station spacing** (= along-centerline sample step; default = cell size),
  **Resolution / cross step**, **Max reach**, **output layer**.

**Corridor algorithm (`grade_corridor`):**
1. Sample the centerline at `station_spacing`: point `P`, unit tangent `T`
   (from finite difference), horizontal normal `N = (−T.Y, T.X, 0)` normalized.
2. `centerline_z = project_z(P.X, P.Y)`.
3. For each station, march across from `−max_reach` to `+max_reach` at
   `cross_step`; at lateral offset `s` (signed), the cross point is `Q = P + s·N`:
   - `|s| ≤ width/2` → carriageway: `design_z = centerline_z − |s|·m_cross`
     (crown) or `centerline_z − s·m_cross` (single crossfall).
   - `|s| > width/2` → skirt: `d = |s| − width/2`, `z_edge` = carriageway Z at the
     near edge; apply the D5 cut/fill clamp against `project_z(Q)`.
4. Build the mesh by triangulating between consecutive stations' cross-section
   sample arrays (a ribbon). Keep a `region_mask` for cells within the daylighted
   corridor so volumes ignore untouched ground.
5. Also produce a heightfield `GradeResult` on a regular grid (rasterize the
   ribbon) so `volumes.cut_fill` and CutFillReport work uniformly — OR compute
   volumes directly from the section ribbon via average-end-area and *also* fill a
   grid for the comparison map. Decide and document; average-end-area is natural
   for corridors (D6) — expose it via `volumes.per_station` for a mass-haul curve.

**Mitered/!sharp corners:** at concave bends offset normals overlap; clip
overlapping skirt samples by taking the design Z nearest the centerline (min |Δ|).
Note this as a known approximation for v1.

**Multiple ways:** allow a list; regenerate all. Each way = its own parameter set
(consider a small list UI; v1 may handle one way at a time and append meshes).

---

## 7. Tool — CutFillReport (compare, visualize, export)

**Inputs (two modes):**
- **A — Two terrains:** pick "Original" and "Modified" objects; build a shared
  grid over their overlap (or a user-supplied analysis boundary curve), sample
  both via `TerrainModel.project_z`, form a `GradeResult`-shaped delta grid.
- **B — From last grading:** read `sc.sticky['terraintools_last_grade']` produced
  by PadGrader/WayGrader (already has design + terrain grids).

**On-screen results:**
- **KPI panel:** Cut, Fill, Net, Cut area, Fill area, Balance ratio (`fill/cut`),
  cell size, grid extent. (Units = document model units; show e.g. "m³ / m²".)
- **Charts** (Eto `forms.Drawable` `.Paint`, same approach as Lindero R1/R2 at
  `Lindero.py` ~lines 1474–1516, 975):
  - Cut vs Fill bar (with net marker).
  - Depth distribution histogram (from `cut_fill`’s `depth_histogram`).
  - For ways: mass-haul / cumulative-volume curve from `per_station`.
- **Colored mesh:** `tint_by_delta` the modified mesh (blue cut → neutral → red
  fill); add a **legend** (reuse Chivito's legend builder pattern,
  `Chivito.py` ~lines 419–536) with min/0/max depth labels.

**Exports:**
- **Excel** (`report.write_xlsx`): Summary KPIs sheet + Per-Station sheet (ways) +
  optional Per-Cell sheet. Use `SaveFileDialog` with
  `_t.prefs_get('cutfill_export_xlsx')` / `prefs_set` (mirror `Lindero.py`
  ~lines 1815–1822). Append `.xlsx` if missing.
- **PNG:** "Export charts as PNG" (Drawable→Bitmap, `Lindero.py` ~line 975) and
  "Capture viewport as PNG" for the tinted mesh + legend (Chivito opens Rhino's
  capture dialog via `Rhino.RhinoApp.RunScript("ViewCaptureToFile", False)`,
  `Chivito.py` ~line 731). Remember folders via `prefs` keys
  `cutfill_export_png` / `cutfill_legend_png`.

---

## 8. Registration (wire each tool into the suite)

For PadGrader, WayGrader, CutFillReport:

1. **`launch.py`** — add to `SCRIPTS` dict:
   ```python
   "PadGrader":     os.path.join(_ROOT, "TerrainTools", "PadGrader",    "PadGrader.py"),
   "WayGrader":     os.path.join(_ROOT, "TerrainTools", "WayGrader",    "WayGrader.py"),
   "CutFillReport": os.path.join(_ROOT, "TerrainTools", "CutFillReport","CutFillReport.py"),
   ```
2. **`install.py`** — add matching `RG_*` entries (kept for parity; file is
   marked deprecated but still lists scripts).
3. **`launch_<name>.py` shims** at repo root (copy `launch_sebucan.py`):
   ```python
   #! python3
   import sys, os
   sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
   from launch import launch
   launch("PadGrader")
   ```
   → `launch_padgrader.py`, `launch_waygrader.py`, `launch_cutfillreport.py`.
4. **`ui/RhinoGuire.rui`** — add toolbar buttons (macro:
   `! _-RunPythonScript "…/launch_padgrader.py"`). The `.rui` is XML; either edit
   it or document the manual button-add steps in `ui/README.md`.
5. **Top-level `README.md`** — add a row per tool under "Tools" and to the
   "Repository Structure" tree. Note `openpyxl` requirement for CutFillReport.
6. **`manifest.yml`** — bump description/keywords if appropriate (`terrain`,
   `grading`, `cut-fill`, `earthwork`).

---

## 9. Phased build order

- **Phase 0 — Scaffold.** Create `_core/` package + empty modules; create the
  three tool folders with skeleton modeless windows that open and close. Register
  in `launch.py` + shims. Verify each opens in Rhino.
- **Phase 1 — Engine.** Implement `slope.py`, `terrain.py` (port Sebucan
  projector/coercion), `grading.grade_pad`, `volumes.cut_fill`, `meshbuild.grid_to_mesh`.
  Unit-test the pure-Python parts headless (see §10).
- **Phase 2 — PadGrader.** Full UI + single pad → graded mesh + KPI readout.
  Then multi-pad, elevation modes, edge-case flags.
- **Phase 3 — WayGrader.** Persistent window, corridor algorithm, crossfall,
  drape profile, regenerate loop. Add `volumes.per_station`.
- **Phase 4 — CutFillReport.** Two-terrain compare + from-sticky mode; KPIs;
  charts; mesh tint + legend; Excel + PNG export.
- **Phase 5 — Polish.** Toolbar buttons, READMEs per tool, top-level README,
  performance pass (resolution warnings, progress in status label), refactor
  Sebucan to import `_core` (optional, non-breaking).

---

## 10. Test / verification plan

**Headless engine tests** (run with plain CPython where possible, or a small
Rhino-run harness for the geometry parts):
- `slope.py`: round-trip `to_gradient`/`from_gradient` for all three units;
  `2:1 ↔ 50% ↔ 26.565°`; reject `≤0`.
- `volumes.cut_fill`: a flat tilted plane vs. a horizontal pad of known size →
  compare computed cut/fill to a hand calculation (within grid tolerance).
- `grading.grade_pad` on a synthetic sloped terrain: pad on a hill → cut on the
  uphill side, fill on the downhill side; daylight line present on both; design
  == terrain beyond it.

**In-Rhino smoke tests** (manual checklist in each tool README):
- Terrain as Mesh and as trimmed Surface both accepted.
- Pad fully above (all fill) and fully below (all cut) terrain.
- Way with crown crossfall over rolling terrain; regenerate after changing width.
- CutFillReport: KPIs match the originating grade's `cut_fill`; Excel opens;
  PNGs written; tinted mesh + legend read correctly.

**Self-check before handoff of any phase:** confirm the file/line references in
this plan still resolve (they were valid at 2026-06-10 against `Sebucan.py`,
`Lindero.py`, `Chivito.py`, `ui/theme.py`).

---

## 11. Backlog / future ideas

- Sloped (ramped) pads; multi-level pads with retaining walls.
- Way longitudinal design grade: target grade, max gradient, vertical curves.
- NURBS surface re-fit of the graded mesh (D3 alternative).
- Catch-line as an actual curve output (for drawings), not just the mesh.
- Iterative offset-and-intersect grading for higher precision (D5 alternative).
- Berms / swales / ditches along ways; superelevation on curves.
- Balanced-earthwork solver: auto-pick pad elevation that makes cut ≈ fill.
- Topo contour regeneration on the modified mesh.

---

## 12. Open questions (resolve as you build; update DECISIONS.md)

1. Volume basis for ways — grid-rasterized (uniform with pads) vs. average-end-
   area (native to corridors). Plan supports both; pick the primary for the KPI.
2. `_core/report.py` Eto boundary — keep `_core` fully Eto-free and put
   Drawable→PNG in each tool, or allow `report.py` to import `Eto.Drawing`?
   (DECISIONS D7 leans Eto-free.)
3. Multiple ways with independent parameter sets in one window — needed for v1,
   or one-way-at-a-time with append? (Plan assumes append for v1.)
4. Model units / labels — read `sc.doc.ModelUnitSystem` and label volumes/areas
   accordingly rather than hard-coding "m".
