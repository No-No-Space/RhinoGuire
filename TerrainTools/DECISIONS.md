# TerrainTools — Decision Log

Lightweight ADR-style record of the choices that shape the grading suite.
Each entry: **decision**, **why**, **alternatives considered**, **status**.
Append new entries; supersede old ones rather than deleting, so history stays.

Created: 2026-06-10. Owner: Aksel.

---

## D1 — Architecture: suite of tools over a shared engine

**Decision.** Build a shared core library (`TerrainTools/_core/`, package
`grading_core`) plus three thin tool front-ends (PadGrader, WayGrader,
CutFillReport). The pads tool and the ways tool both produce a *design
heightfield* through the same engine; the analysis tool consumes terrains/grids.

**Why.** Pads and ways share ~80% of the work (terrain sampling, slope-to-
daylight grading, mesh building, cut/fill). A shared engine avoids duplicating
that math and matches the existing repo split (`MeshTools/…`, `AreaMeasurer/…`).
Each tool stays small and independently launchable from the toolbar.

**Alternatives.** (a) One multi-tab modeless window — fewer launch entries but a
large single file and weaker reuse. (b) Pad tool only first — defers value of
ways and analysis. Rejected in favour of the suite, but the engine is designed
so a unified window could wrap it later if desired.

**Status.** Accepted.

---

## D2 — Slope input: ratio / percent / degrees, user-selectable

**Decision.** Every slope field in the UI accepts a value plus a unit toggle:
**H:V ratio**, **percent (%)**, or **degrees (°)**. Internally everything is
stored as a single canonical gradient `m = rise / run = V / H` (vertical over
horizontal).

Conversions (`grading_core/slope.py`):

| Input | → canonical `m` | `m` → display |
|-------|-----------------|---------------|
| Ratio `H:V` | `m = V / H` | `H:V = 1 : m`  (or `1/m : 1`) |
| Percent `p%` | `m = p / 100` | `p = 100·m` |
| Degrees `θ` | `m = tan(θ)` | `θ = atan(m)` |

Cut slope and fill slope are stored as **positive magnitudes**; the cut-vs-fill
sign is decided per location by comparing design Z to terrain Z (see D5).

**Why.** Civil/earthwork users think in H:V, road/landscape users in %, and some
in degrees. A canonical internal value keeps the geometry code unit-free.

**Alternatives.** Lock to one convention (simpler UI, but forces unit math on
the user). Rejected.

**Status.** Accepted. Note: civil ratio convention here is **H:V** (e.g. `2:1`
= 2 horizontal to 1 vertical = 50% = 26.57°). Label this explicitly in the UI.

---

## D3 — Output model: new mesh, original kept intact

**Decision.** Terrain is read as Surface **or** Mesh, coerced **internally to a
mesh** for sampling. Grading outputs a **new mesh** on a separate layer
(e.g. `TerrainTools::Graded`). The original terrain object is never modified.

**Why.** Non-destructive editing is expected in design work and is required for
a clean before/after cut-fill comparison (we always have both surfaces). Mesh
output is robust for arbitrary graded shapes; rebuilding a trimmed NURBS surface
from a graded heightfield is approximate and not worth it for v1.

**Alternatives.** (a) Also fit a NURBS surface to the result — deferred to the
backlog. (b) Modify in place — rejected (destructive; would force a hidden
backup for comparison).

**Status.** Accepted. NURBS re-fit is a backlog item (see PLAN → Backlog).

---

## D4 — Terrain sampling: mesh + vertical raycast, reusing Sebucan

**Decision.** Sample existing-ground elevation with a vertical-raycast projector
`project_z(x, y)` built on `Rhino.Geometry.Intersect.Intersection.MeshRay`,
exactly the technique already proven in
`MeshTools/WrapeMeshOnMesh/Sebucan.py` (`_make_projector`, `coerce_to_mesh`).
This logic moves into `grading_core/terrain.py`; Sebucan may later import it
(non-breaking refactor, optional).

**Why.** Already works in the repo, handles Mesh/SubD/Surface/Polysurface, and
caches results by (x, y). No new dependency.

**Alternatives.** `Mesh.ClosestPoint` (not strictly vertical), surface `ClosestPoint`
in UV (slower, trimmed-surface edge cases). Raycast is the right primitive for
"elevation directly below/above a plan position".

**Status.** Accepted.

---

## D5 — Grading method: analytic design heightfield over a regular grid

**Decision.** Represent both the design and the existing ground as **Z values on
a shared regular XY grid** (cell size = user "resolution"). For each grid node
the design elevation is computed analytically:

- **Inside a pad / carriageway:** design Z = the platform/section elevation.
- **Outside, in the grading skirt:** let `d` = horizontal distance to the
  feature edge and `z_edge` = feature elevation at the nearest edge point.
  Compare existing terrain `z_t` at the node to `z_edge`:
  - `z_t > z_edge` → **cut**: `design = min(z_t, z_edge + d · m_cut)`
  - `z_t < z_edge` → **fill**: `design = max(z_t, z_edge − d · m_fill)`
  - `z_t ≈ z_edge` → `design = z_t`
  The `min`/`max` clamp is what makes the slope stop automatically at the
  **daylight (catch) line**; beyond it `design = z_t`, so cut/fill there is zero.

This handles mixed cut-and-fill around one feature (cut on the uphill side, fill
on the downhill side) with no special cases.

**Why.** A grid heightfield is simple, robust, easy to mesh, and makes volume a
trivial per-cell sum (D6). The analytic clamp avoids iterative slope-staking.

**Alternatives.** Iterative offset-curve + plane/terrain intersection to find the
catch line (classic civil "grade to daylight"). More exact at the catch line but
much more code and failure modes. The grid method's accuracy is controlled by
resolution and is adequate for design-stage work. Revisit if precision demands it.

**Status.** Accepted for v1.

---

## D6 — Cut/fill volume: grid prism (cell) method

**Decision.** With design Z and terrain Z on the same grid (cell area
`A = cell²`), per cell `Δ = design − terrain`:
`fill += max(Δ,0)·A`, `cut += max(−Δ,0)·A`. Net `= fill − cut`.
Cut/fill areas = cell counts × A where `|Δ|` exceeds a small tolerance.

**Why.** Directly consistent with the design representation (D5), standard
earthwork "grid method", trivially exportable per-cell, and accuracy scales with
the same resolution control the user already sets.

**Alternatives.** Average-end-area along stations (better suited to corridors —
offered as an *additional* per-station report for WayGrader), mesh boolean
volume (fragile). Grid method is the common denominator.

**Status.** Accepted. WayGrader additionally reports per-station areas/volumes
for a mass-haul-style curve.

---

## D7 — Reporting/visualization reuses existing patterns

**Decision.** Charts are drawn with Eto `forms.Drawable` `.Paint` handlers and
exported to PNG via `drawing.Bitmap` (as in `AreaMeasurer/Lindero.py`). The
modified mesh is tinted by cut/fill depth using `mesh.VertexColors` with a
legend, and the viewport is exported through the same `ViewCaptureToFile`
approach used in `DataVisualization/Chivito.py`. Excel uses `openpyxl`, with
last-used folders remembered via `theme.prefs_get/prefs_set`.

**Why.** Consistency with the rest of RhinoGuire; no new dependencies
(`matplotlib` is not reliably present in Rhino's CPython).

**Status.** Accepted.

---

## D8 — Naming: WIP now, rename after workflows settle

**Decision.** Use the descriptive WIP names **PadGrader / WayGrader /
CutFillReport** during development. Rename once the workflows are validated.

**Provisional LatAm-themed option** (consistent with Arriero/Lindero/Sebucan),
for consideration only:

- PadGrader → **Banco** (earthwork *benching/terracing* a platform) or
  **Explanada** (a levelled graded platform).
- WayGrader → **Trazado** (a road/path *alignment*) or **Sendero** (trail).
- CutFillReport → **Balanza** (cut/fill *balance*) or **Cubicación**
  (the civil term for earthwork volume take-off).

**Status.** Open. Revisit before first release.

---

## D9 — Open questions resolved during the v0.1 build

These resolve the four open questions from `PLAN.md` §12, decided while
implementing. Recorded here so the rationale isn't lost.

1. **Volume basis for ways.** The **grid-prism method (D6) is the primary KPI**
   for both pads and ways — the corridor is rasterized onto the same regular
   grid as pads, so `volumes.cut_fill` and CutFillReport behave identically for
   either. The **average-end-area** computation is kept as an *additional*
   per-station output (`volumes.per_station`) feeding the mass-haul curve.

2. **`_core/report.py` Eto boundary.** `_core` stays **Eto-free**.
   `report.py` holds only the `openpyxl` writer (and imports openpyxl lazily so
   PadGrader/WayGrader can import it without the dependency). The
   `Drawable → PNG` chart export lives **inside CutFillReport**, not in `_core`.

3. **Multiple ways per window.** v1 handles **one way at a time**. Regenerate
   replaces the previous result by default (a "Replace previous result on
   Regenerate" checkbox); unchecking it appends additional corridors. A
   per-way list UI is deferred.

4. **Model units.** Volumes/areas are labelled from `sc.doc` via
   `terrain.model_unit_label()` (e.g. `m³`, `m²`) rather than hard-coding "m";
   the engine itself stays unit-agnostic (works in document model units).

**Additional implementation decisions:**

- **Overlapping pad skirts** use the **nearest-feature-wins** rule (a skirt node
  is graded against the closest pad's edge). This resolves most overlaps and is
  documented as a v1 approximation (PLAN §5 / §4.3).
- **Zero/negative slope** on a side means **no skirt** there (design = terrain,
  i.e. a vertical edge) rather than an infinite flat bench.
- A small **shared widgets module** (`TerrainTools/_widgets.py`, Eto-based, *not*
  part of `_core`) holds the slope-input row, labeled-row helper and a few
  document helpers (layer creation, default cell size) so the three tools stay
  DRY.

**Status.** Accepted for v0.1.
