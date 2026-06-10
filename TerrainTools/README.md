# TerrainTools — Terrain Grading Suite

> **Status:** v0.1 implemented. The shared `_core` engine and all three tools
> (PadGrader, WayGrader, CutFillReport) are built and registered in `launch.py`
> with root launch shims. The WIP names are still in use (see D8). See
> [`PLAN.md`](PLAN.md) for the implementation plan and [`DECISIONS.md`](DECISIONS.md)
> for the decision log (open questions resolved at the bottom).
>
> Headless tests for the RhinoCommon-free engine parts (slope conversions and
> grid-prism volumes) live in [`_core/tests/test_headless.py`](_core/tests/test_headless.py)
> and run under plain CPython: `python TerrainTools/_core/tests/test_headless.py`.

A suite of Rhino 8 (CPython 3) tools for **modifying terrains** modelled as
Surfaces or Meshes. The suite lets the user place **building pads** and
**ways/paths** that cut and fill the terrain along controllable grading slopes,
then **compare** the original vs. modified ground and quantify **cut & fill**
volumes with on-screen charts and Excel / PNG export.

All tools follow the existing RhinoGuire conventions: modeless Eto.Forms
windows, the shared `ui/theme.py` design system, `openpyxl` for Excel, and
Eto `Drawable` → `Bitmap` for chart PNGs. They reuse the Z-projection engine
first written in `MeshTools/WrapeMeshOnMesh/Sebucan.py`.

## Planned tools (WIP names — will be renamed once workflows settle)

| WIP name | Role | Input | Output |
|----------|------|-------|--------|
| **`_core`** (`grading_core`) | Shared engine: terrain model, slope math, grading, volumes, mesh build, reporting. Not a user tool. | — | — |
| **PadGrader** | Place a building pad (closed boundary at a target elevation); generate cut/fill grading regions (daylight slopes) around it. | Closed planar curve(s) + terrain | New graded mesh + cut/fill summary |
| **WayGrader** | Define a way by its centerline polyline; persistent window for width, cross slope, and cut/fill grading slopes; generate a graded corridor. | Polyline(s) + terrain | New graded mesh + cut/fill summary |
| **CutFillReport** | Compare original vs. modified terrain, compute cut & fill, show charts, export to Excel + PNG. | Two terrains (or a grading result) | KPIs, colored mesh, charts, .xlsx, .png |

A proposed (also provisional) LatAm-themed naming option is recorded in
[`DECISIONS.md`](DECISIONS.md) under "Naming".

## Conventions inherited from the repo

- **Runtime:** Rhino 8, CPython 3 (`#! python3`), RhinoCommon (`Rhino.Geometry`),
  `rhinoscriptsyntax`, `scriptcontext`.
- **UI:** `Eto.Forms` modeless `forms.Form` with
  `Owner = Rhino.UI.RhinoEtoApp.MainWindow`; styled via `ui/theme.py`.
- **Dependencies:** `openpyxl` only (declared with the `# r: openpyxl` header).
  No new external packages.
- **Registration:** each tool gets a folder + `README.md`, an entry in
  `launch.py`/`install.py`, a `launch_<name>.py` shim, a toolbar button in
  `ui/RhinoGuire.rui`, and a row in the top-level `README.md`.

## How to use these docs

1. Read [`DECISIONS.md`](DECISIONS.md) to understand *why* the suite is shaped
   this way. Amend it as decisions change.
2. Hand [`PLAN.md`](PLAN.md) to Claude Code (in VS Code) to build the suite
   phase by phase.
3. Keep both docs updated as the implementation teaches us things — they are the
   single source of truth for this suite.
