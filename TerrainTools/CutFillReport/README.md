# CutFillReport — Compare · Quantify · Export

Part of the **TerrainTools** suite. Compares an **original** vs a **modified**
terrain (or reads the last PadGrader / WayGrader result), computes **cut & fill
volumes**, shows KPIs and charts, builds a cut/fill-tinted **map mesh** with a
legend, and exports the numbers to **Excel** and the charts to **PNG**.

> Requires **openpyxl** for Excel export — Rhino installs it automatically from
> the `# r: openpyxl` header on first run.

## Two source modes

- **From last grading (Pad/WayGrader)** — reads
  `sc.sticky['terraintools_last_grade']`, which already holds the design and
  terrain grids (and per-station data for ways). Just press **Compute**.
- **Two terrains (Original vs Modified)** — pick the *Original* and *Modified*
  objects and a *cell size*; optionally a closed *analysis boundary* to restrict
  the comparison. Both are sampled on a shared grid over their overlap (or the
  boundary), forming the delta grid.

## Results

- **KPI panel** — Cut, Fill, Net, Cut area, Fill area, Balance ratio
  (`fill/cut`), cell size and grid extent. Volumes/areas are labelled with the
  document's model units (e.g. `m³` / `m²`).
- **Charts** (Eto `Drawable`):
  - **Cut vs Fill** bar with a net/balance readout.
  - **Depth distribution** histogram (cut = teal-blue, fill = salmon).
  - **Mass haul** — cumulative net volume along stations (ways only).
- **Show cut/fill map** — builds a fresh mesh from the graded region, tints it by
  per-vertex depth (blue cut → neutral → red fill) on the
  `TerrainTools::CutFillMap` layer, and shows a **legend** with the depth scale.

## Exports

- **Export to Excel (.xlsx)** — *Summary* KPIs sheet + a *Per Station* sheet for
  ways. The last-used folder is remembered (`cutfill_export_xlsx`).
- **Export charts as PNG** — renders the dashboard (`Drawable` → `Bitmap`) to a
  PNG (`cutfill_export_png`).
- **Capture viewport as PNG…** — opens Rhino's `ViewCaptureToFile` dialog so you
  can capture the tinted map + legend with full control over the settings.

## Method

Volumes use the grid-prism method (D6): over the graded region, per cell
`Δ = design − terrain`, `fill += max(Δ,0)·cell²`, `cut += max(−Δ,0)·cell²`,
`net = fill − cut`. The mass-haul curve uses average-end-area between stations.

## Smoke-test checklist

- [ ] *From last grading* KPIs match the originating tool's reported Cut/Fill.
- [ ] *Two terrains* over an overlapping pair produces sensible cut/fill.
- [ ] Excel opens with a Summary sheet (and Per Station for a way).
- [ ] Charts PNG is written and readable.
- [ ] Tinted map mesh + legend read correctly (blue cut, red fill).
