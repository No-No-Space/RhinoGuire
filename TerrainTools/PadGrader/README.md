# PadGrader — Building Pad Grading

Part of the **TerrainTools** suite. Places one or more building pads (closed
planar boundaries at a target elevation) on a terrain and grades cut/fill
slopes around them down to the daylight line. Outputs a **new graded mesh** on a
separate layer — the original terrain is never modified.

## Workflow

1. **Select Terrain** — a Mesh, SubD, Surface (incl. trimmed) or Polysurface.
   It is coerced to a mesh internally and sampled with a vertical raycast.
2. **Select Pad Curve(s)** — one or more **closed, planar** curves. Open or
   non-planar curves are reported and skipped.
3. **Pad elevation** — choose how each pad's platform Z is set:
   - *From terrain (mean / min / max)* — statistic of the ground under the pad.
   - *Explicit Z* — a single elevation typed in.
4. **Grading slopes** — set the **Cut** and **Fill** slopes. Each accepts a
   value plus a unit toggle: **H:V** (run:rise, e.g. `2:1`), **%**, or **°**.
   The live label shows the equivalent in all three units.
5. **Options**
   - *Cell size* — grid resolution in model units (auto-suggested from the
     terrain size; smaller = more accurate and slower).
   - *Max reach* — how far skirts may extend past a pad before being capped.
   - *Output layer* — default `TerrainTools::Graded` (nested layer created
     automatically).
6. **Generate** — builds the graded mesh, adds it to the output layer, and
   reports **Cut / Fill / Net** volumes and the balance ratio.
7. **Open in CutFillReport** — hands the result to CutFillReport (via
   `sc.sticky`) for charts and Excel/PNG export.

## How grading works (D5)

For each grid node outside a pad, at horizontal distance `d` from the nearest
pad edge (edge elevation `z_edge` = pad platform Z), with terrain `z_t`:

- `z_t > z_edge` → **cut**: `design = min(z_t, z_edge + d·m_cut)`
- `z_t < z_edge` → **fill**: `design = max(z_t, z_edge − d·m_fill)`
- otherwise → `design = z_t`

The `min`/`max` clamp stops the slope automatically at the **daylight line**;
beyond it the design equals the existing ground (zero cut/fill). This naturally
gives cut on the uphill side and fill on the downhill side of a single pad.

Volumes use the grid-prism method (D6): per cell `Δ = design − terrain`,
`fill += max(Δ,0)·cell²`, `cut += max(−Δ,0)·cell²`.

## Edge cases handled / flagged

- Open or non-planar boundary curves (skipped, reported).
- Pad fully above (all fill) or fully below (all cut) the terrain.
- **Overlapping pads** — a skirt node is governed by the **nearest** pad
  (nearest-feature-wins merge rule; an approximation for v1).
- Slope never daylights within *Max reach* — flagged ("still on a slope at the
  grid edge — increase Max reach").
- Boundary or grid partly off the terrain footprint — those nodes are excluded
  from volumes and flagged.

## Notes

- No external dependencies (Excel export lives in CutFillReport).
- Modeless window — Rhino stays interactive while it is open.

## Smoke-test checklist

- [ ] Terrain as a Mesh and as a trimmed Surface are both accepted.
- [ ] A pad on a hillside cuts uphill and fills downhill, with a visible
      daylight line on both sides.
- [ ] A pad raised above the terrain produces all-fill; sunk below, all-cut.
- [ ] Two overlapping pads grade without a seam crashing the run.
- [ ] Reported Cut/Fill match CutFillReport for the same result.
