# WayGrader — Way / Path Corridor Grading

Part of the **TerrainTools** suite. Grades a way/path corridor onto a terrain
from its **centerline**, in a **persistent window**: edit the parameters and
*Regenerate* without re-picking. Outputs a **new graded mesh** on a separate
layer — the original terrain is never modified.

## Workflow

1. **Select Terrain** — Mesh, SubD, Surface or Polysurface.
2. **Pick / Re-pick Centerline** — a polyline or curve. Its plan geometry is the
   way alignment; grading is built from cross-sections sampled along it.
3. **Corridor parameters** (all editable, then Regenerate):
   - *Width* — full carriageway width.
   - *Crossfall* — **Crown** (falls both ways from the centre) or **Single**
     (constant fall left→right), with its own slope value + unit toggle.
   - *Cut* and *Fill* slopes — for the skirts each side (H:V / % / °).
4. **Options**
   - *Station spacing* — along-centerline sample step (for the mass-haul curve).
   - *Cell / cross step* — grid resolution.
   - *Max reach* — lateral cap past the carriageway edge.
   - *Output layer* — default `TerrainTools::Graded`.
   - *Replace previous result on Regenerate* — deletes the prior mesh so
     repeated tweaks don't pile up geometry (on by default).
5. **Generate / Regenerate** — builds the corridor mesh and reports
   Cut / Fill / Net volumes and the number of stations.
6. **Open in CutFillReport** — for charts (including the **mass-haul** curve)
   and Excel/PNG export.

## How grading works

The corridor is rasterized onto the same regular grid used by pads, so volumes
and CutFillReport behave identically for pads and ways. For each grid node the
nearest point on the centerline gives a station, a draped centerline elevation
(`centerline_z = project_z`, the v1 *drape-on-terrain* profile), and a lateral
normal. With signed lateral offset `s`:

- `|s| ≤ width/2` → **carriageway**: `design = centerline_z − |s|·m_cross`
  (crown) or `centerline_z − s·m_cross` (single crossfall).
- `|s| > width/2` → **skirt**: `d = |s| − width/2`, edge Z = carriageway Z at the
  near edge; apply the D5 cut/fill clamp against the terrain.

Per-station cut/fill cross-section areas are accumulated and turned into a
running (mass-haul) volume via the average-end-area method (D6), surfaced in
CutFillReport.

## Known approximations (v1)

- **Concave bends**: nearest-point assignment naturally clips overlapping
  inside-of-curve skirts (min |Δ|), but tight corners are approximate.
- **Profile**: only *drape on terrain* in v1 (constant design grade, max
  longitudinal slope and vertical curves are backlog items).
- **Multiple ways**: handled one at a time; Regenerate replaces the last result
  (uncheck *Replace* to append several corridors).

## Notes

- No external dependencies (Excel export lives in CutFillReport).
- Modeless, persistent window — Rhino stays interactive.

## Smoke-test checklist

- [ ] A way with crown crossfall over rolling terrain grades on both sides.
- [ ] Changing *Width* and pressing Regenerate rebuilds without re-picking.
- [ ] *Replace previous result* removes the old mesh on Regenerate.
- [ ] Mass-haul curve appears in CutFillReport for the result.
