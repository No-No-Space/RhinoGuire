# Trocha — Road on Terrain

| Specification | Details |
| :--- | :--- |
| **Tool Name** | **Trocha** (`RG_Trocha`) |
| **Category** | Road Infrastructure & 3D Draping (Non-destructive) |
| **Runtime** | Rhino 8 (CPython 3) · Modeless `Eto.Forms` UI |
| **Dependencies** | Pure RhinoCommon (Zero external dependencies) |
| **Launch Command** | `! _-RunPythonScript "<path-to-repo>/RhinoGuire/launch_trocha.py"` |
| **Input** | Centerline Curve + Terrain (Mesh, SubD, Surface, Polysurface) |
| **Output** | 3D Solid road slab, create→update links, Boolean-union merged solid |

Drapes a solid road slab onto a terrain from a user-drawn centerline, in a
**persistent window**: pick, set width/thickness, and Generate. The top face
is a clean, crease-free surface; the slab is buried deep enough into the
terrain that it never "flies" over a dip. Non-destructive — the terrain is
never modified. Create → update is linked to the centerline, so re-running
replaces the slab instead of duplicating it.

Full design spec: [`road_tool_plan.md`](road_tool_plan.md).


## Workflow

1. **Select Terrain** — Mesh, SubD, Surface, or Polysurface.
2. **Pick / Re-pick Centerline** — a curve. Picking an already-tagged
   centerline loads its stored width/thickness (update mode); an untagged
   curve starts fresh (create mode). If an orphaned slab is found nearby
   (its original centerline was replaced), you'll be offered to clear it.
3. **Parameters**: *Width*, *Thickness*, *Sample spacing* (along-centerline
   station step), *Smooth centerline before draping*.
4. **Generate / Regenerate** — builds the slab and reports the result
   (solid vs. surface-only fallback, and any warnings).
5. **Remove** — deletes the current centerline's slab and clears its tags.
6. **Update All** — re-drapes every tagged road in the document (run after
   the terrain changes).
7. **Merge** — Boolean-unions all built roads into one presentation solid on
   a `::Merged` sub-layer; re-runnable, replaces its own previous output.

## How draping works

See `road_tool_plan.md` §4/§4a for the full rationale. In short: the
centerline is stationed, offset left/right by half the width, and each rail
is sampled onto the terrain via `TerrainTools._core.terrain.TerrainModel`
(shared with PadGrader/WayGrader/Sebucan — not reimplemented here). The
raw sampled deviation from a smoothed fit (`d`) drives two knobs: the top
face is raised by at least `d` so it never dips below a convex bump, and the
slab thickness is sized so the underside stays buried through the worst dip.
The two smoothed rails are lofted into a single top face, then offset into a
closed solid via `Brep.CreateFromOffsetFace`.

## Known approximations (v1)

- **Terrain misses**: stations outside the terrain footprint are filled by
  interpolating between the nearest valid neighbors, not flagged per-station
  (a single count is reported).
- **Tight corners**: no automatic corner-radius fillet yet (`corner_radius`
  is a config knob but not exposed in the UI); very tight/thick roads can
  fail to thicken — Trocha retries at half thickness a few times, then falls
  back to returning the top surface only with a warning, per
  `road_tool_plan.md` §8.
- **Re-link fallback**: because width/thickness live on the old (now-gone)
  centerline, re-linking an orphaned slab clears it rather than recovering
  its original parameters — you re-enter them and Generate fresh.
- **Cross-section (camber/curbs) and terrain-change detection**: deferred
  per `road_tool_plan.md` §1/§10.

## Notes

- No external dependencies.
- Modeless, persistent window — Rhino stays interactive.
- Distinct from `TerrainTools/WayGrader`: WayGrader *grades/modifies*
  terrain for earthwork design; Trocha *drapes without modifying* it.

## Smoke-test checklist

- [ ] A straight road over flat terrain drapes cleanly, reads as a solid.
- [ ] A curved road over rolling/bumpy terrain stays in contact everywhere
      (no visible gap under the slab, no poke-through on the top).
- [ ] A tight corner does not crash the build (falls back gracefully if the
      offset self-intersects).
- [ ] Regenerate on the same centerline replaces the slab, not duplicates it.
- [ ] Update All re-drapes all tagged roads after moving/editing the terrain.
- [ ] Merge on two crossing roads yields one closed solid.
- [ ] Remove deletes the slab and clears the tags (a fresh Generate on the
      same curve creates cleanly again).
