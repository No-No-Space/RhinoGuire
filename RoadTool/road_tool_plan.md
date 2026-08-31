# Trocha — Road-on-Terrain Tool — Design Spec (rev. 3)

Target: **RhinoCommon**, Rhino 8 Python 3, built as part of the **RhinoGuire** suite.
Location: **`RhinoGuire/RoadTool/Trocha.py`** — own top-level folder inside `RhinoGuire/`
(sibling to `AreaMeasurer/`, `DataExporterImporter/`, `TerrainTools/`), so the standard
path-bootstrap reaches `ui/theme.py` and `TerrainTools/_core/` the same way every other tool does.
Goal: from a user-drawn centerline, generate a road of a given width that **conforms to the terrain
in 3D**, has a **real thickness** (reads as a solid slab in 3D views), never lifts off the ground,
and supports a **create → update** workflow keyed to the source curve. Terrain may be a **NURBS
surface/polysurface, mesh, or SubD** (detect and dispatch).

**Changes in rev. 2 (from your review):** output is now a **thin slab solid**, not a zero-thickness
ribbon (#1); added an explicit **contact rule** so roads never "fly" over dips (#2); tag prefix
`AAJ_` → **`RG_`** (#3); junctions handled by **Boolean-union of the slabs** (#4); cross-section
profiles and terrain-change detection **deferred to post-test config** (#5, #6); centerline-replaced
recovery folded into the self-healing model with one addition (#7, see §5).

**Changes in rev. 3 (suite-integration pass):** codename **Trocha**, living inside `RhinoGuire/`
(not a standalone project); terrain sampling now **reuses `TerrainTools/_core/terrain.py`'s
`TerrainModel.project_z()`** instead of a new dispatcher — this also picks up SubD support for free;
UI is now a **persistent modeless Eto window** (like WayGrader/PadGrader), replacing the raw
command-verb UX in §3; relationship to WayGrader clarified in §1.

---

## 1. Locked decisions
- API: RhinoCommon (full geometry + per-object user strings + undo records).
- Terrain: `Brep`/`Surface`/`Mesh`/`SubD`, sampled via `TerrainTools/_core/terrain.py`'s
  `TerrainModel.project_z()` (§4 step 4) — no separate dispatcher.
- Road form: conforms in 3D and stays in contact with the terrain (§4a).
- Output: **smooth single-surface top + downward thickness → a thin slab solid**. The *top face*
  stays a single crease-free surface; the thickness gives it 3D presence.
- Junctions: **merge** overlapping road slabs with a Boolean union, as a separate re-runnable step
  (§10) — the individual per-curve slabs remain the editable source of truth.
- Cross-section (camber/curbs) and terrain-change detection: designed-for but **deferred** until the
  core tool is tested; later exposed as options in the tool menu.
- **Relationship to `TerrainTools/WayGrader`**: different intent, not a duplicate. WayGrader *grades
  the ground* — it modifies/replaces terrain to fit a corridor and reports cut/fill volumes for
  earthwork design. Trocha *drapes a solid road object onto existing terrain without modifying it* —
  a presentation/documentation object, create→update-linked to its centerline, mergeable at
  junctions. They may look similar in a screenshot; the terrain-mutation question is what tells them
  apart.
- UI: a **persistent modeless Eto window**, matching WayGrader/PadGrader — not raw Rhino command
  verbs (§3 updated accordingly).

---

## 2. Core insight — clean *and* grounded
Two requirements pull in opposite directions: a **clean/smooth** top (approximate the terrain, few
control points) vs. **staying on the ground** (follow every bump exactly). We satisfy both by
splitting them across the slab:

1. **Top face** = a *smoothed approximation* of the terrain (crease-free, presentation-quality).
2. **Thickness + embedment** = we sink the slab into the terrain deep enough that even where the
   smooth top rises slightly over a dip, the **underside is still below grade** — so there is never a
   visible air gap. The smoothing tolerance and the thickness are chosen together (§4a).

Net effect: the visible top is clean, and the road always reads as sitting in/on the ground.

---

## 3. User-facing UI — persistent Eto window (mirrors WayGrader)

Single entry point, `Trocha.py`, opens a **modeless `forms.Form`** (`.Show()`, never `.ShowModal()`),
`Owner = Rhino.UI.RhinoEtoApp.MainWindow`, styled from `ui/theme.py`. Stays open while you edit
params and re-generate, matching the WayGrader/PadGrader pattern instead of re-typing commands.

- **1 — Centerline:** "Pick / Re-pick Centerline" (`rs.GetObject`, curve filter). Picking a
  tagged curve loads its stored params into the form (update mode); an untagged curve starts fresh
  (create mode) — same underlying flow either way, just no separate command verbs.
- **2 — Terrain:** "Select Terrain" (`rs.GetObject`, filter mesh | surface | polysurface | subd) —
  wraps it in `TerrainTools._core.terrain.TerrainModel` (§1, §4 step 4).
- **3 — Parameters** (editable, then Generate/Regenerate): `width`, `thickness`, plus the advanced
  knobs from §7 (`sample_step`, `fit_tol`, `top_rise`, `margin`, `smooth_center`, `corner_radius`)
  behind an "Advanced" disclosure, matching WayGrader's layout style.
- **Generate / Regenerate** button (`_t.BTN_CALC`) — runs create-or-update (§5) for the current
  centerline; wraps the doc edit in `BeginUndoRecord/EndUndoRecord` so one Ctrl+Z reverts it.
- **Remove** button — deletes the current centerline's slab and clears its tags.
- **Update All** button — regenerates every `RG_ROAD`-tagged road in the document (run after the
  terrain changes).
- **Merge** button (§10) — Boolean-union all (or selected) road slabs into a combined presentation
  solid; re-runnable, kept separate so per-road editing/update stays intact.
- Status label reporting result / warnings (reuse the `_set_status` pattern), **Close** button
  (`_t.BTN_CLEAR`).

---

## 4. Geometry pipeline (per road)
Input: centerline `C`, terrain `T`, width `W`, thickness `Th`.

1. **Prep the rail.** If `C` is a polyline / has kinks, optionally `Rebuild`/fit to a smooth degree-3
   curve (`smooth_center`, default on). Corner handling — §8.
2. **Station the path.** `Curve.DivideByLength`/`DivideByCount`, spacing = `sample_step` (e.g. 1–5 m).
   Keep point + **plan tangent** at each station.
3. **Plan offsets.** `perp = Vector3d.CrossProduct(tangent, Vector3d.ZAxis)` (unit). Left = pt+perp·W/2,
   Right = pt−perp·W/2.
4. **Sample terrain Z (the "pull") via the shared engine.** Wrap `T` once in
   `TerrainTools._core.terrain.TerrainModel(T)` (handles Mesh/Surface/Polysurface/SubD coercion
   internally) and call `.project_z(x, y)` per station point — same cached vertical-raycast
   projector used by PadGrader/WayGrader/Sebucan. No hit (outside the terrain footprint) → `None`
   → flag station (§8).
   - Raise sampled Z by `top_rise` (§4a) so the top clears convex bumps.
5. **Smooth each edge.** Fit Left/Right 3D samples to smooth curves:
   `Curve.CreateInterpolatedCurve(pts,3)` then `.Fit(3, fit_tol, 0)` (or `Rebuild`). `fit_tol` is the
   smoothness dial and feeds the contact rule (§4a).
6. **Loft the top.** `Brep.CreateFromLoft([left,right], Point3d.Unset, Point3d.Unset,
   LoftType.Normal, False)[0]` → single crease-free top face. (Two rails also make the top conform
   *across* the width on cross-sloped ground.)
7. **Thicken to a slab.** `Brep.CreateFromOffset(top, -Th, solid=True, extend=True, tol)` → closed
   slab (top face preserved, plus bottom + side walls). Verify `IsSolid`.
8. **Emit.** Add to the roads layer; write tags (§5); optionally assign the
   `RG_Technical_Colour_NoEdges` display mode via `ObjectAttributes.SetDisplayModeOverride(dm, vpId)`
   so it's clean on creation.

### 4a. Contact rule — no flying ribbons (#2)
Let `d` = the max approximation deviation of the smoothed top from the sampled terrain (measurable:
sample the fitted curve vs. the raw stations). Then choose:
- `top_rise ≥ d` → the smooth top never sinks below a convex bump (terrain won't poke through).
- `Th ≥ top_rise + d + margin` → the slab underside stays **below grade through the dips**, so the
  road never lifts off. Practical default: `Th` a few × `top_rise`.
Also: smaller `sample_step` + a bounded `fit_tol` keep `d` small in the first place. This is the knob
set that guarantees continuous ground contact while keeping the top smooth. (Offset in step 7 is along
the surface normal, so on steep grades the *vertical* burial ≈ `Th·cos(slope)` — fine for road
gradients; note it if very steep terrain appears.)

**Defaults for `fit_tol`/`top_rise`/`margin` (bug fix, 2026-07-15):** these three derive from
**`sample_step`** (`fit_tol ≈ 5%`, `top_rise`/`margin ≈ 2%` of it), *not* from
`doc.ModelAbsoluteTolerance` as earlier revisions of this spec said. `sample_step` is a physical,
document-scaled length the user already sets in real model units (like width/thickness);
`ModelAbsoluteTolerance` is a numerical-precision setting that can legitimately be loose (0.1+) on a
large site model without implying the road should float/deviate by tens of centimeters — scaling off
it directly inflated a requested 0.1m thickness to 1.3m and floated the slab off the terrain on a
loose-tolerance file. `tolerance` is kept only as a gentle floor (`× 0.1` — deliberately weak so it
only catches a degenerate near-zero `sample_step`, not re-dominate at realistic "loose" tolerances
like the 0.1 that caused this bug) and for its one legitimate remaining use: the geometric-coincidence
`tol` argument passed into the Brep offset (step 7) and Boolean-union (§10) operations — that use was
never wrong.
Also fixed the same day: `Brep.CreateFromOffsetFace`'s offset sign in step 7 was assumed to always
push downward, but a loft's face normal direction depends on which way the centerline was drawn
(CW vs CCW) — the implementation now tries both offset directions and keeps whichever solid actually
extends downward, rather than trusting a fixed sign.

---

## 5. Create / update state model
Key everything off **user strings** (they travel with the geometry and survive save). Prefix `RG_`:

On the **centerline**:
- `RG_ROAD = "1"` · `RG_ROAD_WIDTH` · `RG_ROAD_THICK` · `RG_ROAD_TERRAIN` (GUID) · `RG_ROAD_CHILD` (GUID).

On the **generated slab**:
- `RG_ROAD_PARENT` (centerline GUID).

Flow:
- **Create**: tag both; store width/thickness/terrain so updates need no re-prompt.
- **Update**: read `RG_ROAD_CHILD` → delete it → rebuild from the *current* centerline + stored params
  → write the new child GUID. One click (optional "change width/thickness" prompt).
- **Self-healing**: if the stored child GUID is stale (slab deleted by hand), scan for any object whose
  `RG_ROAD_PARENT` == the centerline GUID; if none, create fresh.

**On #7 (centerline replaced with a new GUID):** partly covered, with one addition. The self-heal
above matches *child → parent GUID*, which works while the curve keeps its identity. If the curve
itself is **replaced** (explode/rejoin, re-draw), the new curve is untagged and its GUID differs, so
pure GUID matching can't find the orphaned slab. To make this case work we add a **re-link fallback**:
when you pick an untagged curve, look for an orphaned slab (a `RG_ROAD_PARENT` that no longer
resolves) whose footprint is closest to the picked curve, and offer to re-link + rebuild. Same
self-healing spirit, just matched by geometry instead of GUID. (A small document-level registry is the
alternative; user-strings + proximity keeps it stateless — recommended.)

---

## 6. Module layout — `RhinoGuire/RoadTool/`

Separate **pure geometry** (testable, no doc side effects) from **document/UI** ops, same split
`TerrainTools/_core` uses and for the same reason:

- `_core/config.py` — defaults (`sample_step`, `fit_tol`, `top_rise`, `thickness`, `margin`, layer,
  display-mode name, `RG_` tag keys).
- `_core/geometry.py` — `build_slab(center_crv, width, thickness, terrain_model, cfg) -> Brep`.
  RhinoCommon geometry in/out only, no `RhinoDoc` → unit-testable. Takes a
  `TerrainTools._core.terrain.TerrainModel` instance directly (§1, §4 step 4) — no local terrain
  module.
- `_core/state.py` — user-string read/write, parent↔child resolve, scan + proximity re-link.
- `_core/junctions.py` — Boolean-union merge (§10).
- `Trocha.py` — the tool entry point: path bootstrap (mirrors `Sebucan.py`/`WayGrader.py`), the Eto
  form (§3), selection (`rs.GetObject`/`rs.GetObjects`), undo records, doc writes. One file, per the
  suite's "one tool = one file" convention, importing from `_core/` and from
  `TerrainTools._core.terrain`.
- `README.md` — mirrors the other tools' README (workflow, known approximations, smoke-test
  checklist), per `CLAUDE.md`/`TerrainTools/PLAN.md` convention.

Keeping `geometry.py`/`junctions.py` free of `RhinoDoc` is what lets Claude Code write real tests —
same rationale as `TerrainTools/_core`, headless-runnable per `CLAUDE.md`'s "Running Tests" section.

---

## 7. Parameters & knobs
`width`, `thickness` (per road, stored) · `sample_step` · `fit_tol` (smoothness) · `top_rise`,
`margin` (contact rule) · `smooth_center` / optional `corner_radius` · `assign_display_mode`.
`fit_tol`/`top_rise`/`margin` default off **`sample_step`**, not `doc.ModelAbsoluteTolerance` — see
§4a's 2026-07-15 bug-fix note. `doc.ModelAbsoluteTolerance` is used only as the geometric-coincidence
`tol` argument for the underlying Brep offset/Boolean operations, plus a tiny floor on the three knobs
above so they can't collapse toward zero on an unreasonably small `sample_step`.

---

## 8. Edge cases & mitigations
- **Tight corners**: inner offset edge can self-intersect. Mitigate with `corner_radius` fillet
  (`Curve.CreateFilletCornersCurve`) or a min-turn-radius warning; test the offset curve for
  self-intersection before lofting.
- **Centerline past terrain extent**: stations with no hit → clamp to terrain bbox, or drop+interpolate
  neighbors, or warn and skip.
- **Mesh holes/gaps**: ray miss → interpolate Z from neighbors; warn if too many misses.
- **Vertical/overhang terrain**: multiple hits → topmost Z.
- **Closed-loop roads**: loft/offset with closed handling; keep edge curves closed/periodic.
- **Offset/thicken failure** (self-overlap on tight, thick roads): fall back to extrude-down of the top
  edges + cap, or reduce `Th` with a warning.
- **Units/tolerance**: everything off `doc.ModelAbsoluteTolerance`.

---

## 9. Test / acceptance plan
Run `build_slab` against a matrix and assert:
- Terrain types: flat plane, sloped NURBS, bumpy mesh.
- Paths: straight, gentle curve, S-curve, tight corner, closed loop.
- **Clean**: top face is a single face; no interior naked edges on the top.
- **Solid**: slab `IsSolid == True`.
- **Contact (no flying)**: sample the slab underside vs. terrain → underside Z ≤ terrain Z everywhere
  (buried); top Z ≥ terrain Z everywhere (no poke-through). Both within `margin`.
- **Params**: width and thickness hold within tolerance.
- **Update**: replaces (not duplicates) the child; `RoadUpdateAll` re-drapes all after the terrain moves.
- **Merge**: `RoadMerge` on two crossing roads yields one closed solid (`IsSolid`, single shell).

---

## 10. Junctions / merge (#4) and deferred items
**Junctions — chosen approach:** because roads are now solids, merge crossing/overlapping roads with
`Brep.CreateBooleanUnion([slabs], tol)` into one continuous body. Keep it as the separate **Merge**
step (§3) so the individual per-curve slabs stay editable and update-linked; the merged solid is a
derived, re-runnable output (optionally tagged `RG_ROAD_MERGED`). Add corner fillets/cleanup at
junctions as a refinement once the base union is solid.

**Deferred until after concept test (your #5, #6):**
- Cross-section profiles (camber/crown, curbs, thickened variants) exposed as tool-menu config.
- Terrain-change detection (auto-flag roads whose stored terrain moved and prompt an update).

---

## 11. Conventions & registration (mirrors `TerrainTools/PLAN.md` §8)

Conventions are already documented in `RhinoGuire/CLAUDE.md` ("Key Conventions") — Trocha follows
them as-is rather than duplicating them here:

- **Script header**: `#! python3` + `__title__`/`__doc__` block (mirror `Sebucan.py` lines 1–30).
- **Path bootstrap**: same `_rg_root = normpath(join(dirname(__file__), "..", ".."))` pattern
  (Trocha is one level deep — `RoadTool/Trocha.py` — same depth as `AreaMeasurer/Lindero.py`).
- **Eto rules**: modeless only, `Owner` set, no kwargs on .NET constructors, `StackLayout` for
  dynamic rows — see `CLAUDE.md` → "Eto.Forms rules".
- **Geometry convention**: original centerline/terrain are never modified; new slabs go on a
  dedicated layer (e.g. `RoadTool::Roads`); `sc.doc.Views.Redraw()` after every doc edit.

**Registration steps**, once the tool runs:

1. `launch.py` — add `"RG_Trocha": os.path.join(RHINOGUIRE_ROOT, "RoadTool", "Trocha.py")`.
2. `launch_trocha.py` shim at repo root (copy `launch_sebucan.py`).
3. `ui/RhinoGuire.rui` — toolbar button, macro `! _-RunPythonScript ".../launch_trocha.py"`.
4. Top-level `README.md` — add a row under "Tools" and to the repository-structure tree.
5. `CLAUDE.md` — add `RG_Trocha` to the "Available keys" list.

---

## 12. Open items carried from this review

- **Terrain reliability**: per Aksel, the existing `TerrainTools` (PadGrader/WayGrader/CutFillReport)
  "don't work 100% in reality" yet — treat `_core/terrain.py` as reused-but-not-bulletproof; if
  Trocha's smoke tests surface a sampling bug, fix it in `_core` (benefits all four tools) rather
  than working around it locally.
- **Merge-tag scope**: confirm whether `RG_ROAD_MERGED` output should also be excluded from a future
  `RoadUpdateAll`-equivalent sweep (it's derived, not a source road) — flag if unclear when building §10.
