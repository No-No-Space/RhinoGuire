# Lindero — Footprint Area Calculator

**Version:** 0.1
**Author:** Aquelon
**Requires:** Rhino 8, CPython 3 engine (`#! python3`)

---

## What it does

Lindero calculates the **footprint area** of Rhino objects — the plan area as seen from directly above (XY projection). This is distinct from the surface area that Rhino's built-in `Area` command computes, which sums all faces of an object.

The tool runs as a **modeless window**, so Rhino stays fully interactive while the form is open. You can select objects, change layers, and run multiple calculations without reopening the script.

---

## The three scenarios

### S1 — Selected Objects
Calculates the footprint of each selected object **individually**, with no overlap handling. Use this when you want a quick breakdown of specific objects regardless of where they sit relative to each other.

- Results: area per object + sum total.
- Objects are labelled using a user text key (optional); falls back to the Rhino object name or a short GUID.

### S2 — By Layer
Calculates the footprints of all objects on a chosen layer and **merges overlapping regions** using a Boolean Union. This prevents double-counting when two objects physically share the same floor area.

- Results: individual area per object + combined layer total (after overlap removal).
- The difference between the individual sum and the combined total reveals how much area is shared.

### S3 — Layer Hierarchy
Targets a parent layer whose **sublayers each represent one floor or level**. Overlaps are removed within each sublayer independently. The grand total is the **sum of all sublayer totals** — floors are additive (standard Gross Floor Area logic).

Two user text keys drive the breakdown:

| Key | Purpose | Example value |
|---|---|---|
| Object Key | Individual or small-group name | `"Room Name"` → `"Kitchen"` |
| Group Key | Larger classification | `"Department"` → `"Private"` |

- Results: per object, per group subtotal (combined footprint per group within a level), per sublayer total, grand total.

---

## Footprint calculation logic

### Step-by-step

For each **Brep or Extrusion** object:

1. Iterate every face of the solid.
2. Evaluate the face normal at its centre point.
3. Keep only faces where `|normal.Z| > 0.9` — this means the face is within approximately **26 degrees of horizontal**.
4. Among those horizontal faces, identify the one(s) at the **lowest centroid Z** (the actual bottom of the geometry).
5. Extract the **outer border curve** of each bottom face.
6. Project that curve straight down onto the XY plane (`Z = 0`).
7. Compute the area enclosed by the projected curve using `AreaMassProperties`.

For **closed planar curves** (e.g. a room outline drawn as a flat polyline), the curve itself is projected to Z=0 and its enclosed area is used directly.

For any object where no horizontal face is found (e.g. a pure cylinder on its side), the tool falls back to the **XY bounding box** as an approximation.

### What `|normal.Z| > 0.9` means

The `0.9` is **not a distance in model units**. It is a threshold on the dot product between the face normal and the world Z axis. A value of 0.9 means the face normal deviates less than ~26° from vertical, i.e. the face itself is less than ~26° from horizontal. This tolerance handles faces that are nominally flat but have small modelling imperfections.

---

## Known limitation — L-shaped sections (overhangs and cantilevers)

The behaviour depends on which way the L faces.

### Case A — L-shaped floor plan (plan view is L-shaped)
The bottom face of the solid **is** the L-shape. The code finds it correctly and the footprint is exact.

```
Plan view (top down):
┌─────┐
│     │
│     ├───┐
│     │   │
└─────┴───┘
```
**→ Handled correctly.**

---

### Case B1 — L-shaped section, wider at the base (a step or recess going upward)

```
Side section view:
█████████████
█████████
█████████
```

The bottom face spans the full width of the base. The narrowing at the top does not affect which face is identified as the bottom. The footprint is the full base width.

**→ Handled correctly.**

---

### Case B2 — L-shaped section, wider at the top (cantilever / overhang)

```
Side section view:
█████████████
      ███████
      ███████
```

The bottom face covers only the **narrow base** (the column or stem). The overhanging portion at the top has its own horizontal face, but that face is at a **higher Z** than the base — so the code ignores it.

**→ The overhanging area is NOT included in the footprint.**

The result is the footprint of the base only, not the full "shadow" the object casts when viewed from above.

#### When does this matter in practice?

In typical architectural space planning (rooms modelled as simple vertical extrusions), this limitation is never triggered — every room block sits flat on the floor and its bottom face equals its ceiling face in plan area.

It becomes relevant if:
- Objects represent built elements with cantilevers (e.g. balconies, overhanging floors modelled as a single solid).
- Objects are placed at different Z levels but grouped into a single Brep.

#### Potential future fix

Replacing the bottom-face method with a true **top-down silhouette projection** would capture Case B2 correctly. This requires computing `Brep.GetSilhouette()` or projecting all faces to Z=0 and taking the outer boundary of the union — a more expensive but geometrically complete approach.

---

## Overlap removal (Scenarios S2 and S3)

When multiple objects are processed together, their projected footprint curves are passed to `Rhino.Geometry.Curve.CreateBooleanUnion()` at the model's absolute tolerance. This merges all overlapping and adjacent regions into a single set of non-overlapping closed curves, then sums their areas.

If the Boolean Union fails (degenerate curves, non-planar input, tolerance issues), the tool falls back to a plain sum of individual areas and marks the result with `[union failed — sum shown]` in the output. In this case the total may be slightly over-reported if footprints genuinely overlap.

---

## Refresh Model

Because the form is modeless, the Rhino model can change while it is open. Click **Refresh Model** to re-scan all layers and user text keys and update the dropdowns and key ComboBoxes without reopening the script.

---

## To-Do

- Export results to Excel.
- Load an Excel file with target area goals and compare calculated values against them.
- Investigate silhouette-based footprint for Case B2 (cantilever / overhang solids).
