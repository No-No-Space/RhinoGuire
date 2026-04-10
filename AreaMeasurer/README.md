# Lindero — Footprint Area Calculator

**Version:** 0.4
**Author:** Aquelon
**Requires:** Rhino 8, CPython 3 engine (`#! python3`), `openpyxl` (auto-installed)

---

## What it does

Lindero calculates the **footprint area** of Rhino objects — the plan area as seen from directly above (XY projection). This is distinct from the surface area that Rhino's built-in `Area` command computes, which sums all faces of an object.

The tool runs as a **modeless window**, so Rhino stays fully interactive while the form is open. You can select objects, change layers, and run multiple calculations without reopening the script.

---

## Scenarios

### S1 — Selected Objects
Calculates the footprint of each selected object **individually**, with no overlap handling. Use this for a quick breakdown of specific objects regardless of their relative positions.

- Results: area per object + sum total.
- Objects are labelled using a user text key (optional); falls back to the Rhino object name or a short GUID.

---

### S2 — By Layer
Calculates the footprints of all objects on a chosen layer and **merges overlapping regions** using a Boolean Union, preventing double-counting when two objects share floor area.

- Results: individual area per object + combined layer total (after overlap removal).
- An **overlap warning** appears when the individual sum exceeds the combined total, reporting the overlapping area so you can decide whether double-counting is intentional.

---

### S3 — Layer Hierarchy
Targets a parent layer whose **sublayers each represent one floor or level**. Overlaps are removed within each sublayer independently. The grand total is the **sum of all sublayer totals** — floors are additive (standard Gross Floor Area logic).

Two user text keys drive the breakdown:

| Key | Purpose | Example attribute | Example value |
| --- | --- | --- | --- |
| Object Key | Individual or small-group name | `SpaceType` | `"Office"`, `"Kitchen"` |
| Group Key | Larger classification | `Department` | `"Private"`, `"Services"` |

- Results: per object, per group subtotal (combined footprint per group within a level), per sublayer total, grand total.
- **Overlap warnings** are shown per level when objects share footprint area, including a breakdown of cross-group overlap (objects from different groups occupying the same footprint).

---

### S4 — Custom Aggregation

User-defined hierarchy of attribute keys. You define the levels yourself — for example: `Domain → Main Group → Subgroup → Room Type`. The UI lets you add or remove levels dynamically.

Per sublayer, footprints are merged within each leaf group using a **Boolean Union** (same overlap logic as S3). Leaf totals are then summed across all floors.

Results are displayed as an **indented tree** in the text area:

```
  ▸ Residential                              12,450.0000
    ▸ Housing                                 8,200.0000
      ▸ Apartment                             5,100.0000
          Bedroom                             2,400.0000
          Living Room                         2,700.0000
      ▸ Studio                                3,100.0000
          Studio Unit                         3,100.0000
  ▸ Commercial                               4,250.0000
    ...
```

Each non-leaf node shows the cumulative area of all its descendants.

**Using S4 as a data source for R1 / R2:** once keys are defined in S4, you can point R1 and R2 at specific levels of the S4 hierarchy using the level index steppers in Settings — instead of pulling from S3's two-key model.

---

### R1 — Room Analysis

Aggregates individual object areas **across all floors**, grouped by a chosen key value. Compares each room type's total against a **Room Target Key** (set in Settings).

Data source (configurable in Settings): **S3 keys** (default) or **S4 hierarchy** at a chosen level index.

- Results displayed as **bullet charts**: one row per unique key value.
- Each chart row shows the measured total, the target value, and the ±% deviation.
- Colour zones on the bar indicate whether the result is within tolerance (yellow = below goal, orange = above goal).
- Missing target keys are reported as warnings in the tab without blocking the calculation.
- Charts can be **exported as PNG** using the "Export Chart as PNG" button.

---

### R2 — Group Analysis

Same as R1 but aggregates by a **group-level key** and compares against a **Group Target Key** (set in Settings).

Data source (configurable in Settings): **S3 keys** (default) or **S4 hierarchy** at a chosen level index.

- Results displayed as bullet charts, one row per unique group key value.
- Same warning and PNG export behaviour as R1.

---

## Settings tab

### Program Key Mapping

Defines which user text keys hold the target area values used in R1 and R2:

| Setting | Used by | Meaning |
| --- | --- | --- |
| Room Target Key | R1 | Attribute key whose value is the target area for that room type |
| Group Target Key | R2 | Attribute key whose value is the target area for that group |

Dropdowns are populated from all keys present in the model.

### Tolerance

**Global Tolerance (%)** — symmetric percentage applied to both R1 and R2 charts. Default: 10%. Range: 0–50%.

A tolerance of 10% means the yellow band spans `goal × 0.90` to `goal`, and the orange band spans `goal` to `goal × 1.10`.

### R1 / R2 Data Source

Controls which parent layer and keys feed the R1 and R2 bullet charts:

| Option | Behaviour |
| --- | --- |
| **S3 keys** (default) | R1 uses S3 Object Key; R2 uses S3 Group Key; parent layer from S3 tab |
| **S4 hierarchy** | R1 and R2 pull from S4's key sequence using the level indices below |

**R1 Room Level** — 1-based position in the S4 key sequence to use as the room aggregation key (default: 1).

**R2 Group Level** — 1-based position in the S4 key sequence to use as the group aggregation key (default: 2).

### Configuration

- **Save Config** — saves all settings (key mapping, tolerance, S4 key sequence, R1/R2 source) to a JSON file.
- **Load Config** — loads a previously saved config and populates all fields.

Config file structure:

```json
{
  "room_target_key": "TargetArea",
  "group_target_key": "GroupTarget",
  "tolerance_percent": 10.0,
  "s4_parent_layer": "Building",
  "s4_key_sequence": ["Domain", "Main Group", "Subgroup", "Room Type"],
  "r1r2_source": 0,
  "r1_level_index": 1,
  "r2_level_index": 2
}
```

`r1r2_source`: `0` = S3 keys, `1` = S4 hierarchy.

---

## Bullet chart legend

Each row in R1/R2 is drawn as follows (left to right):

```
[Room label]  ░░░░▓▓▓[████████████]░░░▓▓▓░░░░   87.5/100.0
              ↑   ↑  ↑            ↑  ↑         -12.5%  [m²]
              │   │  └─ measured  │  └─ upper tolerance marker
              │   └─ lower        └─ goal line (dark vertical bar)
              │     tolerance
              └─ chart start (0)
```

| Element | Colour | Meaning |
|---|---|---|
| Background | Light grey | Full chart range (0 → goal × 1.35) |
| Yellow band | Yellow | Below-goal tolerance zone: `goal × (1−tol)` to `goal` |
| Orange band | Orange | Above-goal tolerance zone: `goal` to `goal × (1+tol)` |
| Measured bar | Blue-grey | Actual measured area (0 → measured) |
| Goal line | Dark, 2 px | Target value |
| Tolerance markers | Grey, 1 px | Lower and upper tolerance boundaries |

---

## Write Area to Objects

The **"Write Area to Objects"** button opens an inline panel below the button row (S1, S2, S3 only):

1. Choose or type the user text key to write to (default: `Area`).
2. Click **Confirm Write**.

The calculated footprint area is written as a user string to each measured object using `SetUserString`. If the key does not yet exist on an object, it is created. The operation is one-way — no automatic sync; click again to update after recalculating.

Status bar confirms: `Area written to N object(s) using key 'Area'`.

---

## Export to Excel

Available for S1, S2, S3, and S4.

**S1, S2, S3** workbooks contain two sheets:

| Sheet | Contents |
| --- | --- |
| **Objects** | One row per measured object with GUID, layer/level, key values, and footprint area |
| **Summary** | Scenario parameters, totals, group breakdowns, and any overlap warnings |

**S4** workbooks contain two sheets:

| Sheet | Contents |
| --- | --- |
| **Leaf Data** | Flat table — one row per unique key path (leaf), with one column per key level plus area. Useful for pivot tables. |
| **Tree Summary** | Indented hierarchy showing the cumulative area at each node, matching the text area output. |

Overlap warnings are highlighted in amber in Summary sheets.

---

## Export Chart as PNG

Available when the R1 or R2 tab is active. Renders the full bullet chart to a 900 px wide PNG file at the path you choose. The chart height scales automatically with the number of entries (54 px per row).

---

## Footprint calculation logic

### Step-by-step

For each **Brep or Extrusion** object:

1. Iterate every face of the solid.
2. Evaluate the face normal at its centre point.
3. Keep only faces where `|normal.Z| > 0.9` — the face is within approximately **26° of horizontal**.
4. Among those horizontal faces, identify the one(s) at the **lowest centroid Z** (the actual bottom of the geometry).
5. Extract the **outer border curve** of each bottom face.
6. Project that curve straight down onto the XY plane (`Z = 0`).
7. Compute the area enclosed by the projected curve using `AreaMassProperties`.

For **closed planar curves** (e.g. a room outline drawn as a flat polyline), the curve itself is projected to Z=0 and its enclosed area is used directly.

For any object where no horizontal face is found (e.g. a cylinder on its side), the tool falls back to the **XY bounding box** as an approximation.

### What `|normal.Z| > 0.9` means

The `0.9` is a threshold on the dot product between the face normal and the world Z axis — not a distance in model units. It means the face normal deviates less than ~26° from vertical, i.e. the face is less than ~26° from horizontal. This tolerance handles faces that are nominally flat but carry small modelling imperfections.

---

## Known limitation — L-shaped sections (overhangs and cantilevers)

### Case A — L-shaped floor plan (plan view is L-shaped)

The bottom face of the solid *is* the L-shape. The code finds it correctly and the footprint is exact. **→ Handled correctly.**

### Case B1 — L-shaped section, wider at the base

```text
Side section view:
█████████████
█████████
█████████
```

The bottom face spans the full width of the base. The narrowing at the top does not affect which face is identified as the bottom. **→ Handled correctly.**

### Case B2 — L-shaped section, wider at the top (cantilever / overhang)

```text
Side section view:
█████████████
      ███████
      ███████
```

The bottom face covers only the narrow base. The overhanging portion at the top has its own horizontal face, but at a *higher Z* — so the code ignores it. **→ The overhanging area is NOT included in the footprint.**

This is rarely an issue in typical space-planning models (rooms modelled as simple vertical extrusions). It only matters if objects represent built elements with cantilevers, or multiple Z-level forms merged into a single Brep.

A future fix would replace the bottom-face method with a true top-down silhouette projection (`Brep.GetSilhouette()` or a full union of all faces projected to Z=0).

---

## Overlap removal (S2, S3, and S4)

Projected footprint curves are passed to `Rhino.Geometry.Curve.CreateBooleanUnion()` at the model's absolute tolerance. If the Boolean Union fails, the tool falls back to a plain sum and marks the result with `[union failed — sum shown]`.

---

## Refresh Model

Click **Refresh Model** to re-scan all layers and user text keys and update every dropdown and ComboBox without reopening the script. Existing selections are preserved where possible.

---

## To-Do

- Silhouette-based footprint for Case B2 (cantilever / overhang solids).
- Highlight out-of-tolerance objects in the Rhino viewport from R1/R2.
- Excel export for R1/R2 (room/group analysis results with target comparison).
