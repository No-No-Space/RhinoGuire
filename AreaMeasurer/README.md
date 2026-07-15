# Lindero — Footprint Area Calculator

## What it does

Lindero calculates the **footprint area** of Rhino objects — the plan area as seen from directly above (XY projection). This is distinct from the surface area that Rhino's built-in `Area` command computes, which sums all faces of an object.

The tool runs as a **modeless window**, so Rhino stays fully interactive while the form is open. You can select objects, change layers, and run multiple calculations without reopening the script.

---

## Layer structure contract (v2)

S3, S4, R1, and R2 read the model hierarchy **from the layer tree** (since v0.8 — see `PLAN.md` for the design decisions):

```text
07-PROPOSAL-Phase_03            ← PARENT (chosen in the UI)
├── _BuildingVolumes            ← ignored (prefix) — whole subtree excluded
├── Ebene_06                    ← LEVEL (direct child of the parent)
│   ├── _TEXT                   ← ignored (prefix)
│   ├── Diagnostik und Therapie ← CATEGORY — objects measured here
│   ├── Pflege                  ← CATEGORY
│   └── …
├── Ebene_05 … Ebene_-03        ← more levels
└── Ebene_KMA                   ← levels can have any name
```

- **Level** = each direct child of the parent.
- **Category** = the object's **own (lowest) layer**. Deeper nesting is fine — the object's own layer name is always the category.
- **Ignore prefix** (Settings, default `_`): a layer whose name starts with the prefix is excluded **together with its entire subtree**. Empty prefix = nothing ignored. The parent itself is never prefix-tested.
- Objects directly on a level layer are measured with category `—` and reported as a warning. Objects directly on the parent are **not** measured (the preview reports them).

The pure path logic lives in `_paths.py` and is headless-testable:

```sh
python AreaMeasurer/tests/test_paths.py
```

---

## Scenarios

### S1 — Selected Objects
Calculates the footprint of each selected object and **merges overlapping regions**, preventing double-counting when objects share floor area. Objects are merged per layer (and per Z band); per-layer totals are then summed.

- Results: area per object + individual sum + combined total (after overlap removal).
- An **overlap warning** appears when the individual sum exceeds the combined total.
- Objects are labelled using a user text key (optional); falls back to the Rhino object name or a short GUID.

### S2 — By Layer
Same merge logic for all objects on one chosen layer. Objects at different heights (Z-Height Tolerance, Settings) are treated as separate floors and summed.

### S3 — Layer Hierarchy
Reads the **layer structure contract** above. Click **Preview Structure** first: a dry run showing which levels/categories will count, object counts, and the ignored layers — no areas computed.

Aggregation rules:

1. Union per **(level, category)** → category totals per level.
2. **Level total** = union across all the level's objects → `Σ(categories) − level total` = **cross-category overlap**, warned per level.
3. **Grand total** = sum of level totals (floors are additive — GFA logic).
4. **Category totals (whole building)** = per category, the sum of its per-level unions.

Results are shown in two panels: the breakdown grid (levels with category children + building-wide category summary) and the object detail grid (object, category, area per level).

### S4 — Custom Aggregation
User-defined hierarchy of **dimensions**. Each dimension is either:

| Dimension | Meaning |
| --- | --- |
| `Layer @ depth N` | Layer-path segment N levels below the parent (1 = level, 2 = the next segment, …) |
| `UserText key` | A user text value on the object |

Default: `Layer @ 1 → Layer @ 2` (Level → Category — matches the contract out of the box). The old two-key S3 workflow is reproduced with two `UserText` dimensions.

Footprints are merged per leaf group per level, then summed across levels. Results as an indented tree with cumulative areas per node.

### R1 / R2 — Analysis
One shared engine, two tabs so two setups can coexist (e.g. R1 per category, R2 per level). Each tab picks its **Aggregate by** dimension: `Category (layer)`, `Level (layer)`, or a `UserText key`. Scope = the **S3 Parent Layer**.

Merged areas are aggregated across all levels and compared against the **Target Areas** table (Settings) as bullet charts with the ± tolerance zones. Labels without a target chart as "(no target)" and warn. Charts export as PNG.

---

## Settings tab

| Setting | Meaning |
| --- | --- |
| **Ignore Prefix** | Default `_`. Excludes matching layers and their whole subtree (S3, S4, R1, R2). |
| **Global Tolerance (%)** | Symmetric tolerance for the R1/R2 bullet charts. Default 10%. |
| **Decimal Places** | 0–4, applies to all displayed results. |
| **Z-Height Tolerance** | Min. Z gap (model units) to treat objects in one group as separate floors before merging. Applies to all scenarios. |
| **Target Areas** | Label → target value rows for R1/R2. Label must match the aggregation value (category layer name, level name, or key value). Decimal comma accepted. **Fill from last S3** seeds rows from the last S3 result's categories. |

### Configuration (config v2)

Settings **persist automatically per model**: when the window closes (and whenever a config is saved or loaded), the current settings are written into the 3dm as document user text (`Lindero.config_v2`) and restored the next time Lindero opens with that model. Save the 3dm to keep them — they travel inside the model file. A loaded config replaces the persisted one.

**Save Config / Load Config** additionally exchanges the same settings as a JSON file (for sharing between models or machines). v1 config files still load (the S4 key list becomes UserText dimensions; obsolete v1 fields are skipped with a note).

```json
{
  "config_version": 2,
  "tolerance_percent": 10.0,
  "decimal_places": 2,
  "z_height_tol": 0.5,
  "ignore_prefix": "_",
  "s3_parent_layer": "07-PROPOSAL-Phase_03",
  "s4_parent_layer": "07-PROPOSAL-Phase_03",
  "s4_dimensions": [["layer", 1], ["layer", 2]],
  "r1_dimension": ["category", null],
  "r2_dimension": ["level", null],
  "targets": {"Pflege": 1200.0, "Diagnostik und Therapie": 2400.0}
}
```

---

## Bullet chart legend

Each row in R1/R2 is drawn as follows (left to right):

```
[Label]       ░░░░▓▓▓[████████████]░░░▓▓▓░░░░   87.5/100.0
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

The **"Write Area to Objects"** button opens an inline panel below the button row (S1–S4):

1. Choose or type the user text key to write to (default: `Area`).
2. Click **Confirm Write**.

The calculated footprint area is written as a user string to each measured object using `SetUserString`. If the key does not yet exist on an object, it is created. The operation is one-way — no automatic sync; click again to update after recalculating.

Status bar confirms: `Area written to N object(s) using key 'Area'`.

---

## Export to Excel

Available for S1, S2, S3, and S4.

**S1, S2** workbooks: an **Objects** sheet (one row per object) and a **Summary** sheet (parameters, totals, overlap warnings).

**S3** workbooks:

| Sheet | Contents |
| --- | --- |
| **Objects** | Flat, pivot-ready — GUID, parent, **Level**, **Category**, label, area |
| **Summary** | **Level × Category matrix**: levels as rows, categories as columns, level totals on the right, building-wide category totals + grand total at the bottom, warnings highlighted in amber |

**S4** workbooks: **Leaf Data** (one row per unique dimension path — pivot-ready) and **Tree Summary** (indented hierarchy with cumulative areas). Column headers use the dimension labels (e.g. `Layer @ 1`).

---

## Export Chart as PNG

Available when the R1 or R2 tab is active. Renders the full bullet chart to a 900 px wide PNG file at the path you choose. The chart height scales automatically with the number of entries (54 px per row).

---

## Copy Window

The **"Copy Window"** button screenshots the entire Lindero window (including the title bar) straight onto the Windows clipboard — paste it into chat or mail with Ctrl+V. Works from any tab, capturing whatever results are currently visible. The capture rect comes from Win32 in physical pixels, so mixed-DPI multi-monitor setups work. Windows only.

---

## Footprint calculation logic

### Accepted geometry

| Type | How the footprint is extracted |
| --- | --- |
| **Brep / Extrusion** (solid) | Bottom horizontal face(s) are identified; their outer border curves are projected to Z=0 |
| **Closed planar curve** | The curve itself is projected to Z=0 and its enclosed area is used directly |
| **Hatch** | Outer boundary loops are extracted with `Get3dCurves` and projected to Z=0 |
| **Planar surface** (trimmed or untrimmed) | Outer loop extracted; if unavailable, edge curves are joined as a fallback |
| **Other** (no horizontal face found) | Falls back to the **XY bounding box** as an approximation |

### Step-by-step (Brep / Extrusion)

1. Iterate every face of the solid.
2. Evaluate the face normal at its centre point.
3. Keep only faces where `|normal.Z| > 0.9` — the face is within approximately **26° of horizontal**.
4. Among those horizontal faces, identify the one(s) at the **lowest centroid Z** (the actual bottom of the geometry).
5. Extract the **outer border curve** of each bottom face.
6. Project that curve straight down onto the XY plane (`Z = 0`).
7. Compute the area enclosed by the projected curve using `AreaMassProperties`.

### What `|normal.Z| > 0.9` means

The `0.9` is a threshold on the dot product between the face normal and the world Z axis — not a distance in model units. It means the face normal deviates less than ~26° from vertical, i.e. the face is less than ~26° from horizontal. This tolerance handles faces that are nominally flat but carry small modelling imperfections.

### Footprint cache

Within one calculation run every object's projection is computed **once** and reused across the per-object listing and all union calls (`_fp`). On a hospital-scale model (many levels × categories) this cuts the union workload substantially.

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

## Overlap removal (all scenarios)

Projected footprint loops are resolved with a hole-aware planar arrangement (`Curve.CreateBooleanRegions` + point classification), with a pairwise inclusion–exclusion path and a plain `CreateBooleanUnion` as fallbacks — see the `combined_area` docstring for why. If everything fails, the tool falls back to a plain sum and marks the result with `[union failed — sum shown]`.

---

## Refresh Model

Click **Refresh Model** to re-scan all layers and user text keys and update every dropdown and ComboBox without reopening the script. Existing selections are preserved where possible.

---

## Migration notes (v0.7 → v0.8)

- Old **S3 (two keys)** → S4 with two `UserText` dimensions.
- Old **S5** → new S3: sublayer-as-group models work as levels without category layers (objects land in category `—`).
- Old **R1/R2 data source + target keys** → per-tab dimension pickers + the Settings Target Areas table.
- v1 config JSONs load with automatic mapping.

## To-Do

- Silhouette-based footprint for Case B2 (cantilever / overhang solids).
- Highlight out-of-tolerance objects in the Rhino viewport from R1/R2.
- Excel export for R1/R2 (analysis results with target comparison).
- Optional: a "measured but excluded from totals" category class (e.g. Freifläche reported separately) — see PLAN.md §10.
