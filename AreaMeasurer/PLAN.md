# Lindero v2 — Rebuild Plan

Status: **implemented 2026-07-02 (v0.8)** — pending in-Rhino verification against the manual checklist in §8.
Decisions taken with Aksel: consolidate scenarios · generic-depth hierarchy · `_` prefix with cascade · R1/R2 targets from a config table.

---

## 1. Why rebuild

The model structure changed. Objects now live **two levels below** the proposal layer, and the
grouping dimension moved **from user-text keys into the layer tree**:

```text
07-PROPOSAL-Phase_03            ← parent (scope)
├── BuildingVolumes             ← extra stuff → rename _BuildingVolumes to ignore
├── Ebene_Parkhaus              ← Aksel decides: level or _ignored
├── Ebene_06                    ← LEVEL (direct child of parent)
│   ├── TEXT                    ← annotations → _TEXT
│   ├── Freifläche              ← in/out via _ prefix, per level
│   ├── Diagnostik und Therapie ← CATEGORY (DIN 13080) — objects live here
│   ├── Pflege
│   ├── ...
│   └── Verkehrserschließung
├── Ebene_05 … Ebene_-03
├── Ebene_KMA                   ← ordinary level, no special casing
└── Ebene-Logistik Zentrum      ← ordinary level
```

Current code cannot read this:

| Problem | Where |
| --- | --- |
| `get_child_layers()` returns **direct children only** | `Lindero.py:146` |
| `get_layer_objects()` returns objects **directly on** a layer | `Lindero.py:141` |
| → S3/S4/S5/R1/R2 pointed at the proposal find the `Ebene_XX` layers but **zero objects** | `calc_s3:761`, `calc_s4:806`, `calc_s5:863`, `calc_r1:915`, `calc_r2:958` |
| Group dimension assumed to be a user-text key; it is now a layer | `calc_s3` grp_key |
| Scenario redundancy: S5 ≈ S2-per-sublayer; old S3 ⊂ S4; R1 ≈ R2 | — |

## 2. Target data contract

- **Parent** — one proposal layer, chosen in the UI (e.g. `07-PROPOSAL-Phase_03`).
- **Level** — each non-ignored **direct child** of the parent (`Ebene_06`, `Ebene_KMA`, …).
- **Category** — the **short name of the object's own layer** (lowest level). Any depth below the
  level is tolerated; deeper nesting still resolves to the object's own layer name.
- **Objects directly on a level layer** — counted, category `—`, reported as a warning.
- **Ignore rule** — layer whose short name starts with the prefix (default `_`) is excluded
  **together with its entire subtree**. Prefix configurable in Settings, persisted in config.
- Objects are collected from **all** non-ignored descendant layers of each level, not only
  direct grandchildren.

## 3. New scenario set (consolidation)

| Tab | Fate | Behaviour |
| --- | --- | --- |
| **S1 Selection** | keep | unchanged (Z-banding stays) |
| **S2 Single Layer** | keep | unchanged |
| **S3 Layer Hierarchy** | **rebuild** | replaces old S3 **and** S5 — reads Level × Category from the layer tree (contract above) |
| **S4 Custom Aggregation** | generalize | each dimension in the sequence is either `Layer @ depth N` **or** `UserText key`. Old S3 (two keys) and old S4 (key list) are both reproducible as configs of this |
| S5 Group Hierarchy | **remove** | Z-band logic folded into S3 as an advanced option; S2 already bands |
| **R1 / R2 Analysis** | rewire | one shared implementation, two tabs; data source = any S3 dimension (Level or Category) or an S4 level index; targets from the Settings table |

### S3 aggregation & overlap rules

1. Union per **(level, category)** → category totals per level. *(hole-aware
   `CreateBooleanRegions` path in `combined_area:525` stays exactly as is — see its docstring.)*
2. **Level total** = union across all objects of the level → catches cross-category overlap,
   reported per level (same spirit as old S3 cross-group warning).
3. **Grand total** = Σ level totals (floors are additive — GFA logic, unchanged).
4. **Building-wide category totals** = Σ of that category's per-level unions.
5. `z_height_tol` kept as an advanced field (split-level safety, inherited from S5).

## 4. Engine changes

| Item | Detail |
| --- | --- |
| `AreaMeasurer/_paths.py` (new) | **Pure Python, no Rhino imports**: `is_ignored(short, prefix)`, `classify(parent, full_path, prefix) → (level, category) or None` with cascade ignore, tree assembly from a flat path list. Headless-testable like `TerrainTools/_core` |
| `collect_objects(parent, prefix)` | returns `[{guid, level, category, layer}]` using `_paths.classify` |
| **Footprint cache** | `get_footprint_curves` is currently recomputed per object listing **and** inside every `combined_area` call (`:547`). With levels × categories the union count multiplies → cache `{guid: curves}` per calculation run |
| `combined_area(guids, cache=None)` | signature extended, logic untouched |
| Dry-run preview | before calculating, list included levels/categories with object counts and the ignored layers, so surprises surface before a long union run |

## 5. UI changes

- Tabs: S1, S2, S3, S4, R1, R2, Settings (7 instead of 8).
- **S3 tab**: parent dropdown · Preview button + panel (included/ignored tree, object counts) ·
  results: per-level category grid, building-wide category summary, per-object grid.
- **Settings**: ignore prefix field (default `_`) · **target table** (category → target area,
  editable grid, `StackLayout` for dynamic rows per the Eto rules in `CLAUDE.md`) · existing
  decimals/tolerance · config save/load.
- Config JSON gains `"config_version": 2`, `"ignore_prefix": "_"`, `"targets": {"Pflege": 1200.0, …}`;
  loader accepts v1 files and maps them.

## 6. Excel export

- S3 **Objects** sheet gains `Level` and `Category` columns (pivot-ready flat table).
- S3 **Summary** becomes a Level × Category matrix plus totals row/column and overlap warnings.
- R1/R2 export stays on the to-do list (unchanged scope).

## 7. Migration notes

- Old S3 two-key workflow → S4 with two `UserText` dimensions.
- Old S5 → new S3 with no category layers present (objects directly on level layers, category `—`).
- README Scenarios section rewritten; `install.py` remains deprecated and untouched.

## 8. Test plan

- **Headless**: `AreaMeasurer/tests/test_paths.py` for `_paths.py` — cascade ignore, prefix edge
  cases (`_` exactly, prefix mid-name, nested `_TEXT` under kept level), classification at depths
  1/2/3, objects on level layers.
- **In Rhino, manual checklist** (small test 3dm with known areas):
  - two levels × two categories, overlap *within* one category → union removes it;
  - overlap *across* categories on one level → level total < Σ category totals + warning;
  - object directly on `Ebene_XX` → category `—` + warning;
  - `_TEXT` hatch present → excluded; toggle `Freifläche` → totals change accordingly;
  - empty category layer → 0.0 row, no crash; config v1 load → mapped without error.

## 9. Order of work (each phase leaves the tool runnable)

1. `_paths.py` + headless tests.
2. Engine: `collect_objects`, footprint cache, new `calc_s3`.
3. UI: rebuild S3 tab + preview panel, remove S5 tab.
4. S4 dimension picker (`Layer @ depth` | `UserText`).
5. R1/R2 rewire + Settings target table + config v2.
6. Excel export update, README rewrite.

## 10. Explicitly out of scope for v2

- Silhouette footprint for cantilevers (Case B2 — README known limitation).
- Viewport highlighting from R1/R2.
- A "measured but excluded from DIN totals" category class (e.g. Freifläche counted separately).
  For now the `_` prefix is the single in/out switch; revisit if needed.
