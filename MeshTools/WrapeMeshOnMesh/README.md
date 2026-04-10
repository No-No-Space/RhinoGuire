# Sebucan — Wrap Mesh on Mesh

**Version 0.3 · Rhino 8 · CPython 3**

Sebucan projects one or more source meshes onto a destination surface along the Z axis. Every source vertex keeps its X/Y position and its Z is snapped to the destination geometry. An optional adaptive refinement pass splits coarse faces where the terrain detail between vertices would otherwise be lost.

**Typical use case:** road or path meshes that need to follow the contours of a terrain mesh, SubD landscape, or solid polysurface.

---

## Requirements

| Requirement | Notes |
|---|---|
| Rhino 8 | CPython 3 engine (`#! python3` header) |
| `rhinoscriptsyntax`, `Eto.Forms` | Bundled with Rhino 8 — no install needed |

---

## How to run

```
RunPythonScript → Sebucan.py
```

The Sebucan panel opens as a modeless window. Rhino stays fully accessible for viewport selection throughout.

---

## Workflow

### Step 1 — Select Destination
Click **Select Destination** and pick the geometry to project onto.

Accepted types:
- **Mesh** — used directly
- **SubD** — converted to BRep, then meshed internally
- **Surface / Polysurface / Solid** — meshed internally via `Mesh.CreateFromBrep`

The destination type is shown in the panel after selection (e.g. `✓ Terrain  [Mesh]`).

### Step 2 — Select Source Mesh(es)
Click **Select Source Mesh(es)** and pick one or more meshes to wrap. Multiple meshes can be selected at once (e.g. a full set of road meshes).

### Step 3 — Configure options (optional)

| Option | Default | Description |
|---|---|---|
| Replace source mesh(es) | off | Deletes the originals after wrapping. When off, new meshes are added alongside the originals. |
| Adaptive refinement | off | Splits faces where the terrain Z deviation exceeds the tolerance (see below). |
| Tolerance | `0.1` | Maximum allowed Z error at any edge midpoint before a face is split. Match your model's unit scale. |
| Max. passes | `3` | Maximum number of refinement iterations per mesh. |

### Step 4 — Wrap!
Click **Wrap!**. The status bar reports how many vertices were projected, how many vertices were added by refinement, and how many fell outside the destination (their original Z is kept).

New meshes are added to the same layer as their source.

---

## Adaptive Refinement

Without refinement, a coarse source mesh (few vertices) produces projected triangles that cut straight through terrain detail between vertices. Adaptive refinement solves this without uniformly subdividing the whole mesh:

```
For each triangle face:
  1. Compute the midpoint of each edge
  2. Project that midpoint onto the destination → true terrain Z
  3. Compare to the face's interpolated Z at that point
  4. If deviation > tolerance on any edge → split face into 4 sub-triangles
  5. Repeat up to Max. passes, or until no face fails

Flat areas of terrain → 0 splits → no added geometry
Steep / curved areas  → splits only where needed
```

Each unique (x, y) position is only raycasted once per run (cached), so edge midpoints shared between adjacent faces are free on repeated checks.

**Tolerance guidelines:**

| Model unit | Suggested tolerance |
|---|---|
| Metres | `0.05` – `0.20` |
| Centimetres | `5` – `20` |
| Millimetres | `50` – `200` |

**Passes guideline:** 3 passes is sufficient for most terrain. Each pass can at most quadruple the face count of a failing face, but in practice only a fraction of faces are split per pass.

---

## Limitations

- Source objects must be **meshes**. The source type filter is intentionally restricted to meshes since the result is always a mesh.
- Projection is strictly **vertical (Z axis)**. It is not suitable for surfaces where the normal direction is predominantly horizontal.
- Vertices that fall **outside the XY footprint** of the destination geometry keep their original Z. These are reported in the status bar.
- Non-manifold or degenerate destination meshes may produce incorrect intersection results. Clean the destination mesh before use (`MeshRepair` in Rhino).

---

## Changelog

| Version | Date | Change |
|---|---|---|
| 0.3 | 2026-03-02 | Adaptive refinement: splits faces where Z deviation exceeds tolerance |
| 0.2 | 2026-03-02 | Destination now accepts Mesh, SubD, Surface, Polysurface |
| 0.1 | 2026-03-02 | Initial release |
