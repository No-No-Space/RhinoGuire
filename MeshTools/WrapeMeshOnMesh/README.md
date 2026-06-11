# Sebucan — Wrap Mesh on Mesh

Projects one or more source meshes onto a destination surface along the Z axis. Every source vertex keeps its X/Y position and its Z is snapped to the destination. An optional adaptive refinement pass splits coarse faces where terrain detail between vertices would otherwise be lost.

**Typical use case:** road or path meshes that need to follow the contours of a terrain mesh, SubD landscape, or solid polysurface.

## Workflow

### Step 1 — Select Destination

Click **Select Destination** and pick the geometry to project onto.

Accepted types: **Mesh** (used directly), **SubD** (converted to BRep then meshed), **Surface / Polysurface / Solid** (meshed internally via `Mesh.CreateFromBrep`). The destination type is confirmed in the panel.

### Step 2 — Select Source Mesh(es)

Click **Select Source Mesh(es)** and pick one or more meshes. Multiple meshes can be selected at once (e.g. a full set of road meshes).

### Step 3 — Options

| Option | Default | Description |
| --- | --- | --- |
| Replace source mesh(es) | off | Deletes the originals after wrapping |
| Adaptive refinement | off | Splits faces where terrain Z deviation exceeds tolerance |
| Tolerance | `0.1` | Max Z error at any edge midpoint before a face is split |
| Max. passes | `3` | Max refinement iterations per mesh |

### Step 4 — Wrap!

Status bar reports vertices projected, vertices added by refinement, and vertices outside the destination footprint (original Z kept). New meshes are added to the same layer as their source.

## Adaptive Refinement

For each triangle face, the midpoint of each edge is projected onto the destination to get the true terrain Z. If the deviation from the face's interpolated Z exceeds the tolerance on any edge, the face is split into 4 sub-triangles. Repeated up to Max. passes. Each unique (x, y) position is raycasted only once (cached), so shared edge midpoints cost nothing on repeated checks.

**Tolerance guidelines:**

| Model unit | Suggested tolerance |
| --- | --- |
| Metres | `0.05` – `0.20` |
| Centimetres | `5` – `20` |
| Millimetres | `50` – `200` |

## Limitations

- Source objects must be **meshes** (result is always a mesh).
- Projection is strictly **vertical (Z axis)** — not suitable for predominantly horizontal surfaces.
- Vertices outside the XY footprint of the destination keep their original Z.
- Non-manifold or degenerate destination meshes may produce incorrect results — run `MeshRepair` first.
