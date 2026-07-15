# RhinoGuire

A collection of Python 3 tools for managing object metadata and geometry in Rhino 8.

## Tools

### [Arriero](DataExporterImporter/) — Data Exporter/Importer

Export and import object metadata between Rhino and Excel files using GUID-based tracking. Supports backup creation, key creation scope, and flexible handling of empty cells.

### [Chivito](DataVisualization/) — Data Visualization

Color-code Rhino objects based on their metadata values. Three-step workflow: initialize keys from Excel, extract unique values, and visualize with an interactive Color Manager. Includes legend and viewport PNG export.

### [Baquiano](SearchData/) — Search Data

Search and select Rhino objects by their metadata using include/exclude conditions with 8 match types (Contains, Equals, Starts with, Ends with, and their negations). Supports pre-selection filtering and cross-search queries.

### [Pregonero](Tagging/) — Object Tagger

Tag objects with leaders built from a reusable text template, like Revit's *tag by category* tool. Template placeholders (`{KeyName}`) are inserted as **live Rhino text fields** (`%<UserText("guid","Key")>%`), so each leader reflects the tagged object's own user text and updates when it changes. Pick a template object to load its keys, write the template, choose a Dimension Style and optional text height, then click objects to tag (first click sets the arrow, second sets the text). Keys an object is missing are created with the value `TBD`. Modeless window; one Undo per session.

### [Lindero](AreaMeasurer/) — Footprint Area Calculator

Calculates the **footprint area** of Rhino objects — the plan area as seen from directly above (XY projection), distinct from Rhino's built-in `Area` command which sums all faces.

Runs as a modeless window with six calculation tabs:

- **S1 — Selected Objects:** individual footprint per object; overlapping footprints merged to avoid double-counting.
- **S2 — By Layer:** footprints of all objects on a layer, overlaps merged.
- **S3 — Layer Hierarchy:** reads the layer tree — parent → direct children = levels (floors) → the object's own layer = category (e.g. DIN 13080 areas). Layers prefixed `_` are ignored with their whole subtree; a dry-run preview shows what counts before calculating. Union per (level, category), cross-category overlap warned per level, grand total per Gross Floor Area logic.
- **S4 — Custom Aggregation:** user-defined hierarchy of dimensions, each either a layer-path depth or a user text key (e.g. Layer @ 1 → Layer @ 2, or Domain → Room Type). Footprints merged per leaf group per level, summed across levels.
- **R1 / R2 — Analysis:** aggregate merged areas by category, level, or a user text key across all floors and compare against a target table (Settings). Displayed as bullet charts with tolerance bands.

Accepted geometry: solids, extrusions, closed planar curves, planar surfaces, and hatches. Supports labelling via user text keys, configurable decimal places, Write Area to Objects, Excel export (S1–S4, S3 as a Level × Category matrix), and PNG chart export (R1–R2).

### [Sebucan](MeshTools/WrapeMeshOnMesh/) — Wrap Mesh on Mesh

Projects one or more source meshes onto a destination surface along the Z axis. Every source vertex keeps its X/Y position and its Z is snapped to the destination geometry.

Accepted destination types: Mesh, SubD, Surface, Polysurface, Solid. Includes an **adaptive refinement** pass that splits coarse faces only where terrain Z deviation between vertices exceeds a configurable tolerance — flat areas produce no extra geometry.

Typical use case: road or path meshes that need to follow the contours of a terrain mesh or landscape surface.

### [TerrainTools](TerrainTools/) — Terrain Grading Suite

A suite of three tools sharing one grading engine (`_core`) for **modifying terrains** modelled as Surfaces or Meshes. The terrain is sampled with the same Z-projection technique as Sebucan; grading is computed analytically on a regular heightfield (cut/fill slopes auto-stop at the daylight line). All outputs are new meshes — the original terrain is never modified.

- **PadGrader** — place one or more building pads (closed boundaries at a target elevation) and grade cut/fill slopes around them to daylight. Outputs a graded mesh + cut/fill totals.
- **WayGrader** — grade a way/path corridor from its centerline in a persistent window: width, crossfall (crown/single), cut/fill slopes; Regenerate without re-picking. Outputs a graded corridor mesh + per-station mass-haul.
- **CutFillReport** — compare original vs modified terrain (or read the last grading), compute cut & fill volumes, show KPIs/charts, tint a cut/fill map mesh with a legend, and export to **Excel** and **PNG**.

Slopes accept H:V ratio, percent, or degrees. See [`TerrainTools/`](TerrainTools/) for the design docs (`README.md`, `PLAN.md`, `DECISIONS.md`).

### [Trocha](RoadTool/) — Road on Terrain

Drapes a solid road slab onto a terrain from a user-drawn centerline, **without modifying the terrain** — distinct from TerrainTools, which grades/modifies the terrain itself. A persistent window: pick terrain + centerline, set width/thickness, Generate. The top face stays a single clean surface; the slab is buried deep enough into the terrain that it never lifts off over a dip. Create → update is linked to the centerline (re-running replaces, not duplicates), with an Update All sweep and a Merge step that Boolean-unions built roads at junctions.

See [`RoadTool/`](RoadTool/) for the workflow (`README.md`) and full design spec (`road_tool_plan.md`).

## Requirements

- **Rhino 8** with CPython 3
- **openpyxl** — required by Arriero, Chivito and CutFillReport (installed automatically via `# r: openpyxl` header)
- Baquiano, Pregonero, Lindero, Sebucan, PadGrader, WayGrader and Trocha have no external dependencies

## Quick Start

1. Open Rhino 8
2. Type `RunPythonScript` in the command line
3. Navigate to the desired tool's `.py` file and click **Open**

Each tool opens its own GUI window. See the individual README files for detailed usage instructions.

Alternatively, load the toolbar bundle (`ui/RhinoGuire.rui`) for one-click access from the Rhino interface — see [`ui/README.md`](ui/README.md) for setup instructions.

## License

MIT License — see [LICENSE](LICENSE) for details.

## Author

Aquelon — aquelon@pm.me
