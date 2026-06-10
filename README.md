# RhinoGuire

A collection of Python 3 tools for managing object metadata and geometry in Rhino 8.

## Tools

### [Arriero](DataExporterImporter/) — Data Exporter/Importer

Export and import object metadata between Rhino and Excel files using GUID-based tracking. Supports backup creation, key creation scope, and flexible handling of empty cells.

### [Chivito](DataVisualization/) — Data Visualization

Color-code Rhino objects based on their metadata values. Three-step workflow: initialize keys from Excel, extract unique values, and visualize with an interactive Color Manager. Includes legend and viewport PNG export.

### [Baquiano](SearchData/) — Search Data

Search and select Rhino objects by their metadata using include/exclude conditions with 8 match types (Contains, Equals, Starts with, Ends with, and their negations). Supports pre-selection filtering and cross-search queries.

### [Lindero](AreaMeasurer/) — Footprint Area Calculator

Calculates the **footprint area** of Rhino objects — the plan area as seen from directly above (XY projection), distinct from Rhino's built-in `Area` command which sums all faces.

Runs as a modeless window with six calculation scenarios:

- **S1 — Selected Objects:** individual footprint per object; overlapping footprints merged with Boolean Union to avoid double-counting.
- **S2 — By Layer:** footprints of all objects on a layer with Boolean Union to remove overlapping regions.
- **S3 — Layer Hierarchy:** targets a parent layer whose sublayers represent floors. Overlaps removed per floor; results split into a Groups panel and an Objects panel. Grand total follows standard Gross Floor Area logic.
- **S4 — Custom Aggregation:** user-defined key hierarchy (e.g. Domain → Main Group → Subgroup → Room Type). Footprints merged per leaf group per floor (same as S3), summed across all floors. Results displayed as a colored grid; exports to Excel.
- **R1 — Room Analysis:** aggregates areas by a chosen key across all floors and compares totals to target values. Displayed as bullet charts with tolerance bands. Data source: S3 keys or S4 hierarchy.
- **R2 — Group Analysis:** same as R1 but at a group/classification level.

Accepted geometry: solids, extrusions, closed planar curves, planar surfaces, and hatches. Supports labelling via user text keys, configurable decimal places, Write Area to Objects, Excel export (S1–S4), and PNG chart export (R1–R2).

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

## Requirements

- **Rhino 8** with CPython 3
- **openpyxl** — required by Arriero, Chivito and CutFillReport (installed automatically via `# r: openpyxl` header)
- Baquiano, Lindero, Sebucan, PadGrader and WayGrader have no external dependencies

## Quick Start

1. Open Rhino 8
2. Type `RunPythonScript` in the command line
3. Navigate to the desired tool's `.py` file and click **Open**

Each tool opens its own GUI window. See the individual README files for detailed usage instructions.

Alternatively, load the toolbar bundle (`ui/RhinoGuire.rui`) for one-click access from the Rhino interface — see [`ui/README.md`](ui/README.md) for setup instructions.

## Repository Structure

```text
RhinoGuire/
├── README.md
├── LICENSE
├── install.py
├── launch.py
├── manifest.yml
├── ui/
│   └── README.md
├── AreaMeasurer/
│   ├── Lindero.py
│   └── README.md
├── DataExporterImporter/
│   ├── Arriero.py
│   └── README.md
├── DataVisualization/
│   ├── Chivito.py
│   ├── README.md
│   └── _ExcelOutput/
├── MeshTools/
│   └── WrapeMeshOnMesh/
│       ├── Sebucan.py
│       └── README.md
├── SearchData/
│   ├── Baquiano.py
│   └── README.md
└── TerrainTools/
    ├── README.md            # suite overview
    ├── PLAN.md              # implementation plan
    ├── DECISIONS.md         # decision log
    ├── _core/               # shared grading engine (grading_core)
    │   ├── slope.py
    │   ├── terrain.py
    │   ├── grading.py
    │   ├── volumes.py
    │   ├── meshbuild.py
    │   ├── report.py
    │   └── tests/test_headless.py
    ├── _widgets.py          # shared Eto widgets + doc helpers
    ├── PadGrader/
    │   ├── PadGrader.py
    │   └── README.md
    ├── WayGrader/
    │   ├── WayGrader.py
    │   └── README.md
    └── CutFillReport/
        ├── CutFillReport.py
        └── README.md
```

## License

MIT License — see [LICENSE](LICENSE) for details.

## Author

Aquelon — aquelon@pm.me
