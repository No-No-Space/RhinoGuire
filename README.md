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

Runs as a modeless window. Three calculation scenarios:

- **S1 — Selected Objects:** individual footprint per object, no overlap handling.
- **S2 — By Layer:** footprints of all objects on a layer with Boolean Union to remove overlapping regions.
- **S3 — Layer Hierarchy:** targets a parent layer whose sublayers represent floors. Overlaps removed per floor; grand total follows standard Gross Floor Area logic.

Supports labelling via user text keys and optional Excel export (planned).

### [Sebucan](MeshTools/WrapeMeshOnMesh/) — Wrap Mesh on Mesh

Projects one or more source meshes onto a destination surface along the Z axis. Every source vertex keeps its X/Y position and its Z is snapped to the destination geometry.

Accepted destination types: Mesh, SubD, Surface, Polysurface, Solid. Includes an **adaptive refinement** pass that splits coarse faces only where terrain Z deviation between vertices exceeds a configurable tolerance — flat areas produce no extra geometry.

Typical use case: road or path meshes that need to follow the contours of a terrain mesh or landscape surface.

## Requirements

- **Rhino 8** with CPython 3
- **openpyxl** — required by Arriero and Chivito (installed automatically via `# r: openpyxl` header)
- Baquiano, Lindero and Sebucan have no external dependencies

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
└── SearchData/
    ├── Baquiano.py
    └── README.md
```

## License

MIT License — see [LICENSE](LICENSE) for details.

## Author

Aquelon — aquelon@pm.me
