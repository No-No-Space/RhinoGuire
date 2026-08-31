# RhinoGuire

A collection of high-performance Python 3 tools for Rhino 8 for managing BIM metadata, architectural footprint calculations, live-field tagging, mesh projection, road draping, and terrain grading (cut & fill earthwork).

---

## 🧭 Tool Suite Overview

| Tool | Category | Input Geometry / Data | Output | Extra Dependencies | Launcher |
| :--- | :--- | :--- | :--- | :--- | :--- |
| [**Lindero**](AreaMeasurer/) | Area & Space Planning | Solids, Extrusions, Curves, Surfaces, Hatches | Footprint plan area, matrix breakdown, Excel & PNG | None | `launch_lindero.py` |
| [**Arriero**](DataExporterImporter/) | Metadata Management | Rhino Objects with User Text | Two-way Excel sync (`.xlsx`) | `openpyxl` | `launch_arriero.py` |
| [**Chivito**](DataVisualization/) | Data Visualization | User Text Metadata & Excel Color Map | Color-coded Rhino viewports & legend export | `openpyxl` | `launch_chivito.py` |
| [**Baquiano**](SearchData/) | Search & Filtering | Document or pre-selected objects | Viewport selection via 8 condition match types | None | `launch_baquiano.py` |
| [**Pregonero**](Tagging/) | Annotation & BIM Tagging | Rhino Objects with User Text | Dynamic leaders with live `%<UserText(...)>%` fields | None | `launch_pregonero.py` |
| [**Sebucan**](MeshTools/WrapeMeshOnMesh/) | Mesh Operations | Source Meshes + Destination Terrain / Surface | Wrapped mesh with adaptive Z-refinement | None | `launch_sebucan.py` |
| [**PadGrader**](TerrainTools/PadGrader/) | Terrain Grading | Closed Planar Curves + Terrain | Graded terrain mesh + cut/fill quantities | None | `launch_padgrader.py` |
| [**WayGrader**](TerrainTools/WayGrader/) | Terrain Grading | Centerline Polyline + Terrain | Graded corridor mesh + mass-haul station data | None | `launch_waygrader.py` |
| [**CutFillReport**](TerrainTools/CutFillReport/) | Earthwork Reporting | Two Terrains or grading result | KPI dashboard, depth tinted map, Excel & PNG | `openpyxl` | `launch_cutfillreport.py` |
| [**Trocha**](RoadTool/) | Road Infrastructure | Centerline Curve + Terrain | Continuous solid road slab (non-destructive) | None | `launch_trocha.py` |

---

## 🛠️ Tool Details

### [Lindero](AreaMeasurer/) — Footprint Area Calculator
Calculates the **footprint area** of Rhino objects — the true plan area as seen from directly above (XY projection), distinct from Rhino's built-in `Area` command which sums all 3D faces.
- **S1 (Selected Objects) & S2 (By Layer):** Merges overlapping footprints per layer/Z-band to avoid double-counting.
- **S3 (Layer Hierarchy):** Traverses the layer tree (`Parent → Level → Category`) to produce Gross Floor Area (GFA) aggregations and cross-category overlap warnings.
- **S4 (Custom Aggregation):** Flexible N-level hierarchy built from layer-path depths and User Text keys.
- **R1 / R2 (Target Analysis):** Evaluates measured areas against target schedules as bullet charts with customizable tolerance bands.
- **Exports:** Excel workbooks (pivot-ready sheets and Level × Category matrices) and bullet charts as PNG.

### [Arriero](DataExporterImporter/) — Metadata Exporter & Importer
Two-way synchronization between Rhino User Text and Excel spreadsheets using GUID tracking.
- **Export:** Dumps all object GUIDs and key/value pairs into clean, structured spreadsheets.
- **Import:** Updates existing metadata, automatically adds missing keys from new columns, supports customizable empty-cell handling (preserve, set placeholder, or delete key), and generates automatic timestamped backups before updating.

### [Chivito](DataVisualization/) — Metadata Color Visualizer
Color-codes Rhino objects based on metadata values through an interactive, modeless Color Manager.
- **3-Step Workflow:** (1) Initialize keys from template Excel, (2) Scan and export unique values to a color mapping spreadsheet, (3) Apply interactive colors in the viewport.
- **Diagnostics & Output:** Identifies unmapped or missing values with a "Select Problem Objects" diagnostic tool; exports standalone legends and viewports to PNG.

### [Baquiano](SearchData/) — Search & Select Objects by Metadata
Search and isolate Rhino objects across your model using boolean include/exclude query rules.
- **Flexible Matching:** 8 match types (`Contains`, `Equals`, `Starts with`, `Ends with`, and their negations).
- **Scope:** Search across the whole model or restrict queries to pre-selected candidate objects.

### [Pregonero](Tagging/) — Live-Field Object Tagger
Creates annotation leaders driven by reusable text templates, similar to BIM "tag by category" workflows.
- **Live Text Fields:** Placeholders (`{KeyName}`) are inserted as live Rhino text fields (`%<UserText("guid","Key")>%`), automatically staying up-to-date if object metadata changes.
- **Missing Key Safety:** Keys missing on tagged objects are initialized with `TBD` to prevent broken field markers (`####`).

### [Sebucan](MeshTools/WrapeMeshOnMesh/) — Wrap Mesh on Mesh
Projects source meshes vertically onto any target geometry along the Z axis (Mesh, SubD, Surface, Polysurface, or Solid).
- **Adaptive Refinement:** Automatically subdivides coarse faces only where terrain curvature exceeds a user-defined tolerance.

### [TerrainTools](TerrainTools/) — Terrain Grading & Earthwork Suite
A unified suite sharing a fast heightfield grading engine (`_core`) for **modifying site terrain** without altering the original geometry:
- **[PadGrader](TerrainTools/PadGrader/):** Grades building pads (closed boundary curves at target elevations) out to the terrain daylight line along custom cut/fill slopes.
- **[WayGrader](TerrainTools/WayGrader/):** Grades road/path corridors from centerlines with crown or single crossfall, daylight skirts, and station-by-station mass-haul tracking.
- **[CutFillReport](TerrainTools/CutFillReport/):** Compares original vs. modified surfaces, computes cut/fill net balance, renders a tinted depth mesh, and exports KPI reports to Excel and PNG.
- *Documentation:* See [`PLAN.md`](TerrainTools/PLAN.md) and [`DECISIONS.md`](TerrainTools/DECISIONS.md).

### [Trocha](RoadTool/) — Solid Road Slab on Terrain
Drapes a clean, crease-free solid road slab onto terrain from a 3D centerline **without modifying the terrain** (non-destructive presentation and documentation tool).
- **Contact Rule:** Sinks the slab underside into the ground so roads never "fly" over dips, while keeping the top face smooth.
- **Junction Merge:** Boolean-unions intersecting road slabs on demand while keeping source centerlines editable.
- *Documentation:* See [`README.md`](RoadTool/README.md) and the complete [Design Spec (`road_tool_plan.md`)](RoadTool/road_tool_plan.md).

---

## 💻 Requirements & Dependencies

- **Platform:** Rhino 8 for Windows (CPython 3 runtime).
- **Dependencies:**
  - `openpyxl` — Required only for Excel exports in **Arriero**, **Chivito**, and **CutFillReport**. Rhino 8 installs this automatically via the `# r: openpyxl` script header on first run.
  - All other tools (**Lindero**, **Baquiano**, **Pregonero**, **Sebucan**, **PadGrader**, **WayGrader**, **Trocha**) use pure RhinoCommon and standard Python with **zero external dependencies**.

---

## 🚀 Quick Start

### Method 1: Run via Command Line / Script Editor
1. In Rhino 8, run `_-RunPythonScript`.
2. Browse to any launcher shim in the `RhinoGuire/` root:
   - `launch_lindero.py`
   - `launch_arriero.py`
   - `launch_chivito.py`
   - `launch_baquiano.py`
   - `launch_pregonero.py`
   - `launch_sebucan.py`
   - `launch_padgrader.py`
   - `launch_waygrader.py`
   - `launch_cutfillreport.py`
   - `launch_trocha.py`

### Method 2: Load the Rhino Toolbar
Load `ui/RhinoGuire.rui` into Rhino 8 for one-click toolbar access. See the [Toolbar Setup Guide](ui/README.md) for full configuration steps.

---

## 📂 Repository Directory Layout

```text
RhinoGuire/
├── AreaMeasurer/               ← Lindero (Footprint Area Calculator & GFA Matrix)
│   ├── Lindero.py              ← Modeless Eto application
│   ├── _paths.py               ← Headless layer-hierarchy path resolution engine
│   ├── PLAN.md                 ← Architecture & design specification
│   └── tests/                  ← Unit tests (test_paths.py)
├── DataExporterImporter/       ← Arriero (Rhino ↔ Excel metadata synchronization)
├── DataVisualization/          ← Chivito (Metadata color-coding & legend visualizer)
├── MeshTools/WrapeMeshOnMesh/  ← Sebucan (Z-projection & adaptive mesh refinement)
├── RoadTool/                   ← Trocha (Solid road draping on terrain)
│   ├── Trocha.py               ← Modeless Eto application
│   ├── _core/                  ← Geometry engine, configuration, state & junctions
│   ├── road_tool_plan.md       ← Full technical design specification
│   └── tests/                  ← Headless unit tests (test_headless.py)
├── SearchData/                 ← Baquiano (Multi-condition metadata query & selection)
├── Tagging/                    ← Pregonero (Live-field BIM leader annotations)
├── TerrainTools/               ← Terrain grading suite (PadGrader, WayGrader, CutFillReport)
│   ├── _core/                  ← Shared grading engine, terrain raycaster & volumes
│   ├── _widgets.py             ← Reusable slope and UI controls
│   ├── DECISIONS.md            ← Architecture decision record
│   ├── PLAN.md                 ← Implementation roadmap
│   └── tests/                  ← Headless unit tests (test_headless.py)
├── ui/                         ← Theme definitions, color palettes & toolbar RUI
│   ├── theme.py                ← Central Eto styling design system
│   └── RhinoGuire.rui          ← Toolbar definition
├── launch.py                   ← Central script dispatcher
├── launch_*.py                 ← Dedicated root shims for each tool
├── manifest.yml                ← Yak package manager metadata
├── .gitignore / .yakignore     ← Git and Yak package filter definitions
└── README.md                   ← Main repository documentation
```

---

## 🧪 Running Headless Tests

All headless geometry, slope conversion, volume prism, and layer path logic can be tested without Rhino running:

```bash
# Test Lindero layer-hierarchy logic
python AreaMeasurer/tests/test_paths.py

# Test TerrainTools slope & volume engine
python TerrainTools/_core/tests/test_headless.py

# Test Trocha road configuration & parameter derivation
python RoadTool/_core/tests/test_headless.py
```

---

## 📄 License

MIT License — see [LICENSE](LICENSE) for details.

## 👤 Author

**Aksel Alvarez** — [aquelon@pm.me](mailto:aquelon@pm.me)

