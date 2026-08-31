# RhinoGuire Toolbar (`ui/`)

## Files
- `RhinoGuire.rui` — Rhino toolbar file (customized toolbar definition). Note: `.rui` files contain local paths; see instructions below to build or load your toolbar.
- `theme.py` — Shared UI design system, colors, fonts, and styling helpers for all Eto.Forms windows.
- `InfoAboutIcons.txt` — Icon reference sources and guidance.

---

## Creating the Toolbar

### 1. Load or Create Toolbar in Rhino 8
In Rhino 8: `Tools > Toolbar Layout... > New...`
- Name: `RhinoGuire`

### 2. Add Buttons
Right-click the newly created toolbar > **New Button** for each tool.

**Recommended button macro (using direct root launcher shims):**
```text
! _-RunPythonScript "<path-to-repo>/RhinoGuire/launch_<toolname>.py"
```

*Example for Lindero:*
```text
! _-RunPythonScript "D:/404-Github/008-RhinoGuire/RhinoGuire/launch_lindero.py"
```

**Alternative button macro (via central `launch.py` dispatcher):**
```text
! _-RunPythonScript "<path-to-repo>/RhinoGuire/launch.py" "<ToolName>"
```

### 3. Available Tools & Button Mappings

| Label | Script Key / Dispatcher | Dedicated Launcher Shim | Description |
| :--- | :--- | :--- | :--- |
| **Lindero** | `Lindero` | `launch_lindero.py` | Area & Footprint Calculator (XY Plan Projection) |
| **Arriero** | `Arriero` | `launch_arriero.py` | Metadata Exporter / Importer (Rhino ↔ Excel) |
| **Chivito** | `Chivito` | `launch_chivito.py` | Metadata Color-Coder & Visualizer |
| **Baquiano** | `Baquiano` | `launch_baquiano.py` | Metadata Query & Object Selector |
| **Pregonero** | `Pregonero` | `launch_pregonero.py` | Live-Field Object Tagger & Leader Generator |
| **Sebucan** | `Sebucan` | `launch_sebucan.py` | Wrap Mesh on Mesh (Z-Projection Engine) |
| **PadGrader** | `PadGrader` | `launch_padgrader.py` | Terrain Grading — Building Pads |
| **WayGrader** | `WayGrader` | `launch_waygrader.py` | Terrain Grading — Way / Path Corridors |
| **CutFillReport** | `CutFillReport` | `launch_cutfillreport.py` | Cut & Fill Volume Quantification & Excel Export |
| **Trocha** | `Trocha` | `launch_trocha.py` | Solid Road Slab Draped on Terrain |

> **Dependencies:** `Arriero`, `Chivito`, and `CutFillReport` require `openpyxl` for Excel export. Rhino 8 installs it automatically on first run via the `# r: openpyxl` header.

### 4. Save the Toolbar
In Rhino: `File > Save As...` → save as `RhinoGuire/ui/RhinoGuire.rui`.

---

## Quick Setup for Team Members

1. Clone or copy the `008-RhinoGuire/RhinoGuire` repository folder.
2. In Rhino 8, run tools directly with `RunPythonScript` on any `launch_<name>.py` file.
3. To load the toolbar: `Tools > Toolbar Layout... > File > Open...` → Select `ui/RhinoGuire.rui`.

