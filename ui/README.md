# RhinoGuire Toolbar (ui/)

## Files
- `RhinoGuire.rui` — Rhino toolbar (generated from Rhino, see instructions below)

---

## Creating the toolbar for the first time

### 1. Install the scripts
Open Rhino and run in the Python editor (adjust to where you cloned the repo):
```
_RunPythonScript "<path-to-repo>/RhinoGuire/install.py"
```

### 2. Create the toolbar
In Rhino: `Tools > Toolbar Layout > New`
- Name: `RhinoGuire`

### 3. Add a button per script
Right-click on the toolbar > `New Button` for each tool.

**Left button macro (run):**
```
! _RunPythonScript "<path-to-repo>/RhinoGuire/launch.py" "RG_Lindero"
```

**Available buttons:**

| Label         | Script Key        | Description                 |
|---------------|-------------------|-----------------------------|
| Lindero       | RG_Lindero        | Area Measurer               |
| Arriero       | RG_Arriero        | Data Exporter/Importer      |
| Chivito       | RG_Chivito        | Data Visualization          |
| Sebucan       | RG_Sebucan        | Wrap Mesh on Mesh           |
| Baquiano      | RG_Baquiano       | Search Data                 |
| PadGrader     | RG_PadGrader      | TerrainTools — building pad grading   |
| WayGrader     | RG_WayGrader      | TerrainTools — way/path corridor grading |
| CutFillReport | RG_CutFillReport  | TerrainTools — cut/fill compare & export |

> **TerrainTools buttons.** The three TerrainTools above can also be launched
> with the dedicated root shims (mirrors the other tools), which avoids passing
> a script key:
> ```
> ! _-RunPythonScript "<path-to-repo>/RhinoGuire/launch_padgrader.py"
> ! _-RunPythonScript "<path-to-repo>/RhinoGuire/launch_waygrader.py"
> ! _-RunPythonScript "<path-to-repo>/RhinoGuire/launch_cutfillreport.py"
> ```
> CutFillReport needs `openpyxl` for Excel export — Rhino installs it
> automatically from the `# r: openpyxl` header on first run.

### 4. Save the toolbar
`File > Save As` → save as `RhinoGuire/ui/RhinoGuire.rui`

---

## For colleagues (installation)

1. Clone or copy the repo
2. Edit `install.py` if your local path differs
3. Run `install.py` from the Rhino Python editor
4. `Tools > Toolbar Layout > Open > ui/RhinoGuire.rui`
