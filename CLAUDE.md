# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running Tools

All tools run inside **Rhino 8** — they cannot be executed as standalone scripts (they depend on `rhinoscriptsyntax`, `Rhino.Geometry`, `Eto.Forms`, and `scriptcontext`).

**From the Rhino command line:**

```text
RunPythonScript  →  navigate to the tool's .py file
```

**Via toolbar macro (preferred for buttons):**

```text
! _-RunPythonScript "C:/path/to/RhinoGuire/launch_lindero.py"
```

**Via the central launcher (passing a key):**

```text
! _-RunPythonScript "C:/path/to/RhinoGuire/launch.py" "RG_Lindero"
```

Available keys: `RG_Lindero`, `RG_Arriero`, `RG_Chivito`, `RG_Sebucan`, `RG_Baquiano`, `RG_PadGrader`, `RG_WayGrader`, `RG_CutFillReport`.

## Running Tests

The only headless-testable code is `TerrainTools/_core` (pure Python + RhinoCommon-free logic):

```sh
python TerrainTools/_core/tests/test_headless.py
```

Covers slope unit conversions and grid-prism volume calculations. All other code requires a running Rhino session.

## Architecture

### One tool = one file

Each tool is a self-contained Python script. There is no build step, no package install, and no shared state between tool files — each one bootstraps its own import path and reloads `ui/theme.py` at startup.

**Path bootstrap pattern (every tool):**

```python
import sys as _sys, os as _os
_rg_root = _os.path.normpath(_os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "..", ".."))
if _rg_root not in _sys.path:
    _sys.path.insert(0, _rg_root)
from ui import theme as _t
import importlib as _importlib; _importlib.reload(_t)
```

The reload is intentional: it picks up live theme changes without restarting Rhino.

### Shared design system — `ui/theme.py`

Single source for the neo-brutalist color palette, 8pt-grid spacing constants, and factory helpers `lbl()`, `btn()`, `hint()`, `section_header()`. All tools import this. Never hardcode colors or font sizes in a tool file.

`prefs_get(key, fallback)` / `prefs_set(key, path)` persist folder memories to `_prefs.json` at the repo root.

### TerrainTools `_core` — the exception to the "one file" rule

`TerrainTools/_core/` is a shared pure-Python + RhinoCommon engine imported by the three terrain tools. It has no Eto dependency and is designed to be headless-testable. The module boundary is strict: `_core` never calls `rs.*` or touches UI.

| Module | Responsibility |
| --- | --- |
| `slope.py` | Unit conversions: H:V ↔ % ↔ ° ↔ canonical gradient `m = V/H` |
| `terrain.py` | `TerrainModel` wrapping a Mesh; `project_z(x, y)` via vertical raycast (same technique as Sebucan) |
| `grading.py` | Analytical heightfield builder; `GradeResult` dataclass; daylight-line cut/fill rule (see D5 in `DECISIONS.md`) |
| `volumes.py` | Grid-prism cut/fill totals; per-station mass-haul for corridors |
| `meshbuild.py` | `GradeResult` → triangulated Mesh with vertex-color depth tinting |
| `report.py` | openpyxl workbook writer + Eto `Drawable`→`Bitmap` PNG helpers |

`TerrainTools/_widgets.py` sits above `_core` and provides the shared Eto `SlopeInput` row widget used by all three terrain tool UIs.

### Launcher shims

`launch.py` at the root is the canonical entry point. The eight `launch_<name>.py` files are thin shims that call it — they exist so Rhino toolbar macros can reference a stable absolute path without passing a key argument.

`install.py` is **deprecated** (kept for historical reference only — do not update it).

## Key Conventions

### Script header (every tool)

```python
#! python3          # CPython 3 engine
# r: openpyxl       # auto-install via Rhino package manager (only when needed)
```

### Eto.Forms rules

- All windows are **modeless** (`forms.Form` + `.Show()`). Never use `ShowModal()` on a `Form`.
- Always set `self.Owner = Rhino.UI.RhinoEtoApp.MainWindow`.
- `super().__init__()` must be called **before** setting any property.
- Never use keyword arguments in .NET constructors: `label = forms.Label(); label.Text = "x"` not `forms.Label(Text="x")`.
- `TableLayout` does not support dynamic row add/remove after construction — use `StackLayout` for dynamic content.

### openpyxl dependency

Only Arriero, Chivito, and CutFillReport require it. The `# r: openpyxl` header causes Rhino to install it automatically on first run. Do not add other external packages.

### Geometry convention

All tools output **new objects** — original geometry is never modified in-place. New meshes go on dedicated sub-layers (e.g., `TerrainTools::Graded`). Always call `sc.doc.Views.Redraw()` after adding geometry.

### Cross-tool state

PadGrader and WayGrader write their last result to `sc.sticky['terraintools_last_grade']` so CutFillReport can read it in "from last grading" mode.
