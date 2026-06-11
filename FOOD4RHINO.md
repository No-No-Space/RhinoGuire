# Food4Rhino Submission Guide

Checklist and notes for publishing RhinoGuire to [food4rhino.com](https://food4rhino.com) via the Yak Package Manager.

---

## Status: what's ready vs. what's missing

| Item | Status | Notes |
| --- | --- | --- |
| `manifest.yml` — name, version, keywords, url | Ready | Minor fixes needed (see below) |
| `manifest.yml` — `authors` field | Fix needed | Currently `author:` (singular); must be `authors:` list |
| `manifest.yml` — `icon` field | Missing | Must point to a bundled PNG file |
| `manifest.yml` — `description` | Improve | Food4Rhino expects: *name + verb + what it does + who it's for* |
| Icon file (64×64 PNG) | Missing | Food4Rhino rejects blank/missing icons |
| `.yakignore` | Missing | Without it, `yak build` bundles `__pycache__`, `.git`, dev docs, etc. |
| `install.py` | Delete | Deprecated; confusing for Food4Rhino users; replaced entirely by `launch.py` |
| Toolbar (`.rui`) | Decision needed | See "Toolbar portability" below |
| Screenshots | Missing | Needed on the Food4Rhino listing page; not part of the `.yak` file |
| LICENSE | Ready | MIT — compatible with Rhino/Food4Rhino (GPL is not compatible) |

---

## Step-by-step

### 1. Delete `install.py`

The file is deprecated and its own docstring says it can be safely deleted. Shipping it in the Yak package would confuse users who downloaded from Food4Rhino.

### 2. Create the icon

Food4Rhino requires a **custom-drawn icon** — no blank icons, no screenshots used as icons.

- Format: PNG (preferred) or JPEG
- Size: 64×64 px (displayed at 32×32 on the listing)
- Place at the repo root as `icon.png`
- The `ui/InfoAboutIcons.txt` references [line-md](https://icon-sets.iconify.design/line-md/) — those icons could work as a starting point for a composed logo

### 3. Fix `manifest.yml`

```yaml
name: RhinoGuire
version: 0.1.0
authors:
  - Aksel Alvarez
description: >
  RhinoGuire is a collection of Python 3 tools for Rhino 8 covering object
  metadata management, footprint area calculation, data visualization,
  mesh-on-terrain projection, and terrain grading (cut/fill earthwork).
  Aimed at architects and landscape architects working with BIM data in Rhino.
url: https://github.com/No-No-Space/RhinoGuire
icon: icon.png
keywords:
  - rhinoguire
  - python
  - bim
  - data
  - mesh
  - terrain
  - grading
  - cut-fill
  - earthwork
rhino: ">=8.0"
```

Key changes from the current file:

- `author:` → `authors:` (list format — required by Yak)
- Added `icon: icon.png`
- Expanded `description` (name + verb + what it does + audience)

### 4. Create `.yakignore`

Place this at the repo root. Same syntax as `.gitignore`.

```gitignore
# Version control
.git/
.gitignore

# Python cache
__pycache__/
*.py[cod]
*.pyo

# Development-only docs
TerrainTools/PLAN.md
TerrainTools/DECISIONS.md
FOOD4RHINO.md
CLAUDE.md

# Runtime outputs and local state
_prefs.json
DataVisualization/_ExcelOutput/
*.bak
*.rui.bak

# Deprecated
install.py

# IDE / editor
.vscode/
.cursor/
.cursorignore
.cursorindexingignore
```

### 5. Decide on toolbar (`.rui`) distribution

`ui/RhinoGuire.rui` is gitignored because it contains **hardcoded absolute paths** in its button macros. This is the main unresolved issue.

**Option A — Don't ship the `.rui` (recommended for v0.1)**

Omit the toolbar from the Yak package entirely. Document that users run tools via `RunPythonScript` or create their own aliases. This is the standard approach for Python script collections on Food4Rhino. The `.rui` setup guide in `ui/README.md` covers the manual toolbar setup for users who want it.

Add `ui/RhinoGuire.rui` to `.yakignore`.

**Option B — Ship a portable `.rui`**

Rhino toolbar macros can reference the Yak install path via `%RHINOPACKAGES%\RhinoGuire\<version>\`. This would make the `.rui` self-contained, but the exact path format for Yak-installed packages should be verified first against Rhino 8 documentation before committing to this approach.

### 6. Take screenshots

Food4Rhino listings without screenshots get significantly less traffic. Screenshots are uploaded to the Food4Rhino web form — they are **not** part of the `.yak` file.

Suggested shots (one per major tool):

1. Arriero — the export/import window alongside a spreadsheet
2. Chivito — a color-coded Rhino viewport with the legend panel
3. Lindero — bullet charts from R1/R2 analysis
4. PadGrader or WayGrader — graded terrain mesh in the viewport

### 7. Build the `.yak` file

Run from the repo root (adjust path if Rhino is installed elsewhere):

```bat
"C:\Program Files\Rhino 8\System\yak.exe" build
```

This creates a file named `rhinoguire-0.1.0-rh8-win.yak` (or similar). Test the package locally by installing it in Rhino 8 before uploading:

```bat
"C:\Program Files\Rhino 8\System\yak.exe" install rhinoguire-0.1.0-rh8-win.yak
```

### 8. Submit to Food4Rhino

1. Log in at [food4rhino.com](https://food4rhino.com) (McNeel account)
2. Go to **My Content → Add Plugin**
3. Upload the `.yak` file and fill in the web form:
   - Title, description (can reuse the manifest description)
   - Category: **Rhino → Utilities** (or BIM/Architecture depending on the primary audience)
   - Compatible Rhino versions: `Rhino 8`
   - Icon (upload `icon.png`)
   - Screenshots (upload the shots from Step 6)
   - License: MIT
4. Submit — McNeel does a human review, typically 1–2 business days

---

## Reference

- [Yak manifest format](https://developer.rhino3d.com/guides/yak/the-package-manifest/)
- [Anatomy of a Yak package](https://developer.rhino3d.com/guides/yak/the-anatomy-of-a-package/)
- [Yak CLI reference](https://developer.rhino3d.com/guides/yak/yak-cli-reference/)
- [Food4Rhino FAQ](https://www.food4rhino.com/en/faq)
