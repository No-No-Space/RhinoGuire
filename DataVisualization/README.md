# Chivito - Data Visualization Tool for Rhino

A workflow for managing and visualizing object metadata in Rhino 8 through color-coded displays using Excel files.

## Features

- **Three-step workflow** in a single unified interface
- **Non-blocking UI** - interact with Rhino while the tool is open
- **Color Manager** with interactive legend, color picker, and error detection
- **Smart merge** - re-exporting preserves existing color definitions in Excel
- **Export Legend as PNG** for presentations
- **Capture Viewport as PNG** directly from the Color Manager
- **Optimized color updates** using RhinoCommon direct access for speed

## Requirements

- **Rhino 8** with CPython 3
- **openpyxl** library for Excel operations

The script header includes `# r: openpyxl`, which tells Rhino 8 to install the package automatically on first run.

If that does not work, install it manually:

1. Open Rhino 8
2. Go to **Tools > PythonScript > Edit...**
3. In the Script Editor menu: **Tools > Package Manager**
4. Search for `openpyxl` and install it

## How to Use

### Running the Script

1. In Rhino 8, type `RunPythonScript` in the command line
2. Navigate to `Chivito.py` and click **Open**
3. A non-blocking window opens with three step buttons

### Step 1: Initialize Keys

Apply metadata keys from an Excel template to selected objects.

1. Click **"Select Excel & Apply Keys"**
2. Select an Excel file with key definitions (columns: Key Name, Default Value)
3. Select objects in Rhino
4. Keys are added to objects without overwriting existing values

### Step 2: Extract Unique Values

Scan objects and generate a color mapping Excel file.

1. Click **"Scan & Generate Excel"**
2. Select objects with metadata
3. Choose a save location (default: `_ExcelOutput/UniqueValuesColorSettings.xlsx`)
4. If the file already exists, existing color definitions are preserved

**After export:** Open the Excel file and fill in the R, G, B, A columns for each value.

### Step 3: Visualize with Colors

Open the interactive Color Manager.

1. Click **"Open Color Manager"**
2. Select objects to visualize
3. Select the color mapping Excel file
4. The Color Manager opens (non-blocking, Rhino stays interactive):
   - **Select a key** from the dropdown
   - **Click "Update Colors"** to apply colors to objects
   - **Change the default color** for objects without a matching value
   - **Review warnings** for missing color definitions or unused entries
   - **"Select Problem Objects"** to highlight objects with issues
   - **"Export Legend as PNG"** to save the color legend
   - **"Capture Viewport as PNG"** to save the current Rhino viewport

## Folder Structure

```text
DataVisualization/
├── Chivito.py                  <-- Run this script
├── README.md
└── _ExcelOutput/               Excel files go here
```

## Color Format - RGBA

Each color needs 4 values in the Excel file:

| Column | Description | Range |
|--------|-------------|-------|
| R | Red | 0-255 |
| G | Green | 0-255 |
| B | Blue | 0-255 |
| A | Alpha/Opacity | 0-255 (defaults to 255) |

### Example Color Schemes

**Building Age Gradient (Dark to Light):**

```text
1950: 80,80,80,255        (Dark gray - old)
1990: 150,150,150,255     (Medium gray)
2020: 220,220,220,255     (Light gray - new)
```

**Condition (Traffic Light):**

```text
Good:  0,200,0,255        (Green)
Fair:  255,255,0,255      (Yellow)
Bad:   220,20,20,255      (Red)
```

## Excel File Formats

### Key Definitions (Step 1 input)

```text
| Key Name          | Default Value |
|-------------------|---------------|
| Building Year     |               |
| Condition         | Good          |
```

### Color Mapping (Step 2 output / Step 3 input)

```text
| Key           | Value | R   | G   | B   | A   |
|---------------|-------|-----|-----|-----|-----|
| Building Year | 1950  | 80  | 80  | 80  | 255 |
| Condition     | Good  | 0   | 200 | 0   | 255 |
```

## Troubleshooting

### "openpyxl library is not available"

Install openpyxl using the Package Manager in Rhino's Script Editor, or use `# r: openpyxl` in the script header.

### Objects not coloring

- Verify RGB values are filled in the Excel file (not empty)
- Check the Warnings panel in Color Manager
- Use "Select Problem Objects" to identify the issue

### "No keys found in Excel file"

- Ensure the Excel file has the correct format (Key Name in column A)
- Check that the data starts from row 2 (row 1 is the header)

## Version History

- **v0.5** (2026-02-14): Unified release
  - Non-blocking modeless interface (Rhino stays interactive)
  - Viewport capture as PNG
  - Optimized color updates with RhinoCommon direct access
  - Smart merge preserving existing colors on re-export
  - Legend export as PNG
  - Migrated from CSV to Excel (openpyxl)
  - CPython 3 with Eto.Forms

## Author

Aquelon - aquelon@pm.me
