# Data Visualization Tool for Rhino

A workflow for managing and visualizing object metadata in Rhino through color-coded displays.

## 🚀 Quick Start (Recommended)

### Single Unified Tool
**Just run:** `DataVisualizationTool.py` in Rhino

This opens a single interface with all three workflow steps:
1. **Initialize Keys** - Apply metadata to objects
2. **Extract Values** - Generate color mapping
3. **Visualize** - Interactive color manager

The tool stays on top and guides you through each step!

---

## 📋 Complete Workflow Guide

### Initial Setup (One-Time)

1. **Generate Excel Templates:**
   ```bash
   python create_template.py
   python create_uniquevalues_template.py
   ```
   This creates:
   - `MetadataTemplate.xlsx` - For defining metadata keys
   - `UniqueValuesColorSettings.xlsx` - For color mappings

2. **Install Excel CSV Add-in (Recommended):**
   - Install **"CSV Importer Exporter"** Excel add-in
   - This makes CSV export much easier
   - Alternative: Use "Save As" → CSV format in Excel

### Step-by-Step Workflow

#### 1. Define Your Metadata Structure

**Duplicate** `MetadataTemplate.xlsx` for each category:
- `Buildings.xlsx`
- `Trees.xlsx`
- `Roads.xlsx`
- etc.

**Edit** the `_KeyDefinitions` table in each file:
```
Key Name         | Default Value
-----------------|---------------
Building Year    | 
Latest Renovation| 
Condition        | Good
Department       |
```

**Export to CSV:**
- Using CSV Importer Exporter: Select table → Export → Save to `_CSVOutput/Buildings.csv`
- OR Save As: File → Save As → CSV format → `_CSVOutput/Buildings.csv`

#### 2. Apply Keys to Objects (Step 1)

**Option A - Unified Tool (Recommended):**
1. Run `DataVisualizationTool.py`
2. Click "Select CSV & Apply Keys"
3. Select your CSV (e.g., `Buildings.csv`)
4. Select objects in Rhino
5. Keys are added!

**Option B - Individual Script:**
1. Select objects in Rhino
2. Run `01_InitializeKeys.py`
3. Select CSV file

**Result:** Objects now have metadata keys (empty or with default values)

#### 3. Fill in Values

Use Rhino's `SetUserText` command:
- Select object → Type `SetUserText` → Enter
- Key: `Building Year` → Value: `1950`
- Repeat for other keys and objects

#### 4. Generate Color Map (Step 2)

**In the Unified Tool:**
1. Click "Scan & Generate CSV"
2. Select all objects with metadata
3. Save as `_CSVOutput/UniqueValuesColorSettings.csv`

**Result:** CSV file with all unique values found

#### 5. Define Colors in Excel

1. Open `UniqueValuesColorSettings.xlsx`
2. Import the CSV data (or use CSV Importer Exporter)
3. Fill in R, G, B, A columns for each value:
   ```csv
   Key          | Value | R   | G   | B   | A
   -------------|-------|-----|-----|-----|----
   Building Year| 1950  | 100 | 100 | 100 | 255
   Building Year| 1990  | 180 | 180 | 180 | 255
   Condition    | Good  | 0   | 200 | 0   | 255
   Condition    | Bad   | 220 | 20  | 20  | 255
   ```
4. Export back to CSV: `_CSVOutput/UniqueValuesColorSettings.csv`

#### 6. Visualize! (Step 3)

**In the Unified Tool:**
1. Click "Open Color Manager"
2. Select objects to visualize
3. Select `UniqueValuesColorSettings.csv`
4. GUI opens with:
   - Dropdown to select key
   - Color legend display
   - Default color picker
   - Error detection
5. Click "Update Colors"
6. Objects are colored!

**Features:**
- Change keys and update instantly
- See warnings for missing colors
- Select problem objects with one click
- Screenshot-friendly legend

---

## 📁 Folder Structure

```
DataVisualization/
├── DataVisualizationTool.py           ⭐ RUN THIS - Unified interface
├── create_template.py                 # One-time: Generate MetadataTemplate.xlsx
├── create_uniquevalues_template.py    # One-time: Generate UniqueValuesColorSettings.xlsx
├── MetadataTemplate.xlsx              # Template - duplicate for categories
├── UniqueValuesColorSettings.xlsx     # Color mapping workbook
├── 01_InitializeKeys.py               # Optional: Run individually
├── 02_ExtractUniqueValues.py          # Optional: Run individually
├── 03_ColorManager.py                 # Optional: Run individually
├── README.md
├── PROGRESS.md
└── _CSVOutput/                        # CSV files go here
    ├── Buildings.csv
    ├── Trees.csv
    ├── Roads.csv
    └── UniqueValuesColorSettings.csv
```

---

## 🎨 Color Format - RGBA

Each color needs 4 values:
- **R** = Red (0-255)
- **G** = Green (0-255)
- **B** = Blue (0-255)
- **A** = Alpha/Opacity (0-255, optional, defaults to 255)

### Example Color Schemes

**Building Age Gradient (Dark → Light):**
```
1950: 80,80,80,255      # Dark gray (old)
1990: 150,150,150,255   # Medium gray
2020: 220,220,220,255   # Light gray (new)
```

**Condition (Traffic Light):**
```
Good:     0,200,0,255   # Green
Fair:     255,255,0,255 # Yellow
Bad:      220,20,20,255 # Red
```

**Departments (Distinct Colors):**
```
A: 255,100,100,255      # Coral
B: 100,150,255,255      # Blue
C: 150,200,100,255      # Lime
D: 255,180,80,255       # Orange
```

**Priority Actions:**
```
Demolish:      180,0,0,255     # Dark red (urgent)
Refurbish:     255,140,0,255   # Orange (action)
Not necessary: 100,200,100,255 # Light green (good)
```

---

## 💡 Why Excel + CSV?

**Excel (.xlsx):**
- Easy visual editing with tables
- Professional formatting
- Good for organization and version control
- Share with team members who don't use Rhino

**CSV (.csv):**
- Rhino compatibility (no Python dependencies!)
- Lightweight and fast
- Scripts auto-detect `_CSVOutput` folder

**Workflow:** Edit in Excel → Export to CSV → Scripts read CSV

---

## 🔧 Requirements

**For Rhino Scripts:**
- Rhino 7 or 8
- **No external Python libraries needed!** (uses built-in csv module)

**For Template Generators (one-time setup):**
- Python with openpyxl:
  ```bash
  pip install openpyxl
  ```

**For Easy CSV Export (recommended):**
- Excel add-in: "CSV Importer Exporter"

---

## 🎯 Key Features

### Unified Tool (DataVisualizationTool.py)
- ✅ Single interface for entire workflow
- ✅ Stays on top during use
- ✅ Auto-detects `_CSVOutput` folder
- ✅ No external script dependencies
- ✅ Fully self-contained

### Step 1: Initialize Keys
- ✅ Never overwrites existing keys/values
- ✅ Can apply multiple CSV files to same objects
- ✅ Progress tracking

### Step 2: Extract Values
- ✅ Scans all metadata keys automatically
- ✅ Suggests default filename
- ✅ Preserves existing colors when re-run

### Step 3: Color Manager
- ✅ Interactive Eto Forms GUI
- ✅ Dropdown to select active key
- ✅ Color legend with RGB swatches
- ✅ Default color picker
- ✅ "Update" button refreshes from CSV
- ✅ **Comprehensive error detection:**
  - Values in objects but missing from CSV
  - Values in CSV but not in selection
  - Objects without the active key
- ✅ "Select Problem Objects" button
- ✅ Screenshot-friendly for presentations
- ✅ Only affects selected objects

---

## 🔄 Two Ways to Use

### Method 1: Unified Tool (Recommended)
Run `DataVisualizationTool.py` → All steps in one interface

**Best for:**
- New users
- Complete workflows
- Keeping organized

### Method 2: Individual Scripts
Run `01_InitializeKeys.py`, `02_ExtractUniqueValues.py`, `03_ColorManager.py` separately

**Best for:**
- Advanced users
- Partial workflows
- Custom automation

Both methods work identically - choose what fits your workflow!

---

## 📝 CSV Export Instructions

### Recommended: CSV Importer Exporter Add-in

1. Install the Excel add-in "CSV Importer Exporter"
2. In Excel: Select your table data
3. Add-in menu → Export to CSV
4. Save to `_CSVOutput` folder

### Alternative: Excel Save As

1. Select and copy your table data
2. Paste into new workbook
3. File → Save As → "CSV (Comma delimited) (*.csv)"
4. Save to `_CSVOutput` folder

**Note:** Direct Excel table export is not available in all Excel versions, hence the add-in recommendation.

---

## 🐛 Troubleshooting

**"No keys found in CSV":**
- Ensure CSV has header: `Key Name,Default Value`
- Check you exported the table, not blank rows

**"Objects not coloring":**
- Verify RGB values are filled in (not empty)
- Check Warnings panel in Color Manager
- Use "Select Problem Objects" button

**"Unified tool loses focus":**
- Tool is set to stay on top
- Returns focus automatically after steps
- If issues persist, close and restart

**"Missing CSV folder":**
- Create `_CSVOutput` folder in same directory as scripts
- Scripts will auto-detect it

**"Template generators fail":**
- Install openpyxl: `pip install openpyxl`
- Only needed once for setup

---

## 💾 Tips & Best Practices

1. **Organization:**
   - Keep CSV files in `_CSVOutput` folder
   - Version control Excel files
   - Name CSVs clearly (Buildings.csv, Trees.csv)

2. **Workflow:**
   - Use unified tool for complete workflows
   - Export to CSV before running scripts
   - Test with small selection first

3. **Colors:**
   - Use meaningful color schemes
   - Document your color logic
   - Screenshot legends for presentations

4. **Data:**
   - Script 02 preserves colors when re-run
   - Safe to add new values anytime
   - Keys from different CSVs can coexist

5. **Performance:**
   - Select only needed objects for Step 3
   - Unselect objects preserves their colors
   - Close Color Manager when done

---

## 📞 Support

For detailed development notes and feature status, see `PROGRESS.md`

---

## 🎓 Example Complete Workflow

1. **Setup:** Run template generators once
2. **Define:** Duplicate `MetadataTemplate.xlsx` → `Buildings.xlsx`, edit keys, export to CSV
3. **Apply:** Run unified tool → Step 1 → Select `Buildings.csv` + objects
4. **Fill:** Use `SetUserText` in Rhino to add values
5. **Extract:** Unified tool → Step 2 → Scan objects, save CSV
6. **Color:** Open `UniqueValuesColorSettings.xlsx`, import CSV, fill RGB values, export
7. **Visualize:** Unified tool → Step 3 → Select objects, choose key, update colors
8. **Present:** Screenshot the color legend for your presentation!

**Total time after setup: 5-10 minutes** ⚡
