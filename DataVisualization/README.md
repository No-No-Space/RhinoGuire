# Data Visualization Tool for Rhino

A workflow for managing and visualizing object metadata in Rhino through color-coded displays.

## 📋 Quick Start Guide

### Step 1: Generate Excel Templates
Run these Python scripts **once** to create your Excel templates:

```bash
python create_template.py
python create_uniquevalues_template.py
```

This creates:
- `MetadataTemplate.xlsx` - For defining metadata keys
- `UniqueValuesColorSettings.xlsx` - For color mappings (populated by Script 02)

### Step 2: Define Your Metadata Keys
1. **Duplicate** `MetadataTemplate.xlsx` for each category:
   - `Buildings.xlsx`
   - `Trees.xlsx`  
   - `Roads.xlsx`
   - etc.

2. **Edit** the `_KeyDefinitions` table in each file with your metadata keys

3. **Export** each table to CSV:
   - Save to `_CSVOutput/Buildings.csv`
   - Save to `_CSVOutput/Trees.csv`
   - etc.

### Step 3: Apply Keys to Objects in Rhino
1. Select objects in Rhino (e.g., all buildings)
2. Run `01_InitializeKeys.py`
3. Select your CSV (e.g., `Buildings.csv`)
4. Keys are added to selected objects

### Step 4: Fill in Values
Use Rhino's `SetUserText` command to add values:
- Select object → `SetUserText`
- Key: "Building Year" → Value: "1950"

### Step 5: Generate Color Map
1. Select all objects with metadata
2. Run `02_ExtractUniqueValues.py`
3. Save as `_CSVOutput/UniqueValuesColorSettings.csv`
4. **Import CSV to Excel** (or edit CSV directly):
   - Open `UniqueValuesColorSettings.xlsx`
   - Import the CSV data
   - Fill in R, G, B, A color values

### Step 6: Visualize!
1. Select objects to visualize
2. Run `03_ColorManager.py`
3. Select `_CSVOutput/UniqueValuesColorSettings.csv`
4. Choose key from dropdown
5. Click "Update Colors"
6. Take screenshots of the legend for presentations!

---

## 📁 Folder Structure

```
DataVisualization/
├── create_template.py                 # Run once to create MetadataTemplate.xlsx
├── create_uniquevalues_template.py    # Run once to create UniqueValuesColorSettings.xlsx
├── MetadataTemplate.xlsx              # Template (duplicate for Buildings.xlsx, etc.)
├── UniqueValuesColorSettings.xlsx     # Color mapping workbook
├── 01_InitializeKeys.py               # Rhino script: Apply keys to objects
├── 02_ExtractUniqueValues.py          # Rhino script: Generate color CSV
├── 03_ColorManager.py                 # Rhino script: Interactive visualization
├── README.md
├── PROGRESS.md
└── _CSVOutput/                        # Place all CSV exports here
    ├── Buildings.csv
    ├── Trees.csv
    └── UniqueValuesColorSettings.csv
```

---

## 🎨 Why Excel + CSV?

**Excel files** (.xlsx):
- Easy visual editing
- Professional table formatting
- Good for organization and version control

**CSV files** (.csv):
- Rhino compatibility (no external Python dependencies)
- Scripts read from CSV
- Lightweight and fast

**Workflow**: Edit in Excel → Export to CSV → Scripts read CSV

---

## 📝 File Formats

### MetadataTemplate.xlsx → Buildings.csv
Export the `_KeyDefinitions` table as CSV:

```csv
Key Name,Default Value
Building Year,
Latest Renovation,
Condition,Good
Department,
```

### UniqueValuesColorSettings.xlsx → UniqueValuesColorSettings.csv
After running Script 02, fill in colors:

```csv
Key,Value,R,G,B,A
Building Year,1950,120,120,120,255
Building Year,1975,180,180,180,255
Condition,Good,0,255,0,255
Condition,Fair,255,255,0,255
Condition,Poor,255,0,0,255
```

**Color Format - RGBA:**
- R = Red (0-255)
- G = Green (0-255)
- B = Blue (0-255)
- A = Alpha/Opacity (0-255, optional, defaults to 255)

---

## ✨ Key Features

### Script 01: Initialize Keys
- ✅ Auto-detects `_CSVOutput` folder
- ✅ Never overwrites existing keys or values
- ✅ Can apply multiple CSV files to same objects
- ✅ Progress tracking and summary

### Script 02: Extract Unique Values
- ✅ Scans selected objects for all metadata keys
- ✅ Auto-detects `_CSVOutput` folder for saving
- ✅ Preserves existing color assignments when re-run
- ✅ Suggests default filename

### Script 03: Color Manager
- ✅ Interactive Eto Forms GUI
- ✅ Auto-detects `_CSVOutput` folder
- ✅ Dropdown to select active key
- ✅ Color legend with swatches
- ✅ Default color picker
- ✅ "Update" button to refresh from CSV
- ✅ **Error detection and reporting:**
  - Values in objects but missing from CSV
  - Values in CSV but not in selection
  - "Select Problem Objects" button
- ✅ Screenshot-friendly legend
- ✅ Only affects selected objects

---

## 🔧 Requirements

- **Rhino 7 or 8**
- **No external Python libraries needed for Rhino scripts** (uses built-in csv module)
- **openpyxl** only needed for template generators (one-time setup):
  ```bash
  pip install openpyxl
  ```

---

## 💡 Tips & Best Practices

1. **Keep CSV files in `_CSVOutput`** - Scripts auto-detect this folder
2. **Version control Excel files** - Easier to track changes than CSV
3. **Export to CSV before running scripts** - Scripts read from CSV, not Excel
4. **Script 02 preserves colors** - Safe to re-run when adding new values
5. **Use meaningful key names** - They appear in the GUI dropdown
6. **Test with small selection first** - Verify colors before applying to all objects
7. **Take screenshots of legend** - Great for presentations and documentation

---

## 🐛 Troubleshooting

**Scripts can't find `_CSVOutput` folder:**
- Make sure `_CSVOutput` is in the same folder as the scripts
- File dialogs will open normally if folder doesn't exist

**"No keys found in CSV":**
- Check CSV has header row: `Key Name,Default Value`
- Make sure you exported the table, not the entire worksheet

**Objects not coloring:**
- Verify CSV has R,G,B values filled in (not empty)
- Check the Warnings panel in Script 03 for specific errors
- Use "Select Problem Objects" button to find issues

**Template generators fail:**
- Install openpyxl: `pip install openpyxl`
- Only needed once for initial setup

---

## 📊 Example Color Schemes

**Building Age (gradient dark → light):**
- 1950: `80,80,80,255` (dark gray)
- 1990: `150,150,150,255` (medium gray)  
- 2020: `220,220,220,255` (light gray)

**Condition (traffic light):**
- Good: `0,200,0,255` (green)
- Fair: `255,255,0,255` (yellow)
- Poor: `255,0,0,255` (red)

**Priority Actions:**
- Demolish: `200,0,0,255` (dark red - urgent)
- Refurbish: `255,165,0,255` (orange - action needed)
- Maintain: `100,200,100,255` (light green - good status)

---

## 📞 Support

For questions or issues, check `PROGRESS.md` for development notes and feature status.
