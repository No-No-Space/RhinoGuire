# CSV Export Instructions

This folder (`_CSVOutput`) is where you should place all CSV files exported from your Excel templates.

## Why CSV Files?

The Rhino scripts read from **CSV files** instead of Excel files because:
- ✅ No external Python dependencies needed (uses built-in csv module)
- ✅ Works perfectly with Rhino's IronPython 2.7
- ✅ Fast and lightweight
- ✅ Universal format

You still use **Excel for editing** (easier, more visual), then **export to CSV** for the scripts to read.

---

## How to Export from Excel to CSV

### Method 1: CSV Importer Exporter Add-in (Recommended)

**Why recommended:** Excel's built-in CSV export options are limited or missing in many versions.

**Installation:**
1. Open Excel
2. Go to Insert → Get Add-ins (or Office Add-ins)
3. Search for "CSV Importer Exporter"
4. Install the add-in

**Usage:**
1. Open your Excel file (e.g., `Buildings.xlsx`)
2. Select the data range you want to export (including headers)
3. Click the CSV Importer Exporter add-in
4. Choose "Export to CSV"
5. Save to this `_CSVOutput` folder
6. Done!

### Method 2: Copy-Paste Method

**If you don't want to install an add-in:**

1. Open your Excel file
2. Select and copy your table data (Ctrl+C)
3. Open a new blank Excel workbook
4. Paste (Ctrl+V)
5. File → Save As
6. Choose "CSV (Comma delimited) (*.csv)" format
7. Save to this `_CSVOutput` folder
8. Click "Yes" to any warnings about features

### Method 3: Excel Save As (if available)

**Some Excel versions support direct CSV export:**

1. Open your Excel file
2. File → Save As
3. Choose location: this `_CSVOutput` folder
4. Choose format: "CSV (Comma delimited) (*.csv)"
5. Save
6. Only the active sheet will be saved

---

## Expected CSV Files

You should have these CSV files in this folder:

### From MetadataTemplate.xlsx variants:
- `Buildings.csv` - Building metadata keys
- `Trees.csv` - Tree metadata keys
- `Roads.csv` - Road metadata keys
- (Any other category CSVs you create)

**Format example:**
```csv
Key Name,Default Value
Building Year,
Latest Renovation,
Condition,Good
Department,
```

### From Script 02 output:
- `UniqueValuesColorSettings.csv` - Color mappings

**Format example:**
```csv
Key,Value,R,G,B,A
Building Year,1950,100,100,100,255
Building Year,1990,180,180,180,255
Condition,Good,0,200,0,255
Condition,Bad,220,20,20,255
```

---

## Workflow Tips

1. **Edit in Excel** - Use your Excel templates for comfortable editing
2. **Export to CSV** - Place exports in this folder
3. **Run Rhino scripts** - Scripts auto-detect this folder
4. **Update CSV** - When you change Excel, re-export the CSV

---

## Troubleshooting

**Scripts can't find CSV:**
- Make sure files are directly in `_CSVOutput`, not in subfolders
- Check file extension is `.csv` not `.xlsx`
- File dialog will default to this folder automatically

**CSV has wrong data:**
- Make sure you exported the correct sheet/range
- Open CSV in text editor to verify format
- Should be comma-separated, not semicolon or tab

**Special characters in CSV:**
- Save CSV as "UTF-8" encoding if you have special characters
- Most add-ins handle this automatically

---

## Quick Reference

**When to export:**
- After creating/editing key definitions in Excel templates
- After filling in color values in UniqueValuesColorSettings.xlsx

**File naming:**
- Use clear names: `Buildings.csv`, `Trees.csv`
- Match your Excel file names for clarity
- Color map is always: `UniqueValuesColorSettings.csv`

**Scripts that use this folder:**
- `DataVisualizationTool.py` - Auto-detects this folder
- `01_InitializeKeys.py` - Defaults to this folder
- `02_ExtractUniqueValues.py` - Saves here by default
- `03_ColorManager.py` - Opens here by default
