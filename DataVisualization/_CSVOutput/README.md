# Export Instructions

Before running the Rhino scripts, you need to export your Excel tables to CSV format.

## From MetadataTemplate.xlsx (or Buildings.xlsx, Trees.xlsx, etc.)

1. Open your Excel file
2. Click on any cell in the **_KeyDefinitions** table
3. **Excel 2016+:** Data tab → "From Table/Range" → Export
   **OR manually:** Select the table (A3 and below) → Copy → Paste into new workbook → Save As CSV
4. Save as: `_CSVOutput/[YourFileName].csv`
   - Example: `_CSVOutput/Buildings.csv`
   - Example: `_CSVOutput/Trees.csv`

**Important:** The CSV must have this structure:
```
Key Name,Default Value
Building Year,
Latest Renovation,
Condition,Good
```

## From UniqueValuesColorSettings.xlsx

After running Script 02, you'll have color mapping tables. To use them with Script 03:

1. Open `UniqueValuesColorSettings.xlsx`
2. Fill in the R, G, B, A values for each unique value
3. Export the entire worksheet (or the data range) as CSV
4. Save as: `_CSVOutput/UniqueValuesColorSettings.csv`

**Important:** The CSV must have this structure:
```
Key,Value,R,G,B,A
Building Year,1950,120,120,120,255
Building Year,1975,180,180,180,255
Condition,Good,0,255,0,255
Condition,Fair,255,255,0,255
Condition,Poor,255,0,0,255
```

## Quick Export Tips

**Method 1 - Save As CSV:**
- File → Save As → Choose "CSV (Comma delimited)" format
- Only saves the active sheet

**Method 2 - Copy/Paste:**
- Select your data range
- Copy (Ctrl+C)
- Open Notepad or text editor
- Paste
- Save with `.csv` extension

**Method 3 - Export Table (Excel 2016+):**
- Select table
- Right-click → Export → Export to CSV

## File Locations

All CSV files should be placed in:
```
DataVisualization\_CSVOutput\
```

This keeps your main folder clean and organized.
