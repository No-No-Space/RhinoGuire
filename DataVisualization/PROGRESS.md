# Data Visualization Tool - Development Progress

## ✅ Phase 1: Structure & Template (COMPLETE)

### Files Created:
1. **README.md** - Project documentation
2. **create_template.py** - Generates MetadataTemplate.xlsx
3. **MetadataTemplate.xlsx** - Key definitions template
4. **01_InitializeKeys.py** - Applies keys to objects

### Key Features:
- Multiple Excel files approach (Option A)
- User can create Buildings.xlsx, Trees.xlsx, Roads.xlsx from template
- Keys from different files can coexist on same objects
- No overwriting of existing keys or values

## ✅ Phase 2: Extract Unique Values (COMPLETE)

### Files Created:
5. **create_uniquevalues_template.py** - Generates UniqueValuesColorSettings.xlsx
6. **UniqueValuesColorSettings.xlsx** - Color mapping template
7. **02_ExtractUniqueValues.py** - Scans objects and creates color map tables

### Script 02 Features:
- ✅ Scans selected objects for ALL metadata keys (regardless of source Excel)
- ✅ Collects unique values for each key
- ✅ Creates `_ColorMap_[KeyName]` tables in 6-column grid layout
- ✅ Table structure: Title row + Headers (Value | R | G | B | A) + Data rows
- ✅ Updates existing tables when re-run
- ✅ 30-row vertical spacing between table rows
- ✅ 7-column horizontal spacing per table (5 data + 2 gap)
- ✅ Professional formatting with styled headers

### Excel File Strategy:
**Two-file approach:**
1. **MetadataTemplate.xlsx** (user duplicates per category)
   - Contains: `_KeyDefinitions` table
   - Used by: Script 01
   - Examples: Buildings.xlsx, Trees.xlsx, Roads.xlsx

2. **UniqueValuesColorSettings.xlsx** (single file per project)
   - Contains: All `_ColorMap_*` tables in grid layout
   - Used by: Script 02 (writes), Script 03 (reads)
   - Consolidates all categories in one place

### Color Format - RGBA:
- Format: `R,G,B,A` (each 0-255)
- Alpha optional (defaults to 255 if empty)
- Examples:
  - `190,190,190,255` - Opaque gray
  - `255,0,0,128` - Semi-transparent red
  - `0,128,255` - Blue (A defaults to 255)

## 🔄 Phase 3: Color Manager GUI (NEXT)

### Requirements:
- **Eto Forms GUI** (cross-platform)
- **Controls:**
  - Dropdown: Select active key from available keys
  - Color legend: Display all values and their colors
  - Default color picker: For objects without the active key
  - "Update" button: Refresh colors from Excel
  - "Close" button: Exit tool
  
- **Behavior:**
  - Only affects selected objects (preserves other object colors)
  - GUI stays open for multiple updates
  - Screenshot-friendly legend for presentations
  - Real-time color application

### Technical Considerations:
- Parse RGBA from Excel
- Handle missing/invalid color values gracefully
- Use `rs.ObjectColor()` to set display colors
- Detect which keys are available in Excel file
- Show key name + value count in dropdown

## 📁 Current Folder Structure:
```
DataVisualization/
├── README.md
├── PROGRESS.md
├── create_template.py
├── create_uniquevalues_template.py
├── MetadataTemplate.xlsx
├── UniqueValuesColorSettings.xlsx
├── 01_InitializeKeys.py
├── 02_ExtractUniqueValues.py
└── 03_ColorManager.py (TODO)
```

## Complete Workflow (Once Script 03 is done):

1. **Setup:**
   - Duplicate MetadataTemplate.xlsx → Buildings.xlsx, Trees.xlsx, etc.
   - Define keys in each category file
   - Run create_uniquevalues_template.py once

2. **Apply Keys:**
   - Select objects → Run Script 01 → Choose Buildings.xlsx
   - Repeat for different categories

3. **Fill Values:**
   - Use SetUserText in Rhino to fill metadata values
   - Example: SetUserText "Building Year" = "1950"

4. **Extract & Color:**
   - Select all objects → Run Script 02 → Select UniqueValuesColorSettings.xlsx
   - Open Excel, fill in RGBA colors
   - Run Script 03 → Select key → Objects colored automatically

5. **Iterate:**
   - Change colors in Excel
   - Click "Update" in Script 03 GUI
   - Take screenshots of legend for presentations

## Next Steps:
1. ✅ Create Script 02 - DONE
2. 🔄 Create Script 03 (Color Manager GUI) - IN PROGRESS
3. Full integration testing
4. Documentation and examples
