# Rhino Data Exporter/Importer

Export and import object metadata (User Keys/Values) between Rhino 8 and Excel files using GUID-based tracking.

## Features

- **Export**: Extract User Keys/Values from selected objects to timestamped Excel files
- **Import**: Update objects from Excel with flexible options for missing data
- **GUID Tracking**: Uses Rhino's native GUIDs for reliable object identification
- **Backup**: Optional automatic backup before import operations
- **Flexible Import Options**:
  - Create or skip missing keys
  - Handle empty cells with placeholder values or key deletion
  - Comprehensive error reporting

## Requirements

### Python Environment
- Rhino 8 with CPython 3 (Python 3.x)
- **openpyxl** library for Excel operations

### Installing openpyxl

1. Open Rhino 8
2. Go to **Tools → PythonScript → Edit...**
3. Open the **Script Editor**
4. In the menu, go to **Tools → Package Manager**
5. Search for `openpyxl` and install it

**Alternative (Command Line):**
```bash
# Find Rhino's Python executable (usually in Rhino installation folder)
# Then run:
python -m pip install openpyxl
```

## Usage

### Running the Script

1. In Rhino 8, type `RunPythonScript` in the command line
2. Navigate to: `D:\404-Github\008-RhinoGuire\RhinoGuire\DataExporterImporter\DataExporterImporter.py`
3. Select the file and click **Open**

### Export Workflow

1. Select objects in Rhino that have User Keys/Values you want to export
2. Run the script
3. Click **"1. Export Data (Select Objects → Excel)"**
4. Excel file is automatically saved with timestamp: `RhinoData_YYYY-MM-DD_HHMMSS.xlsx`

### Import Workflow

1. Edit the exported Excel file (or create a compatible one)
2. Run the script
3. Click **"2. Import Data (Excel → Update Objects)"**
4. Configure import options (see below)
5. Select the Excel file to import
6. Review the summary report

## Import Options

### Create backup before import
- **Default**: Checked ✓
- Creates a timestamped backup export of all current object data before making changes

### Create missing keys from Excel columns
- **Default**: Checked ✓
- If Excel has columns (keys) that don't exist on objects, they will be created
- If unchecked, only existing keys will be updated

### Update values when Excel cell is empty
- **Default**: Unchecked ☐
- **Unchecked**: Empty Excel cells are ignored, original Rhino values preserved
- **Checked**: Empty cells trigger special behavior based on placeholder setting

#### Placeholder value
- **Default**: `-`
- Only active when "Update values when Excel cell is empty" is checked
- **With placeholder** (e.g., `-`): Empty cells set key value to placeholder
- **Without placeholder** (blank field): Empty cells **DELETE the key entirely**

## Excel File Format

### Structure
```
| GUID                                  | Key1  | Key2      | Key3    | ...
|---------------------------------------|-------|-----------|---------|----
| 12345678-90ab-cdef-1234-567890abcdef | val1  | val2      | val3    | ...
| abcdef12-3456-7890-abcd-ef1234567890 | val1  |           | val3    | ...
```

### Requirements
- **First column MUST be named "GUID"**
- GUIDs must be valid Rhino object GUIDs (with or without dashes)
- Additional columns are treated as User Keys
- Keys are sorted alphabetically on export

## Output Location

All files are saved to:
```
D:\404-Github\008-RhinoGuire\RhinoGuire\DataExporterImporter\
```

This includes:
- Exported Excel files: `RhinoData_YYYY-MM-DD_HHMMSS.xlsx`
- Backup files (created before import)

## Error Handling

### Export Errors
- Shows error dialog if folder cannot be created
- Reports if no objects selected
- Lists any file save errors

### Import Errors
- Validates Excel format (first column must be GUID)
- Reports GUIDs that don't match any objects in the document
- Shows summary with counts:
  - Objects updated
  - Keys created
  - Keys updated
  - Keys deleted
  - GUIDs not found (lists first 5)

## Technical Notes

### GUID Persistence
- Rhino GUIDs are persistent across sessions
- GUIDs remain valid unless object is deleted/recreated
- Operations that preserve GUIDs: Move, Rotate, Scale, Copy
- Operations that create new GUIDs: Boolean operations, Explode/Join, Rebuild surface

### Key Management
- Keys are case-sensitive
- Empty string values are stored as empty strings (not null)
- Deleting a key sets it to `None` in Rhino

## Troubleshooting

### "openpyxl library is not available"
- Install openpyxl using the Package Manager in Rhino's Script Editor
- Or use command line: `python -m pip install openpyxl`

### "Invalid Excel format. First column must be 'GUID'"
- Ensure the first column header is exactly "GUID" (case-sensitive)
- Don't modify the column order from exported files

### "GUIDs not found"
- Objects may have been deleted from the document
- Ensure you're working in the correct Rhino document
- Check if objects were recreated (new GUIDs assigned)

### Import doesn't update values
- Check that "Create missing keys" is enabled if keys don't exist
- Verify "Update values when Excel cell is empty" settings
- Ensure Excel values are not formatted as formulas

## Version History

- **v1.0** (2025-02-04): Initial release
  - Export to timestamped Excel files
  - Import with GUID tracking
  - Flexible import options
  - Automatic backup creation
  - Comprehensive error reporting

## Author

Aksel
