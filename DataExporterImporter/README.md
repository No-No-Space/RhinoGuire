# Arriero - Rhino Data Exporter/Importer

Export and import object metadata (User Keys/Values) between Rhino 8 and Excel files using GUID-based tracking.

## Features

- **Export**: Extract User Keys/Values from selected objects to timestamped Excel files
- **Import**: Update objects from Excel with flexible options for missing data
- **GUID Tracking**: Uses Rhino's native GUIDs for reliable object identification
- **Backup**: Optional automatic backup before import operations
- **Key Creation Scope**: Apply new keys to all objects in the Excel file or only to pre-selected objects
- **Flexible Import Options**:
  - Create or skip missing keys
  - Handle empty cells with placeholder values or key deletion
  - Comprehensive error reporting

## Requirements

- **Rhino 8** with CPython 3
- **openpyxl** library for Excel operations

### Installing openpyxl

The script header includes `# r: openpyxl`, which tells Rhino 8 to install the package automatically on first run.

If that does not work, install it manually:

1. Open Rhino 8
2. Go to **Tools > PythonScript > Edit...**
3. In the Script Editor menu: **Tools > Package Manager**
4. Search for `openpyxl` and install it

## How to Use

### Running the Script

1. In Rhino 8, type `RunPythonScript` in the command line
2. Navigate to `Arriero.py` and click **Open**
3. A window opens with Export/Import buttons and import options

### Export Workflow

1. Click **"1. Export Data (Select Objects > Excel)"**
2. Select objects in Rhino that have User Keys/Values
3. Choose a save location (default filename includes timestamp)
4. Excel file is saved with all object GUIDs and their key/value pairs

### Import Workflow

1. Configure import options in the window (see below)
2. Click **"2. Import Data (Excel > Update Objects)"**
3. Optionally create a backup when prompted
4. Select the Excel file to import
5. Review the summary report (objects updated, keys created/updated/deleted)

## Import Options

### Create backup before import
- **Default**: Checked
- Creates a timestamped backup export of all current object data before making changes

### Create missing keys from Excel columns
- **Default**: Checked
- If Excel has columns (keys) that don't exist on objects, they will be created
- If unchecked, only existing keys will be updated
- **Scope options**:
  - *Apply to all objects in Excel file* (default)
  - *Apply only to pre-selected objects* - limits new key creation to objects you selected before running the script

### Update values when Excel cell is empty
- **Default**: Unchecked
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

- **v0.5** (2026-02-14): Unified release
  - Added save/load file location dialogs
  - Added key creation scope (all objects or pre-selected only)
  - CPython 3 compatible
- **v1.0** (2025-02-04): Initial release
  - Export to timestamped Excel files
  - Import with GUID tracking
  - Flexible import options
  - Automatic backup creation

## Author

Aquelon - aquelon@pm.me
