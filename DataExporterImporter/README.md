# Arriero — Data Exporter/Importer

Export and import object metadata (User Keys/Values) between Rhino 8 and Excel files using GUID-based tracking.

## Workflow

### Export

1. Click **"1. Export Data (Select Objects > Excel)"**
2. Select objects in Rhino that have User Keys/Values
3. Choose a save location (default filename includes timestamp)

Excel file is saved with all object GUIDs and their key/value pairs (keys sorted alphabetically).

### Import

1. Configure import options in the window (see below)
2. Click **"2. Import Data (Excel > Update Objects)"**
3. Optionally create a backup when prompted
4. Select the Excel file to import
5. Review the summary report (objects updated, keys created/updated/deleted)

## Import Options

### Create backup before import

**Default:** Checked. Creates a timestamped backup export of all current object data before making changes.

### Create missing keys from Excel columns

**Default:** Checked. If Excel has columns (keys) that don't exist on objects, they will be created. If unchecked, only existing keys are updated.

**Scope options:**

- *Apply to all objects in Excel file* (default)
- *Apply only to pre-selected objects* — limits new key creation to objects selected before running the script

### Update values when Excel cell is empty

**Default:** Unchecked. Empty Excel cells are ignored and original Rhino values are preserved.

When checked, empty cells trigger special behavior based on the placeholder setting:

- **With placeholder** (e.g., `-`): empty cells set the key value to the placeholder string
- **Without placeholder** (blank field): empty cells **delete the key entirely**

## Excel File Format

First column must be named `GUID` (case-sensitive). Additional columns are User Keys.

```text
| GUID                                 | Key1 | Key2 | ...
|--------------------------------------|------|------|----
| 12345678-90ab-cdef-1234-567890abcdef | val1 | val2 | ...
```

## Technical Notes

- GUIDs are persistent across sessions but are replaced by: Boolean operations, Explode/Join, Rebuild surface
- Keys are case-sensitive
- Deleting a key sets it to `None` in Rhino

## Troubleshooting

**"Invalid Excel format. First column must be 'GUID'"** — don't modify the column order from exported files.

**"GUIDs not found"** — objects may have been deleted or recreated (new GUIDs assigned).

**Import doesn't update values** — check that "Create missing keys" is enabled; verify empty-cell settings; ensure Excel values aren't formatted as formulas.
