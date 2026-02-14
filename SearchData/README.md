# Baquiano - Search Data Tool

Advanced search tool for Rhino 8 object metadata (User Keys/Values) with include/exclude conditions and pre-selection filtering.

## Features

- **Search by User Keys** - find objects based on their metadata key/value pairs
- **Include/Exclude conditions** - build complex queries with AND/OR logic
- **Pre-selection support** - search within previously selected objects or the entire model
- **8 match types** - Contains, Equals, Starts with, Ends with, and their negations
- **Dynamic conditions** - add or remove conditions at runtime
- **Results selection** - matching objects are automatically selected in the viewport

## Requirements

- **Rhino 8** with CPython 3
- No external libraries required

## How to Use

### Running the Script

1. Optionally pre-select objects in Rhino to limit the search scope
2. Type `RunPythonScript` in the command line
3. Navigate to `Baquiano.py` and click **Open**
4. The search window opens

### Search Scope

- **Search all objects in model** - searches every object in the document
- **Search only pre-selected objects** - only available if objects were selected before running the script; shows the count of pre-selected objects

### Building Search Conditions

#### Include Conditions (AND logic)

Objects must match **ALL** include conditions to be included in results.

1. Type the **Key** name (e.g., `Name`)
2. Type the **Value** to search for (e.g., `House`)
3. Select a **Match Type** from the dropdown
4. Click **"+ Add Include Condition"** to add more conditions

#### Exclude Conditions (OR logic)

Objects matching **ANY** exclude condition are removed from results.

1. Click **"+ Add Exclude Condition"**
2. Type the Key, Value, and select a Match Type
3. Any object matching this condition will be excluded even if it passes the include filters

### Match Types

| Match Type | Description |
| --- | --- |
| Contains | Value appears anywhere in the key's value |
| Equals | Exact match (case-insensitive) |
| Starts with | Key's value begins with the search value |
| Ends with | Key's value ends with the search value |
| Does not contain | Value does NOT appear in the key's value |
| Does not equal | NOT an exact match |
| Does not start with | Key's value does NOT begin with the search value |
| Does not end with | Key's value does NOT end with the search value |

### Example: Cross-Search

Find all objects where Key `Name` contains `House`, but exclude objects where Key `User` equals `university`:

1. **Include condition**: Key = `Name`, Value = `House`, Match = `Contains`
2. **Exclude condition**: Key = `User`, Value = `university`, Match = `Equals`
3. Click **Search**

### Example: Finding Outliers

Find objects where Key `Status` does NOT equal `Approved`:

1. **Include condition**: Key = `Status`, Value = `Approved`, Match = `Does not equal`
2. Click **Search** - this selects all objects with a `Status` value that is not `Approved`

## Search Results

After clicking **Search**:

- Matching objects are selected in the Rhino viewport
- A summary message shows:
  - Number of objects found
  - Include and exclude conditions used
  - Search scope and total objects searched
- If no objects match, a "No objects found" message is shown
- Cancelling restores the original selection

## Troubleshooting

### No results found

- Check that the Key name matches exactly (case-sensitive)
- Try using "Contains" instead of "Equals" for a broader match
- Verify that objects have the expected User Keys (use `GetUserText` command in Rhino)

### "Please specify at least one include condition"

- At least one include condition with both Key and Value filled in is required

## Version History

- **v0.5** (2026-02-14): Unified release
  - Eto.Forms GUI with dynamic condition rows
  - Include/exclude conditions with AND/OR logic
  - 8 match types including negations
  - Pre-selection filtering
  - CPython 3 compatible

## Author

Aquelon - aquelon@pm.me
