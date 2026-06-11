# Chivito — Data Visualization

Color-code Rhino objects based on their metadata values. Three-step workflow: initialize keys, extract unique values, then visualize with an interactive Color Manager.

## Workflow

### Step 1 — Initialize Keys

Apply metadata keys from an Excel template to selected objects.

1. Click **"Select Excel & Apply Keys"**
2. Select an Excel file with key definitions (columns: Key Name, Default Value)
3. Select objects in Rhino — keys are added without overwriting existing values

Key Definitions format:

```text
| Key Name      | Default Value |
|---------------|---------------|
| Building Year |               |
| Condition     | Good          |
```

### Step 2 — Extract Unique Values

Scan objects and generate a color mapping Excel file.

1. Click **"Scan & Generate Excel"**
2. Select objects with metadata
3. Choose a save location (default: `_ExcelOutput/UniqueValuesColorSettings.xlsx`)

If the file already exists, existing color definitions are preserved (smart merge). After export, open the file and fill in the R, G, B, A columns for each value.

Color Mapping format:

```text
| Key           | Value | R  | G   | B  | A   |
|---------------|-------|----|-----|----|-----|
| Building Year | 1950  | 80 | 80  | 80 | 255 |
| Condition     | Good  | 0  | 200 | 0  | 255 |
```

A values: 0–255 (defaults to 255 = fully opaque).

### Step 3 — Visualize with Colors

1. Click **"Open Color Manager"**
2. Select objects to visualize, then select the color mapping Excel file
3. In the Color Manager (non-blocking):
   - Select a key from the dropdown, then click **Update Colors**
   - Change the default color for objects without a matching value
   - Review warnings for missing definitions or unused entries
   - **"Select Problem Objects"** to highlight issues in the viewport
   - **"Export Legend as PNG"** / **"Capture Viewport as PNG"**

## Troubleshooting

**Objects not coloring** — verify RGB values are filled in the Excel file; check the Warnings panel; use "Select Problem Objects".

**"No keys found in Excel file"** — Key Name must be in column A with data starting from row 2.
