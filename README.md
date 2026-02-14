# RhinoGuire

A collection of Python 3 tools for managing object metadata (User Keys/Values) in Rhino 8.

## Tools

### [Arriero](DataExporterImporter/) - Data Exporter/Importer

Export and import object metadata between Rhino and Excel files using GUID-based tracking. Supports backup creation, key creation scope, and flexible handling of empty cells.

### [Chivito](DataVisualization/) - Data Visualization

Color-code Rhino objects based on their metadata values. Three-step workflow: initialize keys from Excel, extract unique values, and visualize with an interactive Color Manager. Includes legend and viewport PNG export.

### [Baquiano](SearchData/) - Search Data

Search and select Rhino objects by their metadata using include/exclude conditions with 8 match types (Contains, Equals, Starts with, Ends with, and their negations). Supports pre-selection filtering and cross-search queries.

## Requirements

- **Rhino 8** with CPython 3
- **openpyxl** - required by Arriero and Chivito (installed automatically via `# r: openpyxl` header)
- Baquiano has no external dependencies

## Quick Start

1. Open Rhino 8
2. Type `RunPythonScript` in the command line
3. Navigate to the desired tool's `.py` file and click **Open**

Each tool opens its own GUI window. See the individual README files for detailed usage instructions.

## Repository Structure

```text
RhinoGuire/
├── README.md
├── LICENSE
├── DataExporterImporter/
│   ├── Arriero.py
│   └── README.md
├── DataVisualization/
│   ├── Chivito.py
│   ├── README.md
│   └── _ExcelOutput/
└── SearchData/
    ├── Baquiano.py
    └── README.md
```

## License

MIT License - see [LICENSE](LICENSE) for details.

## Author

Aquelon - aquelon@pm.me
