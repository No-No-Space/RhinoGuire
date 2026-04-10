"""
RhinoGuire - Installer / Path Setup
Run this script once from Rhino Python editor to register the script paths.

Example (adjust to your local path):
  _RunPythonScript "<path-to-repo>/RhinoGuire/install.py"
"""

import rhinoscriptsyntax as rs
import scriptcontext as sc
import os
import sys

# --- CONFIG ---
# Change this path if the repo is cloned to a different location
RHINOGUIRE_ROOT = os.path.dirname(os.path.abspath(__file__))

SCRIPTS = {
    "RG_Lindero":   os.path.join(RHINOGUIRE_ROOT, "AreaMeasurer",        "Lindero.py"),
    "RG_Arriero":   os.path.join(RHINOGUIRE_ROOT, "DataExporterImporter", "Arriero.py"),
    "RG_Chivito":   os.path.join(RHINOGUIRE_ROOT, "DataVisualization",   "Chivito.py"),
    "RG_Sebucan":   os.path.join(RHINOGUIRE_ROOT, "MeshTools",           "WrapeMeshOnMesh", "Sebucan.py"),
    "RG_Baquiano":  os.path.join(RHINOGUIRE_ROOT, "SearchData",          "Baquiano.py"),
}

# Verify all scripts exist
missing = [name for name, path in SCRIPTS.items() if not os.path.exists(path)]
if missing:
    rs.MessageBox(
        "RhinoGuire: Missing scripts:\n" + "\n".join(missing) +
        "\n\nCheck that RHINOGUIRE_ROOT is correct:\n" + RHINOGUIRE_ROOT,
        title="RhinoGuire Install Error"
    )
else:
    # Store paths in Rhino sticky so macros can reference them
    for name, path in SCRIPTS.items():
        sc.sticky[name] = path

    rs.MessageBox(
        "RhinoGuire installed successfully!\n\n" +
        "Scripts registered:\n" +
        "\n".join([f"  {k}" for k in SCRIPTS.keys()]) +
        "\n\nYou can now load the toolbar from:\n" +
        "Tools > Toolbar Layout > Open > RhinoGuire/ui/RhinoGuire.rui",
        title="RhinoGuire"
    )
