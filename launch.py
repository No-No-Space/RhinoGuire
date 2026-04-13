#! python3
"""
RhinoGuire - Script Launcher
Resolves all script paths relative to this file's location.
No install step required — works out of the box after cloning.

Button macro: ! _-RunPythonScript "D:/path/to/RhinoGuire/launch_lindero.py"
"""
import os
import rhinoscriptsyntax as rs

_ROOT = os.path.dirname(os.path.abspath(__file__))

SCRIPTS = {
    "Lindero":  os.path.join(_ROOT, "AreaMeasurer",        "Lindero.py"),
    "Arriero":  os.path.join(_ROOT, "DataExporterImporter", "Arriero.py"),
    "Chivito":  os.path.join(_ROOT, "DataVisualization",   "Chivito.py"),
    "Sebucan":  os.path.join(_ROOT, "MeshTools",           "WrapMeshOnMesh", "Sebucan.py"),
    "Baquiano": os.path.join(_ROOT, "SearchData",          "Baquiano.py"),
}

def launch(key):
    path = SCRIPTS.get(key)
    if not path:
        rs.MessageBox("Unknown script key: " + key, title="RhinoGuire")
        return
    if not os.path.exists(path):
        rs.MessageBox("Script file not found:\n" + path, title="RhinoGuire")
        return
    exec(open(path, encoding="utf-8").read(), {"__file__": path, "__name__": "__main__"})
