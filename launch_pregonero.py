#! python3
"""RhinoGuire - Launch Pregonero (Object Tagger)
Button macro: ! _-RunPythonScript "D:/path/to/RhinoGuire/launch_pregonero.py"
"""
import sys, os, importlib
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import launch as _launch
importlib.reload(_launch)  # pick up launch.py edits without restarting Rhino
_launch.launch("Pregonero")
