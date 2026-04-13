#! python3
"""RhinoGuire - Launch Arriero (Data Exporter/Importer)
Button macro: ! _-RunPythonScript "D:/path/to/RhinoGuire/launch_arriero.py"
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from launch import launch
launch("Arriero")
