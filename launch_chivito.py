#! python3
"""RhinoGuire - Launch Chivito (Data Visualization)
Button macro: ! _-RunPythonScript "D:/path/to/RhinoGuire/launch_chivito.py"
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from launch import launch
launch("Chivito")
