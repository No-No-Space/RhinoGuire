#! python3
"""RhinoGuire - Launch Trocha (Road on Terrain)
Button macro: ! _-RunPythonScript "D:/path/to/RhinoGuire/launch_trocha.py"
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from launch import launch
launch("Trocha")
