#! python3
"""RhinoGuire - Launch Sebucan (Wrap Mesh on Mesh)
Button macro: ! _-RunPythonScript "D:/path/to/RhinoGuire/launch_sebucan.py"
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from launch import launch
launch("Sebucan")
