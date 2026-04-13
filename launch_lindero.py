#! python3
"""RhinoGuire - Launch Lindero (Area Measurer)
Button macro: ! _-RunPythonScript "D:/path/to/RhinoGuire/launch_lindero.py"
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from launch import launch
launch("Lindero")
