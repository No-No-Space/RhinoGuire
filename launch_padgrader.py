#! python3
"""RhinoGuire - Launch PadGrader (Building Pad Grading)
Button macro: ! _-RunPythonScript "D:/path/to/RhinoGuire/launch_padgrader.py"
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from launch import launch
launch("PadGrader")
