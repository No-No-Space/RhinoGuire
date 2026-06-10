#! python3
"""RhinoGuire - Launch WayGrader (Way / Path Corridor Grading)
Button macro: ! _-RunPythonScript "D:/path/to/RhinoGuire/launch_waygrader.py"
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from launch import launch
launch("WayGrader")
