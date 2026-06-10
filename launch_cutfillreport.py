#! python3
"""RhinoGuire - Launch CutFillReport (Compare / Quantify / Export)
Button macro: ! _-RunPythonScript "D:/path/to/RhinoGuire/launch_cutfillreport.py"
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from launch import launch
launch("CutFillReport")
