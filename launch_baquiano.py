#! python3
"""RhinoGuire - Launch Baquiano (Search Data)
Button macro: ! _-RunPythonScript "D:/path/to/RhinoGuire/launch_baquiano.py"
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from launch import launch
launch("Baquiano")
