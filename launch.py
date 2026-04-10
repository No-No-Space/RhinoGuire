"""
RhinoGuire - Script Launcher
Used by toolbar buttons. Pass the script key as argument.

Button macro example:
  ! _RunPythonScript "D:/path/to/RhinoGuire/launch.py" "RG_Lindero"
"""

import scriptcontext as sc
import rhinoscriptsyntax as rs
import sys
import os

def launch(script_key):
    path = sc.sticky.get(script_key)
    if not path:
        rs.MessageBox(
            f"Script '{script_key}' not found.\n\nRun install.py first:\n"
            "_RunPythonScript path/to/RhinoGuire/install.py",
            title="RhinoGuire"
        )
        return
    if not os.path.exists(path):
        rs.MessageBox(f"Script file not found:\n{path}", title="RhinoGuire")
        return

    exec(open(path).read(), {"__file__": path})

# Called from macro with argument
if __name__ == "__main__":
    if len(sys.argv) > 1:
        launch(sys.argv[1])
