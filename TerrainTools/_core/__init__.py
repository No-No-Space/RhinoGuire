#! python3
# -*- coding: utf-8 -*-
"""TerrainTools shared grading engine (``grading_core``).

This package holds the pure engine behind PadGrader, WayGrader and
CutFillReport. It is deliberately UI-free (no ``Eto.*`` imports) so the
geometry/math can be reused and headless-tested.

Submodules:
  slope     — unit conversions (ratio / percent / degrees <-> canonical m=V/H)
  terrain   — TerrainModel: coerce Surface/Mesh/SubD -> mesh, vertical raycast
  grading   — pad + corridor design-heightfield builders, GradeResult
  volumes   — cut/fill from a GradeResult (grid-prism method); per-station
  meshbuild — grid -> Rhino Mesh; tint by cut/fill depth
  report    — openpyxl workbook writer

Note: ``slope`` and ``volumes`` import only the stdlib (no RhinoCommon), so
they can be imported and tested under plain CPython. The other modules need
RhinoCommon and therefore only import successfully inside Rhino.
"""

__version__ = "0.1.0"
