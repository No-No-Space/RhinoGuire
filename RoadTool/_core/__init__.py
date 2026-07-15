#! python3
# -*- coding: utf-8 -*-
"""Pure geometry + state engine behind Trocha (Road-on-Terrain).

No Eto imports (keeps this testable/reusable). Terrain sampling is not
reimplemented here — Trocha wraps terrain via ``TerrainTools._core.terrain
.TerrainModel`` directly (see RoadTool/road_tool_plan.md rev.3, S1/S4/S6).

Submodules:
  config    - defaults, RG_ tag keys (pure stdlib)
  geometry  - build_slab(): centerline + TerrainModel -> slab Brep (RhinoCommon,
              no RhinoDoc access)
  state     - RG_ROAD_* user-string read/write, create/update/self-heal,
              proximity re-link fallback (touches RhinoDoc via rhinoscriptsyntax)
  junctions - Boolean-union merge of road slabs (RhinoCommon, no RhinoDoc access)
"""

__version__ = "0.1.0"
