#! python3
# -*- coding: utf-8 -*-
"""Headless tests for AreaMeasurer/_paths.py (pure Python, no Rhino needed).

Run with plain CPython:

    python AreaMeasurer/tests/test_paths.py

Covers the v2 layer-structure contract from AreaMeasurer/PLAN.md §2:
classification at depths 1/2/3, prefix cascade, partition on a realistic
tree, layer-depth segments for S4, and level display ordering.
"""

import os
import sys

# Make the repo root importable so `AreaMeasurer._paths` resolves.
_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.normpath(os.path.join(_HERE, "..", ".."))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from AreaMeasurer import _paths as p


_failures = []


def check(name, cond):
    status = "ok  " if cond else "FAIL"
    print("[%s] %s" % (status, name))
    if not cond:
        _failures.append(name)


PARENT = "07-PROPOSAL-Phase_03"

# Realistic fixture mirroring Aksel's model (2026-07), plus noise.
LAYERS = [
    "Forma layers",
    "99-DOCUMENTATION",
    "98-DRAFTING AID",
    "98-DRAFTING AID::ClippingPlanes-WIP",
    PARENT,
    PARENT + "::_BuildingVolumes",
    PARENT + "::_BuildingVolumes::Massing",           # cascade: ignored too
    PARENT + "::Ebene_Parkhaus",
    PARENT + "::Ebene_06",
    PARENT + "::Ebene_06::_TEXT",
    PARENT + "::Ebene_06::_Freifläche",
    PARENT + "::Ebene_06::Diagnostik und Therapie",
    PARENT + "::Ebene_06::Pflege",
    PARENT + "::Ebene_06::Ver- und Entsorgung",
    PARENT + "::Ebene_06::Pflege::Station A",         # depth 3 nesting
    PARENT + "::Ebene_-01",
    PARENT + "::Ebene_-01::Technische Gebäudeausrüstung",
    PARENT + "::Ebene_KMA",
    PARENT + "::Ebene_KMA::Verkehrserschließung",
    "06-PROPOSAL-Phase_02",
    "06-PROPOSAL-Phase_02::Ebene_01",
]


# ── rel_segments ────────────────────────────────────────────────────────────

def test_rel_segments():
    check("parent itself -> []", p.rel_segments(PARENT, PARENT) == [])
    check("outside -> None", p.rel_segments(PARENT, "99-DOCUMENTATION") is None)
    check("sibling prefix does not match",
          p.rel_segments("07-PROPOSAL", PARENT) is None)
    check("depth 1", p.rel_segments(PARENT, PARENT + "::Ebene_06") == ["Ebene_06"])
    check("depth 3",
          p.rel_segments(PARENT, PARENT + "::Ebene_06::Pflege::Station A")
          == ["Ebene_06", "Pflege", "Station A"])


# ── classify ────────────────────────────────────────────────────────────────

def test_classify_kinds():
    check("outside", p.classify(PARENT, "Forma layers")[0] == p.OUTSIDE)
    check("parent", p.classify(PARENT, PARENT)[0] == p.PARENT)
    check("level",
          p.classify(PARENT, PARENT + "::Ebene_06")
          == (p.LEVEL, "Ebene_06", p.PLACEHOLDER))
    check("member depth 2",
          p.classify(PARENT, PARENT + "::Ebene_06::Pflege")
          == (p.MEMBER, "Ebene_06", "Pflege"))
    check("member depth 3 -> own layer is category",
          p.classify(PARENT, PARENT + "::Ebene_06::Pflege::Station A")
          == (p.MEMBER, "Ebene_06", "Station A"))
    check("other proposal is outside",
          p.classify(PARENT, "06-PROPOSAL-Phase_02::Ebene_01")[0] == p.OUTSIDE)


def test_classify_ignore_cascade():
    check("prefixed level ignored",
          p.classify(PARENT, PARENT + "::_BuildingVolumes")
          == (p.IGNORED, None, None))
    check("cascade: child of prefixed level ignored",
          p.classify(PARENT, PARENT + "::_BuildingVolumes::Massing")
          == (p.IGNORED, None, None))
    check("prefixed category ignored, level reported",
          p.classify(PARENT, PARENT + "::Ebene_06::_TEXT")
          == (p.IGNORED, "Ebene_06", None))
    check("umlaut name with prefix",
          p.classify(PARENT, PARENT + "::Ebene_06::_Freifläche")[0] == p.IGNORED)
    check("prefix mid-name is NOT ignored",
          p.classify(PARENT, PARENT + "::Ebene_06")[0] == p.LEVEL)  # Ebene_06 contains _
    check("empty prefix disables ignoring",
          p.classify(PARENT, PARENT + "::_TEXT", prefix="")[0] == p.LEVEL)
    check("custom prefix",
          p.classify(PARENT, PARENT + "::xxIgnore", prefix="xx")[0] == p.IGNORED)
    check("parent itself never prefix-tested",
          p.classify("_07-HIDDEN", "_07-HIDDEN::Ebene_01")[0] == p.LEVEL)


# ── partition / build_tree ──────────────────────────────────────────────────

def test_partition():
    included, ignored = p.partition(PARENT, LAYERS)

    inc_paths = [t[0] for t in included]
    check("parent not included", PARENT not in inc_paths)
    check("outside layers not included",
          all(not t[0].startswith("99-") and not t[0].startswith("06-")
              for t in included))
    check("ignored list exact",
          ignored == sorted([
              PARENT + "::_BuildingVolumes",
              PARENT + "::_BuildingVolumes::Massing",
              PARENT + "::Ebene_06::_TEXT",
              PARENT + "::Ebene_06::_Freifläche",
          ]))
    check("level layers included with placeholder",
          (PARENT + "::Ebene_06", "Ebene_06", p.PLACEHOLDER) in included)
    check("member included",
          (PARENT + "::Ebene_06::Pflege", "Ebene_06", "Pflege") in included)
    check("Parkhaus kept when not prefixed",
          (PARENT + "::Ebene_Parkhaus", "Ebene_Parkhaus", p.PLACEHOLDER) in included)

    tree = p.build_tree(included)
    check("tree has all levels",
          sorted(tree.keys()) == sorted(
              ["Ebene_06", "Ebene_-01", "Ebene_KMA", "Ebene_Parkhaus"]))
    check("tree Ebene_06 categories",
          sorted(tree["Ebene_06"].keys()) == sorted(
              [p.PLACEHOLDER, "Diagnostik und Therapie", "Pflege",
               "Ver- und Entsorgung", "Station A"]))
    check("category maps to its layer path",
          tree["Ebene_-01"]["Technische Gebäudeausrüstung"]
          == [PARENT + "::Ebene_-01::Technische Gebäudeausrüstung"])


# ── segment_at_depth (S4 layer dimension) ───────────────────────────────────

def test_segment_at_depth():
    lp = PARENT + "::Ebene_06::Pflege::Station A"
    check("depth 1 = level", p.segment_at_depth(PARENT, lp, 1) == "Ebene_06")
    check("depth 2", p.segment_at_depth(PARENT, lp, 2) == "Pflege")
    check("depth 3", p.segment_at_depth(PARENT, lp, 3) == "Station A")
    check("too deep -> placeholder",
          p.segment_at_depth(PARENT, lp, 4) == p.PLACEHOLDER)
    check("depth 0 -> placeholder",
          p.segment_at_depth(PARENT, lp, 0) == p.PLACEHOLDER)
    check("outside -> placeholder",
          p.segment_at_depth(PARENT, "99-DOCUMENTATION", 1) == p.PLACEHOLDER)
    check("parent itself -> placeholder",
          p.segment_at_depth(PARENT, PARENT, 1) == p.PLACEHOLDER)


# ── level_sort_key ──────────────────────────────────────────────────────────

def test_level_sort_key():
    names = ["Ebene_-03", "Ebene_KMA", "Ebene_06", "Ebene_00-Eingang",
             "Ebene_-01", "Ebene-Logistik Zentrum", "Ebene_05"]
    ordered = sorted(names, key=p.level_sort_key)
    check("numeric levels descending, textual after",
          ordered == ["Ebene_06", "Ebene_05", "Ebene_00-Eingang",
                      "Ebene_-01", "Ebene_-03",
                      "Ebene-Logistik Zentrum", "Ebene_KMA"])
    check("None tolerated", p.level_sort_key(None) == (2, 0, ""))


# ── run ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    test_rel_segments()
    test_classify_kinds()
    test_classify_ignore_cascade()
    test_partition()
    test_segment_at_depth()
    test_level_sort_key()
    print()
    if _failures:
        print("%d FAILURE(S): %s" % (len(_failures), ", ".join(_failures)))
        sys.exit(1)
    print("All tests passed.")
