#! python3
# r: openpyxl
# -*- coding: utf-8 -*-
# __title__ = "Lindero"
# __doc__ = """Version = 0.5
# Date    = 2026-05-07
# Author: Aquelon - aquelon@pm.me
# _____________________________________________________________________
# Description:
# Footprint area calculator for Rhino objects.
# "Footprint" = the plan area (XY projection), NOT the sum of all surfaces.
# Modeless window — Rhino stays accessible while it is open.
# Accepted geometry: solids, extrusions, closed planar curves, planar
# surfaces, and hatches.
# _____________________________________________________________________
# Scenarios (v2 layer-structure contract — see AreaMeasurer/PLAN.md):
#   S1 — Selected Objects:
#        Individual footprint per object. Overlapping footprints are
#        merged with a Boolean Union to avoid double-counting (same as S2).
#
#   S2 — By Layer:
#        All objects on a layer. Overlapping footprints merged with a
#        Boolean Union to avoid double-counting.
#
#   S3 — Layer Hierarchy:
#        Parent layer → direct children = LEVELS (floors) → the object's
#        own (lowest) layer = CATEGORY (e.g. DIN 13080 areas). Objects are
#        collected from ALL descendant layers; layers prefixed with the
#        ignore prefix (Settings, default "_") are excluded with their
#        whole subtree. Union per (level, category); level total = union
#        across the level (cross-category overlap warned); grand total =
#        sum of level totals. Replaces the old key-based S3 and old S5.
#
#   S4 — Custom Aggregation:
#        User-defined hierarchy of dimensions; each dimension is either a
#        'Layer @ depth N' path segment or a user text key. Footprints are
#        merged per leaf group per level, then summed across levels.
#        Results shown as an indented tree. The old two-key S3 = a config
#        with two UserText dimensions.
#
#   R1 / R2 — Analysis:
#        One shared engine: aggregates merged areas by a chosen dimension
#        (Category / Level / UserText key) across all levels and compares
#        against the target table in Settings. Bullet chart per entry.
# _____________________________________________________________________
# Last update:
# - [02.07.2026] - 0.8.3 Copy Window fixed on multi-monitor setups: window
#                     rect now read from Win32 in physical pixels (ctypes,
#                     DwmGetWindowAttribute → GetWindowRect fallback) instead
#                     of scaling Eto's logical Bounds — logical coordinates
#                     shift per monitor when scale factors differ
# - [02.07.2026] - 0.8.2 Settings layout rebuilt on the proven DynamicLayout
#                     patterns (\n-hints in column 2 — wrapping labels blow
#                     up the Scrollable canvas width, see _tab_settings note);
#                     Write Area works from S4 (per-object areas returned by
#                     calc_s4); "Copy Window" button — screenshots the window
#                     to the clipboard (DPI-aware CopyFromScreen)
# - [02.07.2026] - 0.8.1 UI: full-width descriptions — S3/Settings text no
#                     longer squeezed into the first layout column; false
#                     [union failed] fixed when a group contains only
#                     unmeasurable objects (text/annotations); settings
#                     auto-persist per model (3dm document user text,
#                     restored at startup, replaced when a config is loaded)
# - [02.07.2026] - 0.8 v2 rebuild for the new layer structure: levels and
#                     DIN 13080 categories read from the layer tree
#                     (parent → level → … → category = object's own layer);
#                     "_" ignore prefix with subtree cascade + dry-run
#                     preview; footprint cache per calculation run; S5
#                     folded into S3 (Z-banding everywhere); S4 dimensions
#                     generalized to Layer@depth | UserText; R1/R2 targets
#                     from Settings table; config v2 (v1 still loads);
#                     pure path logic split into _paths.py with headless
#                     tests (AreaMeasurer/tests/test_paths.py)
# - [01.07.2026] - 0.7 Arrangement conservation check (regions covered by an
#                     object must reproduce its own area, else the whole
#                     arrangement is rejected — coincident tile edges corrupt
#                     CreateBooleanRegions even at tight tolerance); new
#                     pairwise inclusion–exclusion merge as second path
# - [01.07.2026] - 0.6 Hole-aware footprints: inner loops (courtyards/shafts)
#                     subtracted per object; overlap merge rebuilt on
#                     Curve.CreateBooleanRegions arrangement + point
#                     classification (exact for overlaps, holes, and ring
#                     layouts); 2D boolean tolerance clamped to ~1 mm
#                     (coarse doc tolerances corrupt the arrangement);
#                     merged total bounded by sum of individual regions;
#                     version + file timestamp shown in UI;
#                     Write Area honors Settings decimals
# - [07.05.2026] - 0.5 Overlap removal for S1; closed curves, planar surfaces,
#                     and hatches supported; S3 Group Key first, split result
#                     panels; configurable decimal places in Settings;
#                     colored grid results for S1–S4
# - [10.04.2026] - 0.4 S4 Custom Aggregation, R1/R2 (renamed from S4/S5),
#                     R1/R2 data source toggle (S3 keys or S4 hierarchy)
# - [26.03.2026] - 0.3 Settings tab, S4/S5 bullet-chart analysis, Write Area,
#                     per-tab results areas, overlap warnings
# - [20.02.2026] - 0.2 Export results to Excel (Objects + Summary sheets)
# - [20.02.2026] - 0.1 Initial release
# _____________________________________________________________________


import rhinoscriptsyntax as rs
import Rhino
import Rhino.Geometry as rg
import scriptcontext as sc
import Eto.Drawing as drawing
import Eto.Forms as forms
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import json
import System

import sys as _sys, os as _os
_rg_root = _os.path.normpath(_os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))
if _rg_root not in _sys.path:
    _sys.path.insert(0, _rg_root)
from ui import theme as _t
import importlib as _importlib; _importlib.reload(_t)
from AreaMeasurer import _paths as _lp
_importlib.reload(_lp)

# Version shown in the UI. The file timestamp is added so ANY edit to this
# file is visible in the running window — no stale-instance guessing.
__version__ = "0.8.3"
try:
    import datetime as _dt
    _BUILD_STAMP = _dt.datetime.fromtimestamp(
        _os.path.getmtime(_os.path.abspath(__file__))
    ).strftime("%Y-%m-%d %H:%M")
except Exception:
    _BUILD_STAMP = "unknown"
_VERSION_TEXT = f"v{__version__}  ·  file {_BUILD_STAMP}"

# ============================================================================
# PERSISTENT PREFERENCES (last-used folder per action)
# ============================================================================

_prefs_get = _t.prefs_get
_prefs_set = _t.prefs_set

# Per-document settings persistence: the current config (v2 dict) is stored
# as document user text under this key — it travels inside the 3dm.
DOC_CONFIG_KEY = "Lindero.config_v2"


# ══════════════════════════════════════════════════════════════════════════════
# Model helpers
# ══════════════════════════════════════════════════════════════════════════════

def get_all_user_text_keys():
    """Return sorted list of unique user text keys found on any object in the model."""
    keys = set()
    doc = Rhino.RhinoDoc.ActiveDoc
    if doc:
        for rhino_obj in doc.Objects:
            if rhino_obj and rhino_obj.Attributes:
                user_strings = rhino_obj.Attributes.GetUserStrings()
                if user_strings:
                    for key in user_strings.AllKeys:
                        keys.add(key)
    return sorted(list(keys))


def all_layer_names():
    """Return sorted list of all layer full path names in the model."""
    return sorted(rs.LayerNames() or [])


def short_name(full_path):
    """Return the last segment of a layer full path (last part after '::')."""
    return full_path.split("::")[-1] if "::" in full_path else full_path


def get_layer_objects(layer_name):
    """Return GUIDs of objects directly on the given layer."""
    return rs.ObjectsByLayer(layer_name) or []


def get_child_layers(parent_layer_name):
    """Return full path names of direct child layers of the given layer."""
    prefix = parent_layer_name + "::"
    return sorted(
        name for name in (rs.LayerNames() or [])
        if name.startswith(prefix) and "::" not in name[len(prefix):]
    )


def unit_label():
    """Return the squared-unit label matching the current Rhino model unit system."""
    mapping = {
        "Millimeters": "mm²", "Centimeters": "cm²", "Meters": "m²",
        "Kilometers": "km²", "Feet": "ft²", "Inches": "in²",
    }
    return mapping.get(sc.doc.ModelUnitSystem.ToString(), "units²")


def collect_objects(parent_layer, ignore_prefix=_lp.DEFAULT_IGNORE_PREFIX):
    """
    Collect measurable objects under a parent layer per the v2 contract
    (_paths.py): level = direct child of the parent, category = the
    object's own (lowest) layer; prefixed layers excluded with their
    whole subtree.

    Returns (records, layer_info):
      records:    [{guid, level, category, layer}]  — guid is the raw Guid
      layer_info: {"included": [(layer_path, level, category)],
                   "ignored": [layer_path],
                   "parent_direct_objects": int}
    """
    all_layers = rs.LayerNames() or []
    included, ignored = _lp.partition(parent_layer, all_layers, ignore_prefix)
    records = []
    for layer_path, level, category in included:
        for g in get_layer_objects(layer_path):
            records.append({
                "guid": g, "level": level,
                "category": category, "layer": layer_path,
            })
    return records, {
        "included": included,
        "ignored": ignored,
        "parent_direct_objects": len(get_layer_objects(parent_layer)),
    }


def preview_hierarchy(parent_layer, ignore_prefix=_lp.DEFAULT_IGNORE_PREFIX):
    """
    Dry run for S3: which layers count as what, with object counts —
    NO footprint computation, so it is fast even on heavy models.

    Returns {"tree": {level: {category: n_objects}}, "ignored": [path],
             "parent_direct_objects": int, "total_objects": int}.
    """
    all_layers = rs.LayerNames() or []
    included, ignored = _lp.partition(parent_layer, all_layers, ignore_prefix)
    tree, total = {}, 0
    for layer_path, level, category in included:
        n = len(get_layer_objects(layer_path))
        cats = tree.setdefault(level, {})
        cats[category] = cats.get(category, 0) + n
        total += n
    return {
        "tree": tree,
        "ignored": ignored,
        "parent_direct_objects": len(get_layer_objects(parent_layer)),
        "total_objects": total,
    }


# ══════════════════════════════════════════════════════════════════════════════
# Footprint geometry
# ══════════════════════════════════════════════════════════════════════════════

def _bbox_footprint(obj_guid):
    """Bounding-box fallback: XY rectangle from the object's bounding box."""
    rhobj = sc.doc.Objects.FindId(obj_guid)
    if rhobj is None:
        return None
    bbox = rhobj.GetBoundingBox(True)
    if not bbox.IsValid:
        return None
    pts = [
        rg.Point3d(bbox.Min.X, bbox.Min.Y, 0),
        rg.Point3d(bbox.Max.X, bbox.Min.Y, 0),
        rg.Point3d(bbox.Max.X, bbox.Max.Y, 0),
        rg.Point3d(bbox.Min.X, bbox.Max.Y, 0),
        rg.Point3d(bbox.Min.X, bbox.Min.Y, 0),
    ]
    return rg.PolylineCurve(pts)


def _brep_footprint_curves(brep):
    """
    Find the bottom horizontal face(s) of a Brep and return their outer border
    curves projected to Z=0.
    Falls back to the XY bounding box if no horizontal faces are found.
    """
    tol = sc.doc.ModelAbsoluteTolerance
    proj = rg.Transform.PlanarProjection(rg.Plane.WorldXY)

    horiz = []
    for face in brep.Faces:
        u = (face.Domain(0).Min + face.Domain(0).Max) * 0.5
        v = (face.Domain(1).Min + face.Domain(1).Max) * 0.5
        normal = face.NormalAt(u, v)
        if abs(normal.Z) > 0.9:
            pt = face.PointAt(u, v)
            horiz.append((face, pt.Z))

    if not horiz:
        bbox = brep.GetBoundingBox(True)
        if not bbox.IsValid:
            return []
        pts = [
            rg.Point3d(bbox.Min.X, bbox.Min.Y, 0),
            rg.Point3d(bbox.Max.X, bbox.Min.Y, 0),
            rg.Point3d(bbox.Max.X, bbox.Max.Y, 0),
            rg.Point3d(bbox.Min.X, bbox.Max.Y, 0),
            rg.Point3d(bbox.Min.X, bbox.Min.Y, 0),
        ]
        return [rg.PolylineCurve(pts)]

    min_z = min(z for _, z in horiz)
    # Cap the "same level" test: with a coarse doc tolerance (0.1 m) tol*2
    # could reach 0.2 and pull in the TOP face of a thin slab as well.
    ztol = min(tol * 2.0, _bool_tol() * 20.0)
    bottom = [f for f, z in horiz if abs(z - min_z) <= ztol]

    curves = []
    for face in bottom:
        # All loops: outer boundary AND inner (hole) loops. Holes matter —
        # a floor plate with a courtyard opening must not count the void.
        # Nesting is resolved later by _region_area / CreatePlanarBreps.
        loops = []
        try:
            loops = [lp.To3dCurve() for lp in face.Loops]
            loops = [c for c in loops if c is not None]
        except Exception:
            loops = []
        # Fallback for untrimmed/planar faces where loop extraction fails
        if not loops:
            face_copy = face.DuplicateFace(False)
            if face_copy:
                edge_curves = face_copy.DuplicateEdgeCurves(True)
                if edge_curves:
                    joined = rg.Curve.JoinCurves(edge_curves, tol)
                    loops = [j for j in joined if j.IsClosed]
        for border in loops:
            dup = border.DuplicateCurve()
            dup.Transform(proj)
            if dup.IsClosed:
                curves.append(dup)
    return curves


def get_footprint_curves(obj_guid):
    """
    Return a list of planar closed curves at Z=0 representing the footprint
    of the object. Returns an empty list if the footprint cannot be determined
    or if the geometry type is unsupported / raises an exception.
    """
    try:
        rhobj = sc.doc.Objects.FindId(obj_guid)
        if rhobj is None:
            return []
        geom = rhobj.Geometry

        if isinstance(geom, (rg.Extrusion, rg.Surface)):
            brep = geom.ToBrep()
            if brep:
                return _brep_footprint_curves(brep)

        if isinstance(geom, rg.Brep):
            return _brep_footprint_curves(geom)

        if isinstance(geom, rg.Hatch):
            proj = rg.Transform.PlanarProjection(rg.Plane.WorldXY)
            loops = list(geom.Get3dCurves(True) or [])   # outer loops
            loops += list(geom.Get3dCurves(False) or [])  # inner (hole) loops
            if loops:
                result = []
                for loop in loops:
                    if loop.IsClosed:
                        dup = loop.DuplicateCurve()
                        dup.Transform(proj)
                        result.append(dup)
                if result:
                    return result

        if isinstance(geom, rg.Curve) and geom.IsClosed and geom.IsPlanar():
            proj = rg.Transform.PlanarProjection(rg.Plane.WorldXY)
            dup = geom.DuplicateCurve()
            dup.Transform(proj)
            return [dup]

        fallback = _bbox_footprint(obj_guid)
        return [fallback] if fallback else []
    except Exception:
        return []


def _fp(guid, cache):
    """
    Footprint curves for guid, memoized in cache.

    One dict per calculation run: the same object's projection feeds the
    per-object listing, the per-category union, AND the per-level union —
    without the cache every union call recomputes it (S3 on levels ×
    categories multiplies that). Pass cache=None to bypass.
    """
    if cache is None:
        return get_footprint_curves(guid)
    key = str(guid)
    if key not in cache:
        cache[key] = get_footprint_curves(guid)
    return cache[key]


def curve_area(curve):
    """Return the area enclosed by a planar closed curve, or 0 on failure."""
    try:
        amp = rg.AreaMassProperties.Compute(curve)
        return amp.Area if amp else 0.0
    except Exception:
        return 0.0


def _bool_tol():
    """
    Tolerance for the 2D region math (arrangement, planar breps, coverage),
    clamped to ~1 mm in document units.

    Do NOT pass sc.doc.ModelAbsoluteTolerance straight into these calls: a
    coarse document tolerance (e.g. 0.1 in a meters file) makes Rhino's
    boolean/arrangement operations weld distinct edges and return corrupt
    regions — diagnosed 01.07.2026 with debug_merge.py (RegionCount and
    region areas were garbage at tol=0.1, correct at 0.001).
    """
    tol = sc.doc.ModelAbsoluteTolerance
    try:
        mm = Rhino.RhinoMath.UnitScale(
            Rhino.UnitSystem.Millimeters, sc.doc.ModelUnitSystem
        )
        if mm and mm > 0:
            return min(tol, mm)
    except Exception:
        pass
    return min(tol, 0.001)


def get_footprint_area(obj_guid):
    """Footprint area of a single object (bottom faces, holes subtracted)."""
    return _region_area(get_footprint_curves(obj_guid))


def get_object_bottom_z(guid):
    """Return the lowest horizontal face Z for the object, or bbox min Z as fallback."""
    try:
        rhobj = sc.doc.Objects.FindId(guid)
        if rhobj is None:
            return None
        geom = rhobj.Geometry
        brep = None
        if isinstance(geom, (rg.Extrusion, rg.Surface)):
            brep = geom.ToBrep()
        elif isinstance(geom, rg.Brep):
            brep = geom
        if brep:
            zs = []
            for face in brep.Faces:
                u = (face.Domain(0).Min + face.Domain(0).Max) * 0.5
                v = (face.Domain(1).Min + face.Domain(1).Max) * 0.5
                if abs(face.NormalAt(u, v).Z) > 0.9:
                    zs.append(face.PointAt(u, v).Z)
            if zs:
                return min(zs)
        bbox = rhobj.Geometry.GetBoundingBox(True)
        return bbox.Min.Z if bbox.IsValid else None
    except Exception:
        return None


def _group_by_z(guids, z_tol):
    """Split guids into elevation bands using z_tol. Returns list of guid lists."""
    pairs = sorted(
        ((get_object_bottom_z(g) or 0.0, g) for g in guids),
        key=lambda p: p[0]
    )
    groups = []  # [band_min_z, [guids]]
    for z, g in pairs:
        if not groups or z - groups[-1][0] > z_tol:
            groups.append([z, [g]])
        else:
            groups[-1][1].append(g)
    return [band[1] for band in groups]


def _region_area(curves):
    """
    Area of closed planar curves treated as region boundaries: nested curves
    count as holes (subtracted), never as extra area.

    Curve.CreateBooleanUnion can return inner (hole) loops — e.g. footprints
    arranged in a ring around a courtyard. Summing absolute curve areas would
    then count the enclosed void as footprint (outer loop already contains it,
    and the hole loop would be added again). CreatePlanarBreps resolves the
    nesting: inner loops become brep holes, so GetArea() is correct.
    """
    curves = list(curves)
    if len(curves) == 1:
        return curve_area(curves[0])
    tol = _bool_tol()
    try:
        breps = rg.Brep.CreatePlanarBreps(curves, tol)
    except Exception:
        breps = None
    if breps:
        return sum(b.GetArea() for b in breps)
    return sum(curve_area(c) for c in curves)


def _interior_point(brep):
    """A point strictly inside a planar Brep face (avoids holes and edges)."""
    try:
        face = brep.Faces[0]
        du, dv = face.Domain(0), face.Domain(1)
        for n in (5, 9, 17):
            for iu in range(1, n):
                for iv in range(1, n):
                    u = du.Min + du.Length * iu / float(n)
                    v = dv.Min + dv.Length * iv / float(n)
                    if face.IsPointOnFace(u, v) == rg.PointFaceRelation.Interior:
                        return face.PointAt(u, v)
    except Exception:
        pass
    return None


def _covered_by_object(pt, breps, loops, tol):
    """
    True if pt lies on an object's footprint region.

    Primary test: distance from pt to the object's planar region Breps —
    zero (within tolerance) iff pt is on a face; points in a hole or outside
    are pulled to the boundary, giving a positive distance. This reuses the
    exact same Breps that produce the per-object areas, so merge coverage
    can never disagree with the individual results.
    Fallback (no Breps): even-odd parity over the raw loops via
    Curve.Contains — less robust, some curve types misreport containment.
    """
    if breps:
        for b in breps:
            try:
                cp = b.ClosestPoint(pt)
                if cp and cp.IsValid and cp.DistanceTo(pt) <= tol * 2.0:
                    return True
            except Exception:
                pass
        return False
    plane = rg.Plane.WorldXY
    inside = 0
    for c in loops:
        try:
            if c.Contains(pt, plane, tol) == rg.PointContainment.Inside:
                inside += 1
        except Exception:
            pass
    return inside % 2 == 1


def _loop_brep(curve, tol):
    """Planar Brep of a single closed loop treated as filled, or None."""
    try:
        bs = rg.Brep.CreatePlanarBreps([curve], tol)
        return bs[0] if bs else None
    except Exception:
        return None


def _pairwise_union_area(per_object_regions, tol, sliver):
    """
    Union area via inclusion–exclusion truncated at pairs:

        total = Σ area_i  −  Σ_{i<j} area(R_i ∩ R_j)

    Pairwise region intersections are computed loop-by-loop with
    Curve.CreateBooleanIntersection — two-curve booleans are far better
    conditioned than the global arrangement, which corrupts on coincident
    tile edges (debug sessions 01.07.2026). Holes are handled with signed
    nesting: R = Σ outer − Σ hole, so
        area(R_i ∩ R_j) = Σ_a Σ_b sign_a · sign_b · area(a ∩ b).

    Exact unless three or more objects overlap on the same spot (rare in
    floor-plate models) — such triple overlaps make this an undercount.
    Returns (area, ok).
    """
    intersect = getattr(rg.Curve, "CreateBooleanIntersection", None)
    if intersect is None:
        return 0.0, False

    objs = []
    total = 0.0
    for breps_k, loops_k in per_object_regions:
        # Sign per loop: +1 if inside an even number of the object's other
        # loops (outer boundary / island), -1 if odd (hole).
        loop_breps = [_loop_brep(c, tol) for c in loops_k]
        signs = []
        for li in range(len(loops_k)):
            b = loop_breps[li]
            pt = _interior_point(b) if b else None
            if pt is None and len(loops_k) > 1:
                return 0.0, False  # cannot classify nesting reliably
            depth = 0
            if pt is not None:
                for lj, other in enumerate(loop_breps):
                    if lj == li or other is None:
                        continue
                    try:
                        cp = other.ClosestPoint(pt)
                        if cp and cp.IsValid and cp.DistanceTo(pt) <= tol * 2.0:
                            depth += 1
                    except Exception:
                        pass
            signs.append(1 if depth % 2 == 0 else -1)
        area_k = (sum(b.GetArea() for b in breps_k) if breps_k
                  else _region_area(loops_k))
        total += area_k
        objs.append((loops_k, signs))

    for i in range(len(objs)):
        loops_i, signs_i = objs[i]
        for j in range(i + 1, len(objs)):
            loops_j, signs_j = objs[j]
            overlap = 0.0
            for a, sa in zip(loops_i, signs_i):
                for b, sb in zip(loops_j, signs_j):
                    try:
                        xs = intersect(a, b, tol)
                    except Exception:
                        return 0.0, False
                    if xs and len(xs) > 0:
                        overlap += sa * sb * _region_area(list(xs))
            if overlap > sliver:
                total -= overlap
    return total, True


def combined_area(obj_guids, cache=None):
    """
    Total footprint area for a list of objects with overlaps removed.
    Returns (area: float, union_succeeded: bool).
    cache: optional {str(guid): curves} memo shared across one calculation
    run (see _fp) — the union logic itself is unchanged.

    Primary path (exact for overlaps, holes, and holes covered by other
    objects): Curve.CreateBooleanRegions builds the planar arrangement of
    every footprint loop — the minimal faces the loops cut the plane into.
    Each face is kept iff an interior sample point lies on at least one
    object's planar-Brep region (holes resolved; see _covered_by_object).
    NOTE: do not replace this with Curve.CreateBooleanUnion or
    Brep.CreatePlanarUnion of the raw loops — both treat/collapse hole
    loops as filled, which re-adds courtyard openings to the total.
    Fallback path: Curve.CreateBooleanUnion (hole-blind, gross areas).
    """
    tol = _bool_tol()  # NOT the raw doc tolerance — see _bool_tol()
    sliver = (tol * 10.0) ** 2  # ignore degenerate arrangement slivers

    all_curves = []          # every loop, flat
    per_object_regions = []  # (breps or None, loops) per object, for coverage
    indiv_total = 0.0        # sum of per-object region areas (upper bound)
    for guid in obj_guids:
        curves = _fp(guid, cache)
        if not curves:
            continue
        all_curves.extend(curves)
        breps = None
        try:
            breps = rg.Brep.CreatePlanarBreps(curves, tol)
        except Exception:
            breps = None
        per_object_regions.append((list(breps) if breps else None, curves))
        indiv_total += (sum(b.GetArea() for b in breps) if breps
                        else _region_area(curves))

    if not all_curves:
        return 0.0, False
    if len(all_curves) == 1:
        return curve_area(all_curves[0]), True
    if len(per_object_regions) == 1:
        # Single object: no overlap possible, region area handles its holes.
        return _region_area(per_object_regions[0][1]), True

    boolean_regions = getattr(rg.Curve, "CreateBooleanRegions", None)
    if boolean_regions is not None:
        try:
            arrangement = boolean_regions(all_curves, rg.Plane.WorldXY, False, tol)
        except Exception:
            arrangement = None
        if arrangement and arrangement.RegionCount > 0:
            total = 0.0
            covered_per_obj = [0.0] * len(per_object_regions)
            resolved = True
            for i in range(arrangement.RegionCount):
                loops = list(arrangement.RegionCurves(i) or [])
                if not loops:
                    continue
                try:
                    breps = rg.Brep.CreatePlanarBreps(loops, tol)
                except Exception:
                    breps = None
                if not breps:
                    if _region_area(loops) <= sliver:
                        continue
                    resolved = False
                    break
                for b in breps:
                    b_area = b.GetArea()
                    if b_area <= sliver:
                        continue
                    pt = _interior_point(b)
                    if pt is None:
                        resolved = False
                        break
                    hit = False
                    for k, (breps_k, loops_k) in enumerate(per_object_regions):
                        if _covered_by_object(pt, breps_k, loops_k, tol):
                            covered_per_obj[k] += b_area
                            hit = True
                    if hit:
                        total += b_area
                if not resolved:
                    break
            # Conservation check: the arrangement faces covered by object k
            # partition its footprint, so their areas must sum back to the
            # object's own area. A corrupt arrangement (coincident tile
            # edges are the known trigger) fails this loudly in either
            # direction — under- OR over-counting. Reject it entirely.
            if resolved:
                for k, (breps_k, loops_k) in enumerate(per_object_regions):
                    a_k = (sum(b.GetArea() for b in breps_k) if breps_k
                           else _region_area(loops_k))
                    if abs(covered_per_obj[k] - a_k) > max(sliver * 100.0,
                                                           0.001 * a_k):
                        resolved = False
                        break
            if resolved and total <= indiv_total + sliver:
                return total, True

    # ── Path 2: pairwise inclusion–exclusion. Robust where the global
    # arrangement corrupts (coincident seams); exact except for ≥3-way
    # overlapping objects, which floor-plate models rarely have.
    pw_total, pw_ok = _pairwise_union_area(per_object_regions, tol, sliver)
    if pw_ok:
        return min(max(pw_total, 0.0), indiv_total), True

    # Normalize all curves to CCW when viewed from +Z so that Boolean Union
    # works consistently regardless of object type (solid bottom faces come out
    # CW; user-drawn closed curves and hatch loops are typically CCW).
    xy_plane = rg.Plane.WorldXY
    normalized = []
    for c in all_curves:
        dup = c.DuplicateCurve()
        if rg.Curve.ClosedCurveOrientation(dup, xy_plane) == rg.CurveOrientation.Clockwise:
            dup.Reverse()
        normalized.append(dup)

    try:
        unioned = rg.Curve.CreateBooleanUnion(normalized, tol)
    except Exception:
        unioned = None

    if unioned and len(unioned) > 0:
        return min(_region_area(unioned), indiv_total), True

    return min(sum(curve_area(c) for c in normalized), indiv_total), False


# ══════════════════════════════════════════════════════════════════════════════
# Calculation engines
# ══════════════════════════════════════════════════════════════════════════════

def _label(guid, key):
    """Object display label: user text key → Rhino object name → short GUID."""
    if key:
        val = rs.GetUserText(guid, key)
        if val:
            return val
    name = rs.ObjectName(guid)
    return name if name else str(guid)[:8] + "…"


def _banded_union(guids, z_height_tol, cache):
    """
    Overlap-merged area with Z-banding: guids are split into elevation
    bands first (stacked floors in one group must be SUMMED, not merged),
    then unioned per band. Returns (area, union_ok).

    Bands where NO object has a calculable footprint (text, annotations,
    points) contribute 0 and are skipped WITHOUT flagging union failure —
    combined_area's (0.0, False) for empty input means "nothing to merge",
    not "merge failed".
    """
    total, ok = 0.0, True
    for band in _group_by_z(guids, z_height_tol):
        if not any(_fp(g, cache) for g in band):
            continue
        area, band_ok = combined_area(band, cache)
        total += area
        if not band_ok:
            ok = False
    return total, ok


# ── Aggregation dimensions (S4, R1/R2) ───────────────────────────────────────
# A dimension is a (kind, arg) pair:
#   (DIM_LEVEL, None)      — the record's level (direct child of the parent)
#   (DIM_CATEGORY, None)   — the record's category (its own lowest layer)
#   (DIM_LAYER, depth)     — layer-path segment at 1-based depth below parent
#   (DIM_USERTEXT, key)    — user text value on the object

DIM_LEVEL    = "level"
DIM_CATEGORY = "category"
DIM_LAYER    = "layer"
DIM_USERTEXT = "usertext"


def _dim_value(rec, dim, parent_layer):
    """Value of one dimension for a collect_objects record. Missing values
    become the placeholder '—' so groups stay comparable."""
    kind, arg = dim
    if kind == DIM_LEVEL:
        return rec["level"]
    if kind == DIM_CATEGORY:
        return rec["category"]
    if kind == DIM_LAYER:
        return _lp.segment_at_depth(parent_layer, rec["layer"], int(arg))
    val = rs.GetUserText(rec["guid"], arg) if arg else None
    return val or _lp.PLACEHOLDER


def _dim_label(dim):
    """Human-readable dimension name for UI, status bar, and Excel headers."""
    kind, arg = dim
    if kind == DIM_LEVEL:
        return "Level (layer)"
    if kind == DIM_CATEGORY:
        return "Category (layer)"
    if kind == DIM_LAYER:
        return f"Layer @ {int(arg)}"
    return str(arg) if arg else _lp.PLACEHOLDER


def calc_s1(name_key, z_height_tol=0.5):
    """
    Scenario 1 — Selected objects. Footprints merged per layer to remove
    intra-layer overlaps; per-layer totals are then summed.
    Returns {objects, total, union_ok, skipped, layer_totals, z_warning}.
    """
    guids = rs.SelectedObjects() or []
    per_obj = []
    skipped = 0
    cache = {}
    layer_groups = {}  # {layer_name: [guids]}

    for g in guids:
        curves = _fp(g, cache)
        if not curves:
            skipped += 1
        area = _region_area(curves)
        lname = rs.ObjectLayer(g) or ""
        layer_groups.setdefault(lname, []).append(g)
        per_obj.append({"guid": str(g), "name": _label(g, name_key), "area": area, "layer": lname})

    layer_totals = []
    total = 0.0
    union_ok = True
    for lname, lguids in layer_groups.items():
        layer_area, ok = _banded_union(lguids, z_height_tol, cache)
        if not ok:
            union_ok = False
        layer_totals.append({"layer": lname, "area": layer_area})
        total += layer_area

    # Z-warning: detect cross-layer groups that share the same floor elevation,
    # meaning their footprints may overlap but won't be merged.
    z_warning = None
    if len(layer_groups) > 1:
        layer_z = {}
        for lname, lguids in layer_groups.items():
            zs = [z for z in (get_object_bottom_z(g) for g in lguids) if z is not None]
            if zs:
                layer_z[lname] = min(zs)

        same_z_pairs = []
        names = list(layer_z.keys())
        for i in range(len(names)):
            for j in range(i + 1, len(names)):
                if abs(layer_z[names[i]] - layer_z[names[j]]) <= z_height_tol:
                    same_z_pairs.append((short_name(names[i]), short_name(names[j])))

        if same_z_pairs:
            pairs_str = ", ".join(f"'{a}' & '{b}'" for a, b in same_z_pairs[:3])
            z_warning = (
                f"Same elevation: {pairs_str}. "
                "Cross-layer overlaps are not merged — use S3 for precise results."
            )

    return {
        "objects": per_obj,
        "total": total,
        "union_ok": union_ok,
        "skipped": skipped,
        "layer_totals": layer_totals,
        "z_warning": z_warning,
    }


def calc_s2(layer_name, obj_key, z_height_tol=0.5):
    """
    Scenario 2 — All objects on one layer, footprints merged.
    Returns {objects, total, union_ok, skipped}.
    skipped: number of objects with no calculable footprint (unsupported type).
    """
    guids = get_layer_objects(layer_name)
    per_obj = []
    skipped = 0
    cache = {}
    for g in guids:
        curves = _fp(g, cache)
        if not curves:
            skipped += 1
        area = _region_area(curves)
        per_obj.append({"guid": str(g), "name": _label(g, obj_key), "area": area})
    total, union_ok = _banded_union(guids, z_height_tol, cache)
    return {"objects": per_obj, "total": total, "union_ok": union_ok, "skipped": skipped}


def calc_s3(parent_layer, obj_key, ignore_prefix=_lp.DEFAULT_IGNORE_PREFIX,
            z_height_tol=0.5):
    """
    Scenario 3 — Layer Hierarchy (v2). Levels and categories come from the
    layer tree itself (see _paths.py); objects are collected from ALL
    non-ignored descendant layers of the parent.

    Per level:
      category_totals — union per (level, category), Z-banded
      total           — union across ALL the level's objects, Z-banded;
                        Σ(category_totals) − total = cross-category overlap
    overall_total    = Σ level totals (floors are additive — GFA logic)
    category_overall = per category, Σ of its per-level unions

    Returns {levels: {name: {objects, category_totals, total, union_ok,
                             skipped, warnings}},
             overall_total, category_overall, layer_info, warnings}.
    Level order follows _lp.level_sort_key (top floor first).
    """
    records, info = collect_objects(parent_layer, ignore_prefix)
    cache = {}
    result = {
        "levels": {},
        "overall_total": 0.0,
        "category_overall": {},
        "layer_info": info,
        "warnings": [],
    }

    by_level = {}
    for r in records:
        by_level.setdefault(r["level"], []).append(r)

    for level in sorted(by_level, key=_lp.level_sort_key):
        recs = by_level[level]
        objects, cat_groups = [], {}
        skipped = direct_on_level = 0
        for r in recs:
            curves = _fp(r["guid"], cache)
            if not curves:
                skipped += 1
            if r["category"] == _lp.PLACEHOLDER:
                direct_on_level += 1
            objects.append({
                "guid": str(r["guid"]),
                "name": _label(r["guid"], obj_key),
                "category": r["category"],
                "area": _region_area(curves),
            })
            cat_groups.setdefault(r["category"], []).append(r["guid"])

        total, union_ok = _banded_union(
            [r["guid"] for r in recs], z_height_tol, cache)

        category_totals = {}
        for cat, cat_guids in cat_groups.items():
            area, cat_ok = _banded_union(cat_guids, z_height_tol, cache)
            category_totals[cat] = area
            if not cat_ok:
                union_ok = False
            result["category_overall"][cat] = (
                result["category_overall"].get(cat, 0.0) + area)

        warnings = []
        cross = sum(category_totals.values()) - total
        if cross > 1e-6:
            warnings.append(
                f"Cross-category overlap: {cross:,.4f} — "
                "categories share footprint area on this level")
        if direct_on_level:
            warnings.append(
                f"{direct_on_level} object(s) directly on the level layer "
                f"(category '{_lp.PLACEHOLDER}')")

        result["levels"][level] = {
            "objects": objects,
            "category_totals": category_totals,
            "total": total,
            "union_ok": union_ok,
            "skipped": skipped,
            "warnings": warnings,
        }
        result["overall_total"] += total

    if info["parent_direct_objects"]:
        result["warnings"].append(
            f"[!] {info['parent_direct_objects']} object(s) directly on the "
            "parent layer are not measured")
    if not records:
        result["warnings"].append(
            "[!] No measurable objects found — check the parent layer, "
            f"the ignore prefix ('{ignore_prefix}'), and the layer structure.")
    return result


def calc_s4(parent_layer, dims, ignore_prefix=_lp.DEFAULT_IGNORE_PREFIX,
            z_height_tol=0.5):
    """
    Scenario 4 — Custom Aggregation (v2).
    dims: ordered list of (kind, arg) dimensions — see _dim_value(). Each
    entry is either (DIM_LAYER, depth) or (DIM_USERTEXT, key), so the old
    two-key S3 and the new layer hierarchy are both configs of this.

    Objects come from collect_objects (whole subtree, ignore prefix).
    Per LEVEL, objects are grouped by their full dimension-value path;
    footprints within each leaf group are merged (Z-banded union) per
    level, then leaf totals are summed across levels.

    Returns {tree, overall_total, warnings, layer_info}.
    tree: nested dict {value_str: {"area": float, "children": {...}}} —
    each node's area = cumulative sum of all descendant leaf areas.
    """
    if not dims:
        return {"tree": {}, "overall_total": 0.0,
                "warnings": ["[!] No dimensions defined."], "layer_info": None}

    records, info = collect_objects(parent_layer, ignore_prefix)
    cache = {}

    # Per-object individual areas — feeds "Write Area to Objects" on the S4
    # tab. The cache makes this free: the same projections are reused by the
    # unions below.
    objects = [
        {"guid": str(r["guid"]), "area": _region_area(_fp(r["guid"], cache))}
        for r in records
    ]

    # Group per (level, full dimension path) — the level keeps the union
    # boundary at one floor, exactly like S3.
    level_path_groups = {}
    for r in records:
        path = tuple(_dim_value(r, d, parent_layer) for d in dims)
        level_path_groups.setdefault((r["level"], path), []).append(r["guid"])

    # flat_buckets[path_tuple] = cumulative area across all levels
    flat_buckets = {}
    for (_level, path), path_guids in level_path_groups.items():
        area, _ok = _banded_union(path_guids, z_height_tol, cache)
        flat_buckets[path] = flat_buckets.get(path, 0.0) + area

    # Build nested tree — each ancestor accumulates all descendant leaf areas
    tree = {}
    for path in sorted(flat_buckets):
        leaf_area = flat_buckets[path]
        node = tree
        for key_val in path:
            if key_val not in node:
                node[key_val] = {"area": 0.0, "children": {}}
            node[key_val]["area"] += leaf_area
            node = node[key_val]["children"]

    overall_total = sum(v["area"] for v in tree.values())

    warnings = []
    if not flat_buckets:
        warnings.append("[!] No measurable objects found under the parent "
                        f"(ignore prefix '{ignore_prefix}').")

    return {"tree": tree, "overall_total": overall_total,
            "warnings": warnings, "layer_info": info, "objects": objects}


def calc_r(parent_layer, dim, targets, ignore_prefix=_lp.DEFAULT_IGNORE_PREFIX,
           z_height_tol=0.5):
    """
    R1 / R2 shared engine (v2).
    Aggregates merged footprint areas by ONE dimension (see _dim_value)
    across all levels: per level, footprints within each dimension-value
    group are merged (Z-banded union) to remove same-group overlaps, then
    the merged areas are summed across levels.

    targets: {label: float} from the Settings target table. Labels without
    a target chart as "(no target)" and produce a warning.

    Returns {entries: [{label, measured, goal}], warnings: [str]}.
    """
    records, _info = collect_objects(parent_layer, ignore_prefix)
    cache = {}

    level_groups = {}
    for r in records:
        v = _dim_value(r, dim, parent_layer)
        level_groups.setdefault((r["level"], v), []).append(r["guid"])

    area_by_val = {}
    for (_level, v), guids in level_groups.items():
        area, _ok = _banded_union(guids, z_height_tol, cache)
        area_by_val[v] = area_by_val.get(v, 0.0) + area

    entries, warnings = [], []
    targets = targets or {}
    for v in sorted(area_by_val):
        goal = targets.get(v)
        if goal is None:
            warnings.append(f"[!] No target set for '{v}' (Settings → Target Areas)")
        entries.append({"label": v, "measured": area_by_val[v], "goal": goal})

    return {"entries": entries, "warnings": warnings}


# ══════════════════════════════════════════════════════════════════════════════
# Number formatting (results grids + status bar)
# ══════════════════════════════════════════════════════════════════════════════

_DECIMALS = 2  # decimal places — updated from Settings before each calculation


def _fmt(v):
    return f"{v:,.{_DECIMALS}f}"


# ══════════════════════════════════════════════════════════════════════════════
# Bullet chart drawing helpers (R1 / R2)
# ══════════════════════════════════════════════════════════════════════════════

_CHART_ROW_H   = 54
_CHART_LABEL_W = 128
_CHART_VALUE_W = 108
_CHART_BAR_H   = 18


def _rgb(r, g, b):
    """Create an Eto.Drawing.Color from 0-255 integer RGB values."""
    return drawing.Color(r / 255.0, g / 255.0, b / 255.0)


def _draw_bullet_row(g, i, entry, tol, unit, total_w):
    """Draw one bullet-chart row at vertical slot i."""
    y0    = i * _CHART_ROW_H + 4
    label = entry["label"]
    meas  = entry["measured"]
    goal  = entry.get("goal")

    font    = drawing.Font("Arial", 8)
    sm_font = drawing.Font("Arial", 7)

    # Left label (truncated)
    g.DrawText(font, _t.CHART_LABEL,
               drawing.PointF(4.0, float(y0 + (_CHART_ROW_H - 12) // 2)),
               label[:17])

    chart_x = float(_CHART_LABEL_W)
    chart_w = float(total_w - _CHART_LABEL_W - _CHART_VALUE_W)
    bar_y   = float(y0 + (_CHART_ROW_H - _CHART_BAR_H) // 2)
    bar_h   = float(_CHART_BAR_H)

    if chart_w < 10:
        return

    if goal is None or goal <= 0:
        g.DrawText(font, _t.CHART_NOTGT,
                   drawing.PointF(chart_x + 4.0, bar_y),
                   f"{meas:,.2f} (no target)")
        return

    max_val = max(goal * 1.35, meas * 1.05) if meas > goal else goal * 1.35
    if max_val <= 0:
        return

    def px(v):
        return chart_x + float(v) / max_val * chart_w

    # 1. Background bar
    g.FillRectangle(_t.CHART_BG,
                    drawing.RectangleF(chart_x, bar_y, chart_w, bar_h))

    # 2. Below-target zone: goal*(1-tol) → goal
    yl = px(max(0.0, goal * (1.0 - tol)))
    yr = px(goal)
    if yr > yl:
        g.FillRectangle(_t.CHART_LOW,
                        drawing.RectangleF(yl, bar_y, yr - yl, bar_h))

    # 3. Above-target zone: goal → goal*(1+tol)
    ol  = px(goal)
    or_ = px(min(max_val, goal * (1.0 + tol)))
    if or_ > ol:
        g.FillRectangle(_t.CHART_HIGH,
                        drawing.RectangleF(ol, bar_y, or_ - ol, bar_h))

    # 4. Measured bar
    mx = px(meas) - chart_x
    if mx > 0:
        g.FillRectangle(_t.CHART_BAR,
                        drawing.RectangleF(chart_x, bar_y,
                                           min(float(mx), chart_w), bar_h))

    # 5. Goal line (2 px)
    gx = px(goal)
    g.DrawLine(drawing.Pen(_t.CHART_GOAL, 2.0),
               drawing.PointF(gx, bar_y - 2.0),
               drawing.PointF(gx, bar_y + bar_h + 2.0))

    # 6. Tolerance marker lines
    tpen = drawing.Pen(_t.CHART_TOL, 1.0)
    for tx in (px(goal * (1.0 - tol)), px(goal * (1.0 + tol))):
        if chart_x <= tx <= chart_x + chart_w:
            g.DrawLine(tpen,
                       drawing.PointF(tx, bar_y),
                       drawing.PointF(tx, bar_y + bar_h))

    # 7. Right labels: measured/goal and delta %
    delta = (meas - goal) / goal * 100.0
    sign  = "+" if delta >= 0 else ""
    vx    = float(total_w - _CHART_VALUE_W + 4)
    g.DrawText(font,    _t.CHART_LABEL,
               drawing.PointF(vx, bar_y - 1.0),
               f"{meas:,.1f}/{goal:,.1f}")
    g.DrawText(sm_font, _t.CHART_DELTA,
               drawing.PointF(vx, bar_y + 11.0),
               f"{sign}{delta:.1f}%  [{unit}]")


def _export_chart_png(entries, tol, unit, path, chart_width=900):
    """Render bullet-chart entries to a PNG file at the given path."""
    n = max(1, len(entries))
    height = n * _CHART_ROW_H + 20
    bmp = drawing.Bitmap(chart_width, height, drawing.PixelFormat.Format32bppRgba)
    g = drawing.Graphics(bmp)
    try:
        g.FillRectangle(drawing.Colors.White,
                        drawing.RectangleF(0.0, 0.0, float(chart_width), float(height)))
        for i, entry in enumerate(entries):
            _draw_bullet_row(g, i, entry, tol, unit, chart_width)
    finally:
        g.Dispose()
    bmp.Save(path)
    bmp.Dispose()


# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════

class LinderoForm(forms.Form):
    """Modeless footprint area calculator — stays open between calculations."""

    def __init__(self):
        super().__init__()
        self.Title = f"Lindero — Footprint Area Calculator  {_VERSION_TEXT}"
        self.Padding = drawing.Padding(10)
        self.Resizable = True
        self.MinimumSize = drawing.Size(480, 540)
        self.ClientSize  = drawing.Size(680, 720)
        self.BackgroundColor = _t.BG
        self.Owner = Rhino.UI.RhinoEtoApp.MainWindow

        self.available_keys   = get_all_user_text_keys()
        self.available_layers = all_layer_names()

        # Last-calculation state (for Write Area and Export)
        self._export_data = None
        self._last_s1 = None   # list of {guid, name, area}
        self._last_s2 = None   # {objects, total, union_ok}
        self._last_s3 = None   # calc_s3 result (levels, category_overall, …)
        self._last_s4 = None   # calc_s4 result (tree, overall_total, …)

        # Search-filter updater functions for key ComboBoxes (populated by tab builders)
        self._ks_combos = []   # [updater, ...] for static key combos
        self._ks_write  = None # updater for _write_key_combo (has extra "Area" item)
        self._ks_s4     = {}   # {id(cb): updater} for dynamic S4 usertext rows
        self._ks_layer_s2     = None   # layer search updaters — set by tab builders
        self._ks_layer_parent = None
        self._ks_layer_s4_par = None

        # Settings target table rows (dynamic)
        self._target_rows = []  # [{"label_tb", "value_tb", "row"}]

        # R1 / R2 chart state
        self._r1_entries = []
        self._r1_tol     = 0.10
        self._r1_unit    = ""
        self._r2_entries = []
        self._r2_tol     = 0.10
        self._r2_unit    = ""

        self._build_ui()

        # Per-document persistence: restore the settings stored in the 3dm
        # and write them back when the window closes (see _store_doc_config).
        self.Closed += self.on_form_closed
        if self._load_doc_config():
            self.status_label.Text += "   ·   settings restored from document"

    # ------------------------------------------------------------------
    # Layout builders
    # ------------------------------------------------------------------

    def _build_ui(self):
        outer = forms.StackLayout()
        outer.Orientation = forms.Orientation.Vertical
        outer.Spacing = 8
        outer.HorizontalContentAlignment = forms.HorizontalAlignment.Stretch

        # Tabs — expand to fill available height
        self.tabs = forms.TabControl()
        self.tabs.Pages.Add(self._tab_s1())
        self.tabs.Pages.Add(self._tab_s2())
        self.tabs.Pages.Add(self._tab_s3())
        self.tabs.Pages.Add(self._tab_s4())
        self.tabs.Pages.Add(self._tab_r1())
        self.tabs.Pages.Add(self._tab_r2())
        self.tabs.Pages.Add(self._tab_settings())
        outer.Items.Add(forms.StackLayoutItem(self.tabs, True))

        # Button row
        btn_row = forms.StackLayout()
        btn_row.Orientation = forms.Orientation.Horizontal
        btn_row.Spacing = 6

        calc_btn = forms.Button()
        calc_btn.Text = "Calculate"
        calc_btn.Font = _t.F_SANS_B
        calc_btn.BackgroundColor = _t.BTN_CALC
        calc_btn.Click += self.on_calculate

        clear_btn = forms.Button()
        clear_btn.Text = "Clear"
        clear_btn.Font = _t.F_SANS_B
        clear_btn.BackgroundColor = _t.BTN_CLEAR
        clear_btn.Click += self.on_clear

        refresh_btn = forms.Button()
        refresh_btn.Text = "Refresh Model"
        refresh_btn.Font = _t.F_SANS_B
        refresh_btn.BackgroundColor = _t.BTN_DEFAULT
        refresh_btn.Click += self.on_refresh_model

        export_btn = forms.Button()
        export_btn.Text = "Export to Excel"
        export_btn.Font = _t.F_SANS_B
        export_btn.BackgroundColor = _t.BTN_DEFAULT
        export_btn.Click += self.on_export

        write_btn = forms.Button()
        write_btn.Text = "Write Area to Objects"
        write_btn.Font = _t.F_SANS_B
        write_btn.BackgroundColor = _t.BTN_DEFAULT
        write_btn.Click += self.on_write_area_toggle

        png_btn = forms.Button()
        png_btn.Text = "Export Chart as PNG"
        png_btn.Font = _t.F_SANS_B
        png_btn.BackgroundColor = _t.BTN_DEFAULT
        png_btn.Click += self.on_export_png

        copy_btn = forms.Button()
        copy_btn.Text = "Copy Window"
        copy_btn.Font = _t.F_SANS_B
        copy_btn.BackgroundColor = _t.BTN_DEFAULT
        copy_btn.ToolTip = "Screenshot this window to the clipboard (paste with Ctrl+V)"
        copy_btn.Click += self.on_copy_window

        btn_row.Items.Add(forms.StackLayoutItem(calc_btn))
        btn_row.Items.Add(forms.StackLayoutItem(clear_btn))
        btn_row.Items.Add(forms.StackLayoutItem(refresh_btn))
        btn_row.Items.Add(forms.StackLayoutItem(export_btn))
        btn_row.Items.Add(forms.StackLayoutItem(write_btn))
        btn_row.Items.Add(forms.StackLayoutItem(png_btn))
        btn_row.Items.Add(forms.StackLayoutItem(copy_btn))
        outer.Items.Add(forms.StackLayoutItem(btn_row))

        # Write panel (hidden until toggled)
        outer.Items.Add(forms.StackLayoutItem(self._build_write_panel()))

        # Status label
        self.status_label = forms.Label()
        self.status_label.Text = (
            f"Ready  —  {len(self.available_layers)} layer(s), "
            f"{len(self.available_keys)} key(s) loaded."
        )
        self.status_label.Font = _t.F_SANS
        self.status_label.TextColor = _t.TEXT_MUTED

        version_label = forms.Label()
        version_label.Text = _VERSION_TEXT
        version_label.Font = _t.F_SANS
        version_label.TextColor = _t.TEXT_MUTED

        status_row = forms.StackLayout()
        status_row.Orientation = forms.Orientation.Horizontal
        status_row.Spacing = _t.SPACE_2
        status_row.Items.Add(forms.StackLayoutItem(self.status_label, True))
        status_row.Items.Add(forms.StackLayoutItem(version_label))
        outer.Items.Add(forms.StackLayoutItem(status_row))

        self.Content = outer

    def _build_write_panel(self):
        self._write_panel = forms.StackLayout()
        self._write_panel.Orientation = forms.Orientation.Horizontal
        self._write_panel.Spacing = 6
        self._write_panel.Padding = drawing.Padding(0, 2, 0, 2)
        self._write_panel.Visible = False

        wk_lbl = forms.Label()
        wk_lbl.Text = "Write to Key:"

        self._write_key_combo = forms.ComboBox()
        self._write_key_combo.DataStore = ["Area"] + list(self.available_keys)
        self._write_key_combo.Text = "Area"
        self._write_key_combo.Width = 180
        self._ks_write = _t.bind_key_search(self._write_key_combo, ["Area"] + list(self.available_keys))

        confirm_btn = forms.Button()
        confirm_btn.Text = "Confirm Write"
        confirm_btn.Font = _t.F_SANS_B
        confirm_btn.BackgroundColor = _t.BTN_CALC
        confirm_btn.Click += self.on_confirm_write

        cancel_btn = forms.Button()
        cancel_btn.Text = "Cancel"
        cancel_btn.Font = _t.F_SANS_B
        cancel_btn.BackgroundColor = _t.BTN_CLEAR
        cancel_btn.Click += self._on_write_cancel

        self._write_panel.Items.Add(forms.StackLayoutItem(wk_lbl))
        self._write_panel.Items.Add(forms.StackLayoutItem(self._write_key_combo))
        self._write_panel.Items.Add(forms.StackLayoutItem(confirm_btn))
        self._write_panel.Items.Add(forms.StackLayoutItem(cancel_btn))

        return self._write_panel

    # ------------------------------------------------------------------
    # Tab builders — S1 / S2 / S3
    # ------------------------------------------------------------------

    def _tab_s1(self):
        page = forms.TabPage()
        page.Text = "S1 — Selected Objects"

        controls = forms.DynamicLayout()
        controls.DefaultSpacing = drawing.Size(5, 6)
        controls.Padding = drawing.Padding(8)

        desc = forms.Label()
        desc.Text = "Individual footprint per selected object. No overlap handling."
        desc.TextColor = _t.TEXT_MUTED
        controls.AddRow(desc)
        controls.AddRow(None)

        name_lbl = forms.Label()
        name_lbl.Text = "Name Key (optional):"
        self.name_key_combo = forms.ComboBox()
        self.name_key_combo.DataStore = self.available_keys
        self.name_key_combo.PlaceholderText = "Falls back to object name / GUID"
        self.name_key_combo.Width = 300
        self._ks_combos.append(_t.bind_key_search(self.name_key_combo, self.available_keys))
        controls.AddRow(name_lbl, self.name_key_combo)
        controls.AddRow(None)

        note = forms.Label()
        note.Text = "Select objects in Rhino (form stays open), then click Calculate."
        note.TextColor = _t.TEXT_MUTED
        controls.AddRow(note)

        self.results_s1_grid = self._make_results_grid([(True, 0), (False, 110)])

        layout = forms.StackLayout()
        layout.Orientation = forms.Orientation.Vertical
        layout.HorizontalContentAlignment = forms.HorizontalAlignment.Stretch
        layout.Items.Add(forms.StackLayoutItem(controls))
        layout.Items.Add(forms.StackLayoutItem(self.results_s1_grid, True))

        page.Content = layout
        return page

    def _tab_s2(self):
        page = forms.TabPage()
        page.Text = "S2 — By Layer"

        controls = forms.DynamicLayout()
        controls.DefaultSpacing = drawing.Size(5, 6)
        controls.Padding = drawing.Padding(8)

        desc = forms.Label()
        desc.Text = "All objects on a layer. Overlapping footprints merged (Boolean Union)."
        desc.TextColor = _t.TEXT_MUTED
        controls.AddRow(desc)
        controls.AddRow(None)

        layer_lbl = forms.Label()
        layer_lbl.Text = "Layer:"
        self.layer_s2_dd = forms.ComboBox()
        self.layer_s2_dd.DataStore = list(self.available_layers)
        self.layer_s2_dd.PlaceholderText = "Type to filter layers…"
        self.layer_s2_dd.Width = 300
        self._ks_layer_s2 = _t.bind_key_search(self.layer_s2_dd, self.available_layers)
        controls.AddRow(layer_lbl, self.layer_s2_dd)

        obj_key_lbl = forms.Label()
        obj_key_lbl.Text = "Object Key:"
        self.obj_key_s2 = forms.ComboBox()
        self.obj_key_s2.DataStore = self.available_keys
        self.obj_key_s2.PlaceholderText = "User text key for object labels"
        self.obj_key_s2.Width = 300
        self._ks_combos.append(_t.bind_key_search(self.obj_key_s2, self.available_keys))
        controls.AddRow(obj_key_lbl, self.obj_key_s2)

        self.results_s2_grid = self._make_results_grid([(True, 0), (False, 110)])

        layout = forms.StackLayout()
        layout.Orientation = forms.Orientation.Vertical
        layout.HorizontalContentAlignment = forms.HorizontalAlignment.Stretch
        layout.Items.Add(forms.StackLayoutItem(controls))
        layout.Items.Add(forms.StackLayoutItem(self.results_s2_grid, True))

        page.Content = layout
        return page

    def _tab_s3(self):
        page = forms.TabPage()
        page.Text = "S3 — Layer Hierarchy"

        # Full-width description block — kept OUT of the DynamicLayout:
        # a single-control AddRow sits in column 1, whose width is set by
        # the "Parent Layer:" label, so long text wraps into a tall sliver.
        head = forms.StackLayout()
        head.Orientation = forms.Orientation.Vertical
        head.Spacing = 4
        head.Padding = drawing.Padding(8, 8, 8, 0)
        head.HorizontalContentAlignment = forms.HorizontalAlignment.Stretch

        desc = forms.Label()
        desc.Text = (
            "Levels and categories come from the layer tree — objects are "
            "collected from every non-ignored layer under the parent."
        )
        desc.TextColor = _t.TEXT_MUTED
        desc.Wrap = forms.WrapMode.Word
        head.Items.Add(forms.StackLayoutItem(desc))

        sub3 = forms.Label()
        sub3.Text = "Level = direct child of parent   ·   Category = object's own layer"
        sub3.Font = _t.F_SANS_B
        sub3.TextColor = _t.HEADER
        sub3.Wrap = forms.WrapMode.Word
        head.Items.Add(forms.StackLayoutItem(sub3))

        controls = forms.DynamicLayout()
        controls.DefaultSpacing = drawing.Size(5, 6)
        controls.Padding = drawing.Padding(8)

        parent_lbl = forms.Label()
        parent_lbl.Text = "Parent Layer:"
        self.parent_layer_dd = forms.ComboBox()
        self.parent_layer_dd.DataStore = list(self.available_layers)
        self.parent_layer_dd.PlaceholderText = "Type to filter layers…"
        self.parent_layer_dd.Width = 300
        self._ks_layer_parent = _t.bind_key_search(self.parent_layer_dd, self.available_layers)
        controls.AddRow(parent_lbl, self.parent_layer_dd)

        obj_key_lbl = forms.Label()
        obj_key_lbl.Text = "Object Key:"
        self.obj_key_s3 = forms.ComboBox()
        self.obj_key_s3.DataStore = self.available_keys
        self.obj_key_s3.PlaceholderText = "Optional — object labels only"
        self.obj_key_s3.Width = 300
        self._ks_combos.append(_t.bind_key_search(self.obj_key_s3, self.available_keys))
        obj_key_hint = forms.Label()
        obj_key_hint.Text = "User text key used to label objects in the detail list"
        obj_key_hint.TextColor = _t.TEXT_MUTED
        controls.AddRow(obj_key_lbl, self.obj_key_s3)
        controls.AddRow(None, obj_key_hint)

        # Preview row — own horizontal stack so the button keeps its natural
        # width and the hint wraps in the remaining space.
        prev_row = forms.StackLayout()
        prev_row.Orientation = forms.Orientation.Horizontal
        prev_row.Spacing = 8
        prev_row.Padding = drawing.Padding(8, 0, 8, 4)
        prev_row.VerticalContentAlignment = forms.VerticalAlignment.Center

        preview_btn = forms.Button()
        preview_btn.Text = "Preview Structure"
        preview_btn.Font = _t.F_SANS_B
        preview_btn.BackgroundColor = _t.BTN_DEFAULT
        preview_btn.Click += self.on_preview_s3
        preview_hint = forms.Label()
        preview_hint.Text = (
            "Dry run: included levels/categories with object counts, plus "
            "ignored layers (prefix in Settings, default '_') — no areas computed."
        )
        preview_hint.TextColor = _t.TEXT_MUTED
        preview_hint.Wrap = forms.WrapMode.Word
        prev_row.Items.Add(forms.StackLayoutItem(preview_btn))
        prev_row.Items.Add(forms.StackLayoutItem(preview_hint, True))

        # ── Two-panel split results ─────────────────────────────────────
        brk_hdr = forms.Label()
        brk_hdr.Text = "Breakdown — Level × Category"
        brk_hdr.Font = _t.F_HEAD
        # breakdown grid: label (expand) + area (fixed)
        self.results_s3_breakdown_grid = self._make_results_grid([(True, 0), (False, 110)])
        brk_pane = forms.StackLayout()
        brk_pane.Orientation = forms.Orientation.Vertical
        brk_pane.HorizontalContentAlignment = forms.HorizontalAlignment.Stretch
        brk_pane.Items.Add(forms.StackLayoutItem(brk_hdr))
        brk_pane.Items.Add(forms.StackLayoutItem(self.results_s3_breakdown_grid, True))

        obj_hdr = forms.Label()
        obj_hdr.Text = "Objects — Detail"
        obj_hdr.Font = _t.F_HEAD
        # objects grid: object name (expand) + category (fixed) + area (fixed)
        self.results_s3_objects_grid = self._make_results_grid([(True, 0), (False, 120), (False, 90)])
        obj_pane = forms.StackLayout()
        obj_pane.Orientation = forms.Orientation.Vertical
        obj_pane.HorizontalContentAlignment = forms.HorizontalAlignment.Stretch
        obj_pane.Items.Add(forms.StackLayoutItem(obj_hdr))
        obj_pane.Items.Add(forms.StackLayoutItem(self.results_s3_objects_grid, True))

        splitter = forms.Splitter()
        splitter.Orientation = forms.Orientation.Vertical
        splitter.Panel1 = brk_pane
        splitter.Panel2 = obj_pane
        splitter.Position = 220

        layout = forms.StackLayout()
        layout.Orientation = forms.Orientation.Vertical
        layout.HorizontalContentAlignment = forms.HorizontalAlignment.Stretch
        layout.Items.Add(forms.StackLayoutItem(head))
        layout.Items.Add(forms.StackLayoutItem(controls))
        layout.Items.Add(forms.StackLayoutItem(prev_row))
        layout.Items.Add(forms.StackLayoutItem(splitter, True))

        page.Content = layout
        return page

    # ------------------------------------------------------------------
    # Tab builder — S4 (Custom Aggregation)
    # ------------------------------------------------------------------

    def _tab_s4(self):
        page = forms.TabPage()
        page.Text = "S4 — Custom Aggregation"

        layout = forms.StackLayout()
        layout.Orientation = forms.Orientation.Vertical
        layout.Spacing = 4
        layout.Padding = drawing.Padding(8)
        layout.HorizontalContentAlignment = forms.HorizontalAlignment.Stretch

        desc = forms.Label()
        desc.Text = (
            "Groups objects by a user-defined hierarchy of dimensions — each "
            "one is either a layer-path segment ('Layer @ depth', 1 = level) "
            "or a user text key. Footprints are merged per leaf group per "
            "level (same as S3), then summed across all levels. Ignore "
            "prefix from Settings applies."
        )
        desc.TextColor = _t.TEXT_MUTED
        desc.Wrap = forms.WrapMode.Word
        layout.Items.Add(forms.StackLayoutItem(desc))

        # Parent layer row
        pl_row = forms.DynamicLayout()
        pl_row.DefaultSpacing = drawing.Size(5, 4)
        pl_row.Padding = drawing.Padding(0, 4, 0, 4)
        pl_lbl = forms.Label()
        pl_lbl.Text = "Parent Layer:"
        self.parent_layer_s4_dd = forms.ComboBox()
        self.parent_layer_s4_dd.DataStore = list(self.available_layers)
        self.parent_layer_s4_dd.PlaceholderText = "Type to filter layers…"
        self.parent_layer_s4_dd.Width = 300
        self._ks_layer_s4_par = _t.bind_key_search(self.parent_layer_s4_dd, self.available_layers)
        pl_row.AddRow(pl_lbl, self.parent_layer_s4_dd)
        layout.Items.Add(forms.StackLayoutItem(pl_row))

        # Dimensions section header + Add button
        keys_header = forms.StackLayout()
        keys_header.Orientation = forms.Orientation.Horizontal
        keys_header.Spacing = 10
        keys_header.Padding = drawing.Padding(0, 2, 0, 2)
        keys_lbl = forms.Label()
        keys_lbl.Text = "Dimensions  (top row = hierarchy level 1, bottom = deepest):"
        add_btn = forms.Button()
        add_btn.Text = "+ Add Dimension"
        add_btn.Click += self._on_s4_add_dim
        keys_header.Items.Add(forms.StackLayoutItem(keys_lbl))
        keys_header.Items.Add(forms.StackLayoutItem(add_btn))
        layout.Items.Add(forms.StackLayoutItem(keys_header))

        # Dynamic dimension rows container.
        # Default seed = Layer @ 1 → Layer @ 2 (Level → Category — matches
        # the v2 layer contract out of the box).
        self._s4_keys_layout = forms.StackLayout()
        self._s4_keys_layout.Orientation = forms.Orientation.Vertical
        self._s4_keys_layout.Spacing = 3
        self._s4_dim_rows = []
        self._s4_add_dim_row(DIM_LAYER, 1)
        self._s4_add_dim_row(DIM_LAYER, 2)
        layout.Items.Add(forms.StackLayoutItem(self._s4_keys_layout))

        # Results tree grid
        self.results_s4_grid = forms.TreeGridView()
        self.results_s4_grid.ShowHeader = False
        self.results_s4_grid.AllowColumnReordering = False
        self.results_s4_grid.AllowMultipleSelection = False
        self.results_s4_grid.Font = _t.F_MONO

        _lbl_col = forms.GridColumn()
        _lbl_col.Editable = False
        _lbl_col.DataCell = forms.TextBoxCell(0)
        _lbl_col.Expand = True
        self.results_s4_grid.Columns.Add(_lbl_col)

        _area_col = forms.GridColumn()
        _area_col.Editable = False
        _area_col.DataCell = forms.TextBoxCell(1)
        _area_col.Width = 110
        self.results_s4_grid.Columns.Add(_area_col)

        self.results_s4_grid.CellFormatting += self._on_cell_format
        layout.Items.Add(forms.StackLayoutItem(self.results_s4_grid, True))

        page.Content = layout
        return page

    def _s4_add_dim_row(self, kind=DIM_USERTEXT, arg=None):
        """Append one dimension row: [kind dropdown][key combo | depth stepper][✕]."""
        row = forms.StackLayout()
        row.Orientation = forms.Orientation.Horizontal
        row.Spacing = 4

        kind_dd = forms.DropDown()
        kind_dd.Items.Add("UserText key")
        kind_dd.Items.Add("Layer @ depth")
        kind_dd.Width = 120
        kind_dd.SelectedIndex = 1 if kind == DIM_LAYER else 0

        key_cb = forms.ComboBox()
        key_cb.DataStore = self.available_keys
        key_cb.Text = str(arg) if (kind == DIM_USERTEXT and arg) else ""
        key_cb.Width = 190
        self._ks_s4[id(key_cb)] = _t.bind_key_search(key_cb, self.available_keys)

        depth_st = forms.NumericStepper()
        depth_st.MinValue      = 1
        depth_st.MaxValue      = 8
        depth_st.DecimalPlaces = 0
        depth_st.Increment     = 1
        depth_st.Width         = 60
        depth_st.Value         = int(arg) if (kind == DIM_LAYER and arg) else 1

        rm_btn = forms.Button()
        rm_btn.Text = "✕"
        rm_btn.Width = 28

        rec = {"kind_dd": kind_dd, "key_cb": key_cb,
               "depth_st": depth_st, "row": row}

        def on_kind_changed(s, e, r=rec):
            is_layer = r["kind_dd"].SelectedIndex == 1
            r["key_cb"].Visible   = not is_layer
            r["depth_st"].Visible = is_layer
        kind_dd.SelectedIndexChanged += on_kind_changed

        def on_remove(s, e, r=rec):
            self._s4_remove_dim_row(r)
        rm_btn.Click += on_remove

        row.Items.Add(forms.StackLayoutItem(kind_dd))
        row.Items.Add(forms.StackLayoutItem(key_cb, True))
        row.Items.Add(forms.StackLayoutItem(depth_st))
        row.Items.Add(forms.StackLayoutItem(rm_btn))
        on_kind_changed(None, None)

        self._s4_dim_rows.append(rec)
        self._s4_keys_layout.Items.Add(forms.StackLayoutItem(row))

    def _on_s4_add_dim(self, _s, _e):
        self._s4_add_dim_row()

    def _s4_remove_dim_row(self, rec):
        """Remove a dimension row, keeping at least one."""
        if len(self._s4_dim_rows) <= 1:
            return
        if rec in self._s4_dim_rows:
            self._s4_dim_rows.remove(rec)
            self._ks_s4.pop(id(rec["key_cb"]), None)
        for i in range(self._s4_keys_layout.Items.Count):
            if self._s4_keys_layout.Items[i].Control is rec["row"]:
                self._s4_keys_layout.Items.RemoveAt(i)
                break

    def _s4_clear_dim_rows(self):
        """Remove all dimension rows (used by config load)."""
        while self._s4_keys_layout.Items.Count > 0:
            self._s4_keys_layout.Items.RemoveAt(0)
        for rec in self._s4_dim_rows:
            self._ks_s4.pop(id(rec["key_cb"]), None)
        self._s4_dim_rows = []

    def _s4_dims(self):
        """Current dimension list [(kind, arg)], blank usertext rows dropped."""
        dims = []
        for rec in self._s4_dim_rows:
            if rec["kind_dd"].SelectedIndex == 1:
                dims.append((DIM_LAYER, int(rec["depth_st"].Value)))
            else:
                key = rec["key_cb"].Text.strip()
                if key:
                    dims.append((DIM_USERTEXT, key))
        return dims

    # ------------------------------------------------------------------
    # Tab builders — R1 / R2
    # ------------------------------------------------------------------

    def _tab_r1(self):
        page = forms.TabPage()
        page.Text = "R1 — Room Analysis"

        layout = forms.StackLayout()
        layout.Orientation = forms.Orientation.Vertical
        layout.Spacing = 4
        layout.Padding = drawing.Padding(8)
        layout.HorizontalContentAlignment = forms.HorizontalAlignment.Stretch

        desc = forms.Label()
        desc.Text = (
            "Aggregates merged footprint areas by the chosen dimension across "
            "all levels of the S3 Parent Layer. Compares totals against the "
            "Target Areas table (Settings)."
        )
        desc.TextColor = _t.TEXT_MUTED
        desc.Wrap = forms.WrapMode.Word

        dim_row = forms.StackLayout()
        dim_row.Orientation = forms.Orientation.Horizontal
        dim_row.Spacing = 6
        dim_lbl = forms.Label()
        dim_lbl.Text = "Aggregate by:"
        self.r1_dim_dd = forms.DropDown()
        self.r1_dim_dd.Items.Add("Category (layer)")
        self.r1_dim_dd.Items.Add("Level (layer)")
        self.r1_dim_dd.Items.Add("UserText key")
        self.r1_dim_dd.SelectedIndex = 0
        self.r1_dim_dd.Width = 150
        self.r1_key_cb = forms.ComboBox()
        self.r1_key_cb.DataStore = self.available_keys
        self.r1_key_cb.PlaceholderText = "User text key"
        self.r1_key_cb.Width = 180
        self.r1_key_cb.Visible = False
        self._ks_combos.append(_t.bind_key_search(self.r1_key_cb, self.available_keys))

        def _r1_dim_changed(s, e):
            self.r1_key_cb.Visible = (self.r1_dim_dd.SelectedIndex == 2)
        self.r1_dim_dd.SelectedIndexChanged += _r1_dim_changed

        dim_row.Items.Add(forms.StackLayoutItem(dim_lbl))
        dim_row.Items.Add(forms.StackLayoutItem(self.r1_dim_dd))
        dim_row.Items.Add(forms.StackLayoutItem(self.r1_key_cb))

        self.warn_r1 = forms.Label()
        self.warn_r1.TextColor = _t.TEXT_ERROR
        self.warn_r1.Wrap = forms.WrapMode.Word
        self.warn_r1.Visible = False

        self.chart_r1 = forms.Drawable()
        self.chart_r1.Size = drawing.Size(400, 10)
        self.chart_r1.Paint += self._paint_r1

        scroll_r1 = forms.Scrollable()
        scroll_r1.ExpandContentWidth = True
        scroll_r1.ExpandContentHeight = False
        scroll_r1.Content = self.chart_r1

        layout.Items.Add(forms.StackLayoutItem(desc))
        layout.Items.Add(forms.StackLayoutItem(dim_row))
        layout.Items.Add(forms.StackLayoutItem(self.warn_r1))
        layout.Items.Add(forms.StackLayoutItem(scroll_r1, True))

        page.Content = layout
        return page

    def _tab_r2(self):
        page = forms.TabPage()
        page.Text = "R2 — Group Analysis"

        layout = forms.StackLayout()
        layout.Orientation = forms.Orientation.Vertical
        layout.Spacing = 4
        layout.Padding = drawing.Padding(8)
        layout.HorizontalContentAlignment = forms.HorizontalAlignment.Stretch

        desc = forms.Label()
        desc.Text = (
            "Same engine as R1 with its own dimension — e.g. R1 per category, "
            "R2 per level. Uses the S3 Parent Layer and the Target Areas "
            "table (Settings)."
        )
        desc.TextColor = _t.TEXT_MUTED
        desc.Wrap = forms.WrapMode.Word

        dim_row = forms.StackLayout()
        dim_row.Orientation = forms.Orientation.Horizontal
        dim_row.Spacing = 6
        dim_lbl = forms.Label()
        dim_lbl.Text = "Aggregate by:"
        self.r2_dim_dd = forms.DropDown()
        self.r2_dim_dd.Items.Add("Category (layer)")
        self.r2_dim_dd.Items.Add("Level (layer)")
        self.r2_dim_dd.Items.Add("UserText key")
        self.r2_dim_dd.SelectedIndex = 1
        self.r2_dim_dd.Width = 150
        self.r2_key_cb = forms.ComboBox()
        self.r2_key_cb.DataStore = self.available_keys
        self.r2_key_cb.PlaceholderText = "User text key"
        self.r2_key_cb.Width = 180
        self.r2_key_cb.Visible = False
        self._ks_combos.append(_t.bind_key_search(self.r2_key_cb, self.available_keys))

        def _r2_dim_changed(s, e):
            self.r2_key_cb.Visible = (self.r2_dim_dd.SelectedIndex == 2)
        self.r2_dim_dd.SelectedIndexChanged += _r2_dim_changed

        dim_row.Items.Add(forms.StackLayoutItem(dim_lbl))
        dim_row.Items.Add(forms.StackLayoutItem(self.r2_dim_dd))
        dim_row.Items.Add(forms.StackLayoutItem(self.r2_key_cb))

        self.warn_r2 = forms.Label()
        self.warn_r2.TextColor = _t.TEXT_ERROR
        self.warn_r2.Wrap = forms.WrapMode.Word
        self.warn_r2.Visible = False

        self.chart_r2 = forms.Drawable()
        self.chart_r2.Size = drawing.Size(400, 10)
        self.chart_r2.Paint += self._paint_r2

        scroll_r2 = forms.Scrollable()
        scroll_r2.ExpandContentWidth = True
        scroll_r2.ExpandContentHeight = False
        scroll_r2.Content = self.chart_r2

        layout.Items.Add(forms.StackLayoutItem(desc))
        layout.Items.Add(forms.StackLayoutItem(dim_row))
        layout.Items.Add(forms.StackLayoutItem(self.warn_r2))
        layout.Items.Add(forms.StackLayoutItem(scroll_r2, True))

        page.Content = layout
        return page

    # ------------------------------------------------------------------
    # Tab builder — Settings
    # ------------------------------------------------------------------

    def _tab_settings(self):
        # LAYOUT CONVENTIONS (empirically stable inside this Scrollable —
        # broke twice before settling here, v0.8→v0.8.2):
        #   · label + control rows:      layout.AddRow(lbl, ctrl)
        #   · long hints: explicit "\n" breaks, column 2 via AddRow(None, x).
        #     Never rely on Wrap inside the Scrollable — a wrapping Label
        #     reports its UNWRAPPED width as preferred size and blows the
        #     scroll canvas out horizontally.
        #   · short headers/separators:  AddRow(x) in column 1 is fine.
        #   · the dynamic target table (StackLayout) also lives in column 2 —
        #     rows have fixed widths, so it cannot widen the canvas.
        page = forms.TabPage()
        page.Text = "Settings"

        layout = forms.DynamicLayout()
        layout.DefaultSpacing = drawing.Size(5, 8)
        layout.Padding = drawing.Padding(12)

        # ── General ──────────────────────────────────────────────────
        sec1 = forms.Label()
        sec1.Text = "General"
        sec1.Font = _t.F_HEAD
        layout.AddRow(sec1)
        layout.AddRow(None)

        pref_lbl = forms.Label()
        pref_lbl.Text = "Ignore Prefix:"
        self.ignore_prefix_tb = forms.TextBox()
        self.ignore_prefix_tb.Text = _lp.DEFAULT_IGNORE_PREFIX
        self.ignore_prefix_tb.Width = 60
        pref_hint = forms.Label()
        pref_hint.Text = (
            "Layers whose name starts with this prefix are excluded together\n"
            "with their whole subtree (S3, S4, R1, R2). Empty = nothing ignored."
        )
        pref_hint.TextColor = _t.TEXT_MUTED
        layout.AddRow(pref_lbl, self.ignore_prefix_tb)
        layout.AddRow(None, pref_hint)
        layout.AddRow(None)

        tol_lbl = forms.Label()
        tol_lbl.Text = "Global Tolerance (%):"
        self.tolerance_stepper = forms.NumericStepper()
        self.tolerance_stepper.MinValue    = 0.0
        self.tolerance_stepper.MaxValue    = 50.0
        self.tolerance_stepper.Value       = 10.0
        self.tolerance_stepper.DecimalPlaces = 1
        self.tolerance_stepper.Increment   = 0.5
        self.tolerance_stepper.Width       = 80
        tol_hint = forms.Label()
        tol_hint.Text = "Symmetric tolerance applied in R1 and R2 bullet charts"
        tol_hint.TextColor = _t.TEXT_MUTED
        layout.AddRow(tol_lbl, self.tolerance_stepper)
        layout.AddRow(None, tol_hint)
        layout.AddRow(None)

        dec_lbl = forms.Label()
        dec_lbl.Text = "Decimal Places:"
        self.decimal_stepper = forms.NumericStepper()
        self.decimal_stepper.MinValue     = 0
        self.decimal_stepper.MaxValue     = 4
        self.decimal_stepper.Value        = 2
        self.decimal_stepper.DecimalPlaces = 0
        self.decimal_stepper.Increment    = 1
        self.decimal_stepper.Width        = 60
        dec_hint = forms.Label()
        dec_hint.Text = "Number of decimal places shown in all results (0 – 4)"
        dec_hint.TextColor = _t.TEXT_MUTED
        layout.AddRow(dec_lbl, self.decimal_stepper)
        layout.AddRow(None, dec_hint)
        layout.AddRow(None)

        zh_lbl = forms.Label()
        zh_lbl.Text = "Z-Height Tolerance:"
        self.z_height_tol_stepper = forms.NumericStepper()
        self.z_height_tol_stepper.MinValue      = 0.0
        self.z_height_tol_stepper.MaxValue      = 100.0
        self.z_height_tol_stepper.Value         = 0.5
        self.z_height_tol_stepper.DecimalPlaces = 2
        self.z_height_tol_stepper.Increment     = 0.1
        self.z_height_tol_stepper.Width         = 80
        zh_hint = forms.Label()
        zh_hint.Text = (
            "Min. Z gap (model units) to treat objects in one group as\n"
            "separate floors before merging overlaps. Applies to all scenarios."
        )
        zh_hint.TextColor = _t.TEXT_MUTED
        layout.AddRow(zh_lbl, self.z_height_tol_stepper)
        layout.AddRow(None, zh_hint)
        layout.AddRow(None)

        # ── Target Areas (R1 / R2) ────────────────────────────────────
        sep2 = forms.Label()
        sep2.Text = "─" * 42
        sep2.TextColor = _t.TEXT_MUTED
        layout.AddRow(sep2)

        sec3 = forms.Label()
        sec3.Text = "Target Areas (R1 / R2)"
        sec3.Font = _t.F_HEAD
        layout.AddRow(sec3)
        layout.AddRow(None)

        tgt_hint = forms.Label()
        tgt_hint.Text = (
            "Label must match the aggregation value — e.g. a category layer\n"
            "name ('Pflege'), a level name, or a user-text value.\n"
            "Values in model units²; decimal comma accepted."
        )
        tgt_hint.TextColor = _t.TEXT_MUTED
        layout.AddRow(None, tgt_hint)

        self._targets_layout = forms.StackLayout()
        self._targets_layout.Orientation = forms.Orientation.Vertical
        self._targets_layout.Spacing = 3
        layout.AddRow(None, self._targets_layout)

        add_tgt_btn = forms.Button()
        add_tgt_btn.Text = "+ Add Target"
        add_tgt_btn.Click += self._on_add_target

        fill_tgt_btn = forms.Button()
        fill_tgt_btn.Text = "Fill from last S3"
        fill_tgt_btn.Click += self._on_fill_targets_from_s3

        tgt_btn_row = forms.StackLayout()
        tgt_btn_row.Orientation = forms.Orientation.Horizontal
        tgt_btn_row.Spacing = 6
        tgt_btn_row.Items.Add(forms.StackLayoutItem(add_tgt_btn))
        tgt_btn_row.Items.Add(forms.StackLayoutItem(fill_tgt_btn))
        layout.AddRow(None, tgt_btn_row)
        layout.AddRow(None)

        # ── Configuration ────────────────────────────────────────────
        sep3 = forms.Label()
        sep3.Text = "─" * 42
        sep3.TextColor = _t.TEXT_MUTED
        layout.AddRow(sep3)

        sec4 = forms.Label()
        sec4.Text = "Configuration"
        sec4.Font = _t.F_HEAD
        layout.AddRow(sec4)
        layout.AddRow(None)

        save_btn = forms.Button()
        save_btn.Text = "Save Config"
        save_btn.Click += self.on_save_config

        load_btn = forms.Button()
        load_btn.Text = "Load Config"
        load_btn.Click += self.on_load_config

        cfg_hint = forms.Label()
        cfg_hint.Text = (
            "Settings persist automatically per model: stored in the 3dm as\n"
            "document user text when the window closes, restored on the next\n"
            "launch (save the model to keep them). Save / Load Config\n"
            "exchanges the same settings as a JSON file between models."
        )
        cfg_hint.TextColor = _t.TEXT_MUTED
        layout.AddRow(save_btn, load_btn)
        layout.AddRow(None, cfg_hint)

        scroll = forms.Scrollable()
        scroll.ExpandContentWidth = True
        scroll.ExpandContentHeight = False
        scroll.Content = layout
        page.Content = scroll
        return page

    # ------------------------------------------------------------------
    # Layer dropdown helpers
    # ------------------------------------------------------------------

    def _selected_layer(self, dd):
        text = (dd.Text or "").strip()
        return text if text in self.available_layers else None

    def _restore_layer_dd(self, dd, prev_name, updater):
        updater(self.available_layers)
        if not prev_name or prev_name not in self.available_layers:
            dd.Text = ""

    # ------------------------------------------------------------------
    # Settings helpers — ignore prefix, target table, R dimensions
    # ------------------------------------------------------------------

    def _ignore_prefix(self):
        """Current ignore prefix from Settings. Empty = nothing ignored."""
        return self.ignore_prefix_tb.Text.strip()

    def _add_target_row(self, label="", value=""):
        """Append one [label][value][✕] row to the Settings target table."""
        row = forms.StackLayout()
        row.Orientation = forms.Orientation.Horizontal
        row.Spacing = 4

        label_tb = forms.TextBox()
        label_tb.Text = label
        label_tb.PlaceholderText = "Label (category / level / key value)"
        label_tb.Width = 240

        value_tb = forms.TextBox()
        value_tb.Text = value
        value_tb.PlaceholderText = "Target"
        value_tb.Width = 90

        rm_btn = forms.Button()
        rm_btn.Text = "✕"
        rm_btn.Width = 28

        rec = {"label_tb": label_tb, "value_tb": value_tb, "row": row}

        def on_remove(s, e, r=rec):
            self._remove_target_row(r)
        rm_btn.Click += on_remove

        row.Items.Add(forms.StackLayoutItem(label_tb))
        row.Items.Add(forms.StackLayoutItem(value_tb))
        row.Items.Add(forms.StackLayoutItem(rm_btn))

        self._target_rows.append(rec)
        self._targets_layout.Items.Add(forms.StackLayoutItem(row))

    def _on_add_target(self, _s, _e):
        self._add_target_row()

    def _remove_target_row(self, rec):
        if rec in self._target_rows:
            self._target_rows.remove(rec)
        for i in range(self._targets_layout.Items.Count):
            if self._targets_layout.Items[i].Control is rec["row"]:
                self._targets_layout.Items.RemoveAt(i)
                break

    def _clear_target_rows(self):
        while self._targets_layout.Items.Count > 0:
            self._targets_layout.Items.RemoveAt(0)
        self._target_rows = []

    def _targets_dict(self):
        """(targets {label: float}, invalid_count) from the Settings table.
        Decimal commas accepted; blank labels skipped."""
        targets, invalid = {}, 0
        for rec in self._target_rows:
            label = rec["label_tb"].Text.strip()
            raw   = rec["value_tb"].Text.strip().replace(",", ".")
            if not label:
                continue
            try:
                targets[label] = float(raw)
            except (ValueError, TypeError):
                invalid += 1
        return targets, invalid

    def _on_fill_targets_from_s3(self, _s, _e):
        """Seed target rows from the categories of the last S3 result."""
        if not self._last_s3 or not self._last_s3.get("category_overall"):
            self.status_label.Text = (
                "Run Calculate on S3 first — categories are taken from its result."
            )
            self.status_label.TextColor = _t.TEXT_WARN
            return
        existing = {rec["label_tb"].Text.strip() for rec in self._target_rows}
        added = 0
        for cat in sorted(self._last_s3["category_overall"]):
            if cat != _lp.PLACEHOLDER and cat not in existing:
                self._add_target_row(cat, "")
                added += 1
        self.status_label.Text = (
            f"{added} categor{'y' if added == 1 else 'ies'} added — "
            "fill in the target values."
        )
        self.status_label.TextColor = _t.TEXT_OK if added else _t.TEXT_MUTED

    def _r_dim(self, dd, cb):
        """Dimension from an R-tab picker, or None if usertext with no key."""
        idx = dd.SelectedIndex
        if idx == 0:
            return (DIM_CATEGORY, None)
        if idx == 1:
            return (DIM_LEVEL, None)
        key = cb.Text.strip()
        return (DIM_USERTEXT, key) if key else None

    def _apply_r_dim(self, serial, dd, cb):
        """Restore an R-tab picker from its config serialization."""
        if not serial:
            return
        kind = serial[0]
        arg  = serial[1] if len(serial) > 1 else None
        if kind == DIM_CATEGORY:
            dd.SelectedIndex = 0
        elif kind == DIM_LEVEL:
            dd.SelectedIndex = 1
        else:
            dd.SelectedIndex = 2
            cb.Text = str(arg or "")
        cb.Visible = (dd.SelectedIndex == 2)

    # ------------------------------------------------------------------
    # Event handlers — model / navigation
    # ------------------------------------------------------------------

    def on_refresh_model(self, _sender, _e):
        prev_s2     = self._selected_layer(self.layer_s2_dd)
        prev_parent = self._selected_layer(self.parent_layer_dd)
        prev_s4_par = self._selected_layer(self.parent_layer_s4_dd)

        self.available_keys   = get_all_user_text_keys()
        self.available_layers = all_layer_names()

        for upd in self._ks_combos:
            upd(self.available_keys)

        for upd in self._ks_s4.values():
            upd(self.available_keys)

        self._restore_layer_dd(self.layer_s2_dd,        prev_s2,     self._ks_layer_s2)
        self._restore_layer_dd(self.parent_layer_dd,    prev_parent, self._ks_layer_parent)
        self._restore_layer_dd(self.parent_layer_s4_dd, prev_s4_par, self._ks_layer_s4_par)

        self._ks_write(["Area"] + list(self.available_keys))

        self.status_label.Text = (
            f"Refreshed  —  {len(self.available_layers)} layer(s), "
            f"{len(self.available_keys)} key(s)."
        )
        self.status_label.TextColor = _t.TEXT_MUTED

    def on_calculate(self, _sender, _e):
        global _DECIMALS
        _DECIMALS = int(self.decimal_stepper.Value)
        unit = unit_label()
        try:
            idx = self.tabs.SelectedIndex
            runners = {
                0: self._run_s1,
                1: self._run_s2,
                2: self._run_s3,
                3: self._run_s4,
                4: self._run_r1,
                5: self._run_r2,
            }
            fn = runners.get(idx)
            if fn:
                fn(unit)
            else:
                self.status_label.Text = "Switch to a calculation tab (S1–S4, R1–R2) to calculate."
                self.status_label.TextColor = _t.TEXT_MUTED
        except Exception as ex:
            self.status_label.Text = f"Error: {ex}"
            self.status_label.TextColor = _t.TEXT_ERROR

    def on_clear(self, _sender, _e):
        idx = self.tabs.SelectedIndex
        if idx == 0:
            self.results_s1_grid.DataStore = forms.TreeGridItemCollection()
        elif idx == 1:
            self.results_s2_grid.DataStore = forms.TreeGridItemCollection()
        elif idx == 2:
            self.results_s3_breakdown_grid.DataStore = forms.TreeGridItemCollection()
            self.results_s3_objects_grid.DataStore   = forms.TreeGridItemCollection()
            self._last_s3 = None
        elif idx == 3:
            self.results_s4_grid.DataStore = forms.TreeGridItemCollection()
            self._last_s4 = None
        elif idx == 4:
            self._r1_entries = []
            self.warn_r1.Visible = False
            self.chart_r1.Invalidate()
        elif idx == 5:
            self._r2_entries = []
            self.warn_r2.Visible = False
            self.chart_r2.Invalidate()
        self.status_label.Text = "Cleared."
        self.status_label.TextColor = _t.TEXT_MUTED

    def on_export(self, _sender, _e):
        if self._export_data is None:
            self.status_label.Text = "Nothing to export — run Calculate on S1, S2, S3, or S4 first."
            self.status_label.TextColor = _t.TEXT_WARN
            return

        path = rs.SaveFileName(
            "Export Results to Excel",
            "Excel Files (*.xlsx)|*.xlsx||",
            _prefs_get('lindero_export_xlsx'), "Lindero_Results", "xlsx"
        )
        if not path:
            return
        _prefs_set('lindero_export_xlsx', path)
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"

        try:
            export_to_excel(self._export_data, path)
            self.status_label.Text = f"Exported → {path}"
            self.status_label.TextColor = _t.TEXT_OK
        except Exception as ex:
            self.status_label.Text = f"Export failed: {ex}"
            self.status_label.TextColor = _t.TEXT_ERROR

    def on_export_png(self, _sender, _e):
        idx = self.tabs.SelectedIndex
        if idx == 4:
            entries, tol, unit = self._r1_entries, self._r1_tol, self._r1_unit
            default_name = "Lindero_R1_RoomAnalysis"
        elif idx == 5:
            entries, tol, unit = self._r2_entries, self._r2_tol, self._r2_unit
            default_name = "Lindero_R2_GroupAnalysis"
        else:
            self.status_label.Text = "Switch to R1 or R2 tab to export a chart as PNG."
            self.status_label.TextColor = _t.TEXT_WARN
            return

        if not entries:
            self.status_label.Text = "No chart data — run Calculate first."
            self.status_label.TextColor = _t.TEXT_WARN
            return

        path = rs.SaveFileName(
            "Export Chart as PNG",
            "PNG Images (*.png)|*.png||",
            _prefs_get('lindero_export_png'), default_name, "png"
        )
        if not path:
            return
        _prefs_set('lindero_export_png', path)
        if not path.lower().endswith(".png"):
            path += ".png"

        try:
            _export_chart_png(entries, tol, unit, path)
            self.status_label.Text = f"Chart exported → {path}"
            self.status_label.TextColor = _t.TEXT_OK
        except Exception as ex:
            self.status_label.Text = f"PNG export failed: {ex}"
            self.status_label.TextColor = _t.TEXT_ERROR

    def on_copy_window(self, _sender, _e):
        """
        Screenshot this window (incl. title bar) onto the Windows clipboard,
        ready to paste into chat/mail — replaces the Snagit round-trip.

        The window rectangle is read from Win32 in PHYSICAL pixels
        (DwmGetWindowAttribute EXTENDED_FRAME_BOUNDS, GetWindowRect as
        fallback) — the same virtual-desktop coordinate space
        CopyFromScreen captures in, so no DPI math is involved.
        Do NOT derive the rect from Eto's logical Bounds × LogicalPixelSize:
        with monitors of different scale factors the logical→physical
        mapping shifts per screen, which grabbed the wrong monitor
        (bug 02.07.2026). Windows-only by design (WinForms clipboard).
        """
        try:
            import ctypes
            from ctypes import wintypes
            import clr
            clr.AddReference("System.Drawing")
            clr.AddReference("System.Windows.Forms")
            import System.Drawing as _sd
            import System.Windows.Forms as _swf

            try:
                hwnd = self.NativeHandle.ToInt64()
            except AttributeError:
                hwnd = int(self.NativeHandle)
            if not hwnd:
                raise RuntimeError("no native window handle")

            # Visually tight window rect (excludes the invisible resize
            # border Win10/11 adds around GetWindowRect).
            rect = wintypes.RECT()
            got = False
            try:
                DWMWA_EXTENDED_FRAME_BOUNDS = 9
                res = ctypes.windll.dwmapi.DwmGetWindowAttribute(
                    wintypes.HWND(hwnd), DWMWA_EXTENDED_FRAME_BOUNDS,
                    ctypes.byref(rect), ctypes.sizeof(rect))
                got = (res == 0)
            except Exception:
                got = False
            if not got:
                if not ctypes.windll.user32.GetWindowRect(
                        wintypes.HWND(hwnd), ctypes.byref(rect)):
                    raise RuntimeError("GetWindowRect failed")

            x = rect.left
            y = rect.top
            w = max(1, rect.right - rect.left)
            h = max(1, rect.bottom - rect.top)

            bmp = _sd.Bitmap(w, h)
            g = _sd.Graphics.FromImage(bmp)
            try:
                g.CopyFromScreen(x, y, 0, 0, _sd.Size(w, h))
            finally:
                g.Dispose()
            _swf.Clipboard.SetImage(bmp)

            self.status_label.Text = (
                "Window copied to clipboard — paste with Ctrl+V.")
            self.status_label.TextColor = _t.TEXT_OK
        except Exception as ex:
            self.status_label.Text = f"Copy failed: {ex}"
            self.status_label.TextColor = _t.TEXT_ERROR

    # ------------------------------------------------------------------
    # Event handlers — Write Area
    # ------------------------------------------------------------------

    def on_write_area_toggle(self, _s, _e):
        self._write_panel.Visible = not self._write_panel.Visible

    def _on_write_cancel(self, _s, _e):
        self._write_panel.Visible = False

    def on_confirm_write(self, _s, _e):
        key = self._write_key_combo.Text.strip()
        if not key:
            self.status_label.Text = "Enter a key name to write to."
            self.status_label.TextColor = _t.TEXT_ERROR
            return

        idx = self.tabs.SelectedIndex
        objects = []
        if idx == 0 and self._last_s1:
            objects = self._last_s1
        elif idx == 1 and self._last_s2:
            objects = self._last_s2["objects"]
        elif idx == 2 and self._last_s3:
            for lv in self._last_s3["levels"].values():
                objects += lv["objects"]
        elif idx == 3 and self._last_s4:
            objects = self._last_s4.get("objects") or []

        if not objects:
            self.status_label.Text = "No calculated data — run Calculate first on S1, S2, S3, or S4."
            self.status_label.TextColor = _t.TEXT_WARN
            return

        count = 0
        dp = int(self.decimal_stepper.Value)  # honor Settings decimal places
        for obj in objects:
            try:
                guid   = System.Guid(obj["guid"])
                rh_obj = sc.doc.Objects.FindId(guid)
                if rh_obj:
                    rh_obj.Attributes.SetUserString(key, f"{obj['area']:.{dp}f}")
                    rh_obj.CommitChanges()
                    count += 1
            except Exception:
                pass

        sc.doc.Views.Redraw()
        self._write_panel.Visible = False
        self.status_label.Text = f"Area written to {count} object(s) using key '{key}'"
        self.status_label.TextColor = _t.TEXT_OK

    # ------------------------------------------------------------------
    # Event handlers — Settings config
    # ------------------------------------------------------------------

    def _gather_config(self):
        """Current settings as a config-v2 dict — shared by the JSON export
        and the per-document auto-persist."""
        def _dim_serial(dd, cb):
            idx = dd.SelectedIndex
            if idx == 0:
                return [DIM_CATEGORY, None]
            if idx == 1:
                return [DIM_LEVEL, None]
            return [DIM_USERTEXT, cb.Text.strip()]

        return {
            "config_version":    2,
            "tolerance_percent": self.tolerance_stepper.Value,
            "decimal_places":    int(self.decimal_stepper.Value),
            "z_height_tol":      self.z_height_tol_stepper.Value,
            "ignore_prefix":     self._ignore_prefix(),
            "s3_parent_layer":   self._selected_layer(self.parent_layer_dd) or "",
            "s4_parent_layer":   self._selected_layer(self.parent_layer_s4_dd) or "",
            "s4_dimensions":     [[k, a] for k, a in self._s4_dims()],
            "r1_dimension":      _dim_serial(self.r1_dim_dd, self.r1_key_cb),
            "r2_dimension":      _dim_serial(self.r2_dim_dd, self.r2_key_cb),
            "targets":           self._targets_dict()[0],
        }

    def _apply_config(self, cfg):
        """Apply a config dict (v1 or v2) to all controls. Returns version."""
        version = int(cfg.get("config_version", 1))

        self.tolerance_stepper.Value    = float(cfg.get("tolerance_percent", 10.0))
        self.decimal_stepper.Value      = float(cfg.get("decimal_places", 2))
        self.z_height_tol_stepper.Value = float(cfg.get("z_height_tol", 0.5))
        self.ignore_prefix_tb.Text      = cfg.get(
            "ignore_prefix", _lp.DEFAULT_IGNORE_PREFIX)

        s3_par = cfg.get("s3_parent_layer", "")
        if s3_par and s3_par in self.available_layers:
            self.parent_layer_dd.Text = s3_par
        s4_par = cfg.get("s4_parent_layer", "")
        if s4_par and s4_par in self.available_layers:
            self.parent_layer_s4_dd.Text = s4_par

        if version >= 2:
            dims = cfg.get("s4_dimensions") or []
            if dims:
                self._s4_clear_dim_rows()
                for d in dims:
                    kind = d[0] if len(d) > 0 else DIM_USERTEXT
                    arg  = d[1] if len(d) > 1 else None
                    self._s4_add_dim_row(kind, arg)
                if not self._s4_dim_rows:
                    self._s4_add_dim_row()
            self._apply_r_dim(cfg.get("r1_dimension"),
                              self.r1_dim_dd, self.r1_key_cb)
            self._apply_r_dim(cfg.get("r2_dimension"),
                              self.r2_dim_dd, self.r2_key_cb)
            targets = cfg.get("targets") or {}
            if targets:
                self._clear_target_rows()
                for label in sorted(targets):
                    self._add_target_row(label, str(targets[label]))
        else:
            # v1 migration: key sequence becomes UserText dimensions;
            # target keys and the old data-source fields are obsolete.
            key_seq = [k for k in (cfg.get("s4_key_sequence") or []) if k]
            if key_seq:
                self._s4_clear_dim_rows()
                for k in key_seq:
                    self._s4_add_dim_row(DIM_USERTEXT, k)
        return version

    # Per-document persistence -----------------------------------------
    # The current settings live in the 3dm (document user text) so each
    # model remembers its own setup: written when the window closes and
    # whenever a config is saved/loaded, restored at startup. They stay
    # until a new config replaces them. NOTE: the 3dm must be saved for
    # the settings to survive — they ride inside the model file.

    def _store_doc_config(self):
        try:
            rs.SetDocumentUserText(
                DOC_CONFIG_KEY,
                json.dumps(self._gather_config(), ensure_ascii=False))
        except Exception:
            pass  # never block closing on a persistence hiccup

    def _load_doc_config(self):
        """Restore settings persisted in the document. True if applied."""
        try:
            raw = rs.GetDocumentUserText(DOC_CONFIG_KEY)
            if raw:
                self._apply_config(json.loads(raw))
                return True
        except Exception:
            pass
        return False

    def on_form_closed(self, _s, _e):
        self._store_doc_config()

    def on_save_config(self, _s, _e):
        path = rs.SaveFileName(
            "Save Lindero Configuration",
            "JSON Files (*.json)|*.json||",
            _prefs_get('lindero_config'), "lindero_config", "json"
        )
        if not path:
            return
        _prefs_set('lindero_config', path)
        if not path.lower().endswith(".json"):
            path += ".json"
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self._gather_config(), f, indent=2, ensure_ascii=False)
            self._store_doc_config()
            self.status_label.Text = f"Config saved → {path}"
            self.status_label.TextColor = _t.TEXT_OK
        except Exception as ex:
            self.status_label.Text = f"Save failed: {ex}"
            self.status_label.TextColor = _t.TEXT_ERROR

    def on_load_config(self, _s, _e):
        path = rs.OpenFileName(
            "Load Lindero Configuration",
            "JSON Files (*.json)|*.json||",
            folder=_prefs_get('lindero_config')
        )
        if not path:
            return
        _prefs_set('lindero_config', path)
        try:
            with open(path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            version = self._apply_config(cfg)
            # The loaded config becomes the persistent one for this model.
            self._store_doc_config()
            if version >= 2:
                self.status_label.Text = f"Config loaded ← {path}"
                self.status_label.TextColor = _t.TEXT_OK
            else:
                self.status_label.Text = (
                    f"v1 config loaded ← {path}  —  target keys / data-source "
                    "fields are obsolete; define Target Areas in Settings."
                )
                self.status_label.TextColor = _t.TEXT_WARN
        except Exception as ex:
            self.status_label.Text = f"Load failed: {ex}"
            self.status_label.TextColor = _t.TEXT_ERROR

    # ------------------------------------------------------------------
    # Per-scenario runners — S1 / S2 / S3
    # ------------------------------------------------------------------

    def _run_s1(self, unit):
        name_key = self.name_key_combo.Text.strip()
        data = calc_s1(name_key, float(self.z_height_tol_stepper.Value))

        if not data["objects"]:
            self.results_s1_grid.DataStore = forms.TreeGridItemCollection()
            self.status_label.Text = "No selection."
            self.status_label.TextColor = _t.TEXT_WARN
            return

        self._populate_flat_grid(self.results_s1_grid, data)
        self._last_s1 = data["objects"]
        self._export_data = {
            "scenario": 1, "unit": unit,
            "params": {"name_key": name_key},
            "objects": data["objects"],
        }
        self.status_label.Text = (
            f"S1  —  {len(data['objects'])} object(s)  |  Total: {_fmt(data['total'])} {unit}"
        )
        self.status_label.TextColor = _t.TEXT_OK

    def _run_s2(self, unit):
        layer_name = self._selected_layer(self.layer_s2_dd)
        if not layer_name:
            self.status_label.Text = "Please select a layer."
            self.status_label.TextColor = _t.TEXT_ERROR
            return

        obj_key = self.obj_key_s2.Text.strip()
        data    = calc_s2(layer_name, obj_key, float(self.z_height_tol_stepper.Value))

        if not data["objects"]:
            self.results_s2_grid.DataStore = forms.TreeGridItemCollection()
            self.status_label.Text = f"Layer '{short_name(layer_name)}' is empty."
            self.status_label.TextColor = _t.TEXT_WARN
            return

        self._populate_flat_grid(self.results_s2_grid, data)
        self._last_s2 = data
        self._export_data = {
            "scenario": 2, "unit": unit,
            "params": {"layer_name": layer_name, "obj_key": obj_key},
            **data,
        }
        self.status_label.Text = (
            f"S2  —  '{short_name(layer_name)}'  |  "
            f"{len(data['objects'])} object(s)  |  Total: {_fmt(data['total'])} {unit}"
        )
        self.status_label.TextColor = _t.TEXT_OK

    def _run_s3(self, unit):
        parent = self._selected_layer(self.parent_layer_dd)
        if not parent:
            self.status_label.Text = "Please select a parent layer."
            self.status_label.TextColor = _t.TEXT_ERROR
            return

        obj_key = self.obj_key_s3.Text.strip()
        prefix  = self._ignore_prefix()
        data    = calc_s3(parent, obj_key, prefix,
                          float(self.z_height_tol_stepper.Value))

        if not any(d["objects"] for d in data["levels"].values()):
            self.results_s3_breakdown_grid.DataStore = forms.TreeGridItemCollection()
            self.results_s3_objects_grid.DataStore   = forms.TreeGridItemCollection()
            self.status_label.Text = (
                f"No measurable objects under '{short_name(parent)}' "
                f"(ignore prefix '{prefix}') — try Preview Structure."
            )
            self.status_label.TextColor = _t.TEXT_WARN
            return

        self._populate_s3_breakdown_grid(data)
        self._populate_s3_objects_grid(data)
        self._last_s3 = data
        self._export_data = {
            "scenario": 3, "unit": unit,
            "params": {"parent": parent, "obj_key": obj_key,
                       "ignore_prefix": prefix},
            "levels": data["levels"],
            "overall_total": data["overall_total"],
            "category_overall": data["category_overall"],
            "layer_info": data["layer_info"],
            "warnings": data["warnings"],
        }
        n_warn = sum(len(d["warnings"]) for d in data["levels"].values())
        n_warn += len(data["warnings"])
        self.status_label.Text = (
            f"S3  —  '{short_name(parent)}'  |  "
            f"{len(data['levels'])} level(s)  |  "
            f"Overall total: {_fmt(data['overall_total'])} {unit}"
            + (f"  |  {n_warn} warning(s)" if n_warn else "")
        )
        self.status_label.TextColor = _t.TEXT_WARN if n_warn else _t.TEXT_OK

    def on_preview_s3(self, _s, _e):
        """Dry-run preview: layer roles + object counts, no area math."""
        parent = self._selected_layer(self.parent_layer_dd)
        if not parent:
            self.status_label.Text = "Please select a parent layer."
            self.status_label.TextColor = _t.TEXT_ERROR
            return
        prefix = self._ignore_prefix()
        pv = preview_hierarchy(parent, prefix)
        self._populate_s3_preview_grid(pv, prefix)
        self.status_label.Text = (
            f"Preview  —  {len(pv['tree'])} level(s), "
            f"{pv['total_objects']} object(s) included, "
            f"{len(pv['ignored'])} layer(s) ignored. Nothing calculated yet."
        )
        self.status_label.TextColor = _t.TEXT_MUTED

    # ------------------------------------------------------------------
    # S4 grid helpers
    # ------------------------------------------------------------------

    def _on_cell_format(self, _, e):
        """
        Shared CellFormatting handler for all result grids.
        Row type is stored in the last Values slot:
          0/1  data row even/odd  → white / ROW_ALT stripe
          10   section header     → BTN_DEFAULT + bold
          20   subtotal           → BTN_DEFAULT
          30   grand total        → TOTAL_BG + bold
          40   warning/info       → TEXT_WARN foreground
        """
        item = e.Item
        if not isinstance(item, forms.TreeGridItem):
            return
        vals = item.Values
        if vals is None or len(vals) == 0:
            return
        try:
            rt = int(vals[len(vals) - 1])
        except (TypeError, ValueError):
            return
        if rt == 1:
            e.BackgroundColor = _t.ROW_ALT
        elif rt == 10:
            e.BackgroundColor = _t.BTN_DEFAULT
            e.Font = _t.F_MONO_B
        elif rt == 15:
            e.BackgroundColor = _t.BTN_DEFAULT
        elif rt == 20:
            e.BackgroundColor = _t.BTN_DEFAULT
        elif rt == 30:
            e.BackgroundColor = _t.TOTAL_BG
            e.Font = _t.F_MONO_B
        elif rt == 40:
            e.ForegroundColor = _t.TEXT_WARN

    def _populate_s4_grid(self, data):
        """Build TreeGridItemCollection from calc_s4 data and assign to grid."""
        collection = forms.TreeGridItemCollection()

        def add_nodes(parent_col, node_dict):
            for val in sorted(node_dict):
                entry = node_dict[val]
                has_children = bool(entry["children"])
                item = forms.TreeGridItem()
                item.Values = [val, _fmt(entry["area"]).rjust(10), 10 if has_children else 0]
                item.Expanded = True
                if has_children:
                    add_nodes(item.Children, entry["children"])
                parent_col.Add(item)

        add_nodes(collection, data["tree"])

        total_item = forms.TreeGridItem()
        total_item.Values = ["OVERALL TOTAL", _fmt(data["overall_total"]).rjust(10), 30]
        collection.Add(total_item)

        self.results_s4_grid.DataStore = collection

    # ------------------------------------------------------------------
    # Shared grid factory
    # ------------------------------------------------------------------

    def _make_results_grid(self, col_specs):
        """
        Return a styled TreeGridView.
        col_specs: list of (expand: bool, width_px: int) — one per data column.
        The last Values slot is always the hidden type-code; no column is added for it.
        """
        grid = forms.TreeGridView()
        grid.ShowHeader = False
        grid.AllowColumnReordering = False
        grid.AllowMultipleSelection = False
        grid.Font = _t.F_MONO
        for i, (expand, width) in enumerate(col_specs):
            col = forms.GridColumn()
            col.Editable = False
            col.DataCell = forms.TextBoxCell(i)
            if expand:
                col.Expand = True
            else:
                col.Width = width
            grid.Columns.Add(col)
        grid.CellFormatting += self._on_cell_format
        return grid

    # ------------------------------------------------------------------
    # Grid populate helpers — S1 / S2
    # ------------------------------------------------------------------

    def _populate_flat_grid(self, grid, data):
        """
        Shared populator for S1 and S2: flat object list + sum + total + warnings.
        data must have: objects [{name, area}], total, union_ok, skipped.
        S1 may also carry: layer_totals [{layer, area}], z_warning (str or None).
        """
        collection = forms.TreeGridItemCollection()
        raw_sum = sum(o["area"] for o in data["objects"])
        for i, o in enumerate(data["objects"]):
            item = forms.TreeGridItem()
            item.Values = [o["name"], _fmt(o["area"]).rjust(10), i % 2]
            collection.Add(item)
        sum_item = forms.TreeGridItem()
        sum_item.Values = ["Sum (individual totals)", _fmt(raw_sum).rjust(10), 20]
        collection.Add(sum_item)

        layer_totals = data.get("layer_totals") or []
        multi_layer = len(layer_totals) > 1
        if multi_layer:
            for lt in layer_totals:
                lt_item = forms.TreeGridItem()
                lt_item.Values = [f"  {short_name(lt['layer'])}", _fmt(lt["area"]).rjust(10), 15]
                collection.Add(lt_item)

        union_note = "  [union failed — sum shown]" if not data["union_ok"] else ""
        total_label = "TOTAL (per-layer merged)" if multi_layer else "TOTAL (overlaps merged)"
        total_item = forms.TreeGridItem()
        total_item.Values = [f"{total_label}{union_note}", _fmt(data["total"]).rjust(10), 30]
        collection.Add(total_item)

        overlap = raw_sum - data["total"]
        if overlap > 1e-6 and not multi_layer:
            w = forms.TreeGridItem()
            w.Values = [f"  Overlap: {_fmt(overlap)} — some objects share footprint area", "", 40]
            collection.Add(w)
        z_warning = data.get("z_warning")
        if z_warning:
            w = forms.TreeGridItem()
            w.Values = [f"  [!] {z_warning}", "", 40]
            collection.Add(w)
        if data.get("skipped", 0) > 0:
            w = forms.TreeGridItem()
            w.Values = [f"  {data['skipped']} object(s) skipped (no calculable footprint)", "", 40]
            collection.Add(w)
        grid.DataStore = collection

    # ------------------------------------------------------------------
    # Grid populate helpers — S3 (breakdown panel)
    # ------------------------------------------------------------------

    def _populate_s3_breakdown_grid(self, data):
        """Level rows with category children, building-wide category summary,
        overall total, and global warnings."""
        collection = forms.TreeGridItemCollection()
        for level, lv in data["levels"].items():
            union_note = "  [union failed]" if not lv["union_ok"] else ""
            level_item = forms.TreeGridItem()
            level_item.Values = [
                f"▸ {level}{union_note}",
                _fmt(lv["total"]).rjust(10),
                10,
            ]
            level_item.Expanded = True
            for i, (cat, ca) in enumerate(sorted(lv["category_totals"].items())):
                child = forms.TreeGridItem()
                child.Values = [f"  {cat}", _fmt(ca).rjust(10), i % 2]
                level_item.Children.Add(child)
            for w_text in lv["warnings"]:
                w = forms.TreeGridItem()
                w.Values = [f"  [!] {w_text}", "", 40]
                level_item.Children.Add(w)
            collection.Add(level_item)

        if data["category_overall"]:
            cat_hdr = forms.TreeGridItem()
            cat_hdr.Values = ["▸ CATEGORIES — whole building", "", 10]
            cat_hdr.Expanded = True
            for i, (cat, ca) in enumerate(sorted(data["category_overall"].items())):
                child = forms.TreeGridItem()
                child.Values = [f"  {cat}", _fmt(ca).rjust(10), i % 2]
                cat_hdr.Children.Add(child)
            collection.Add(cat_hdr)

        total_item = forms.TreeGridItem()
        total_item.Values = ["OVERALL TOTAL (sum of levels)",
                             _fmt(data["overall_total"]).rjust(10), 30]
        collection.Add(total_item)

        for w_text in data.get("warnings", []):
            w = forms.TreeGridItem()
            w.Values = [f"  {w_text}", "", 40]
            collection.Add(w)
        self.results_s3_breakdown_grid.DataStore = collection

    # ------------------------------------------------------------------
    # Grid populate helpers — S3 (objects panel)
    # ------------------------------------------------------------------

    def _populate_s3_objects_grid(self, data):
        """Level header rows with object children (name + category + area)."""
        collection = forms.TreeGridItemCollection()
        for level, lv in data["levels"].items():
            level_item = forms.TreeGridItem()
            level_item.Values = [f"▸ {level}", "", "", 10]
            level_item.Expanded = True
            for i, o in enumerate(lv["objects"]):
                child = forms.TreeGridItem()
                child.Values = [o["name"], o["category"],
                                _fmt(o["area"]).rjust(10), i % 2]
                level_item.Children.Add(child)
            individual_sum = sum(o["area"] for o in lv["objects"])
            overlap = individual_sum - lv["total"]
            if overlap > 1e-6:
                w = forms.TreeGridItem()
                w.Values = [f"  Overlap: {_fmt(overlap)}", "", "", 40]
                level_item.Children.Add(w)
            if lv.get("skipped", 0) > 0:
                w = forms.TreeGridItem()
                w.Values = [f"  {lv['skipped']} object(s) skipped", "", "", 40]
                level_item.Children.Add(w)
            collection.Add(level_item)
        self.results_s3_objects_grid.DataStore = collection

    # ------------------------------------------------------------------
    # Grid populate helpers — S3 (structure preview)
    # ------------------------------------------------------------------

    def _populate_s3_preview_grid(self, pv, prefix):
        """Dry-run view in the breakdown panel: counts, not areas."""
        collection = forms.TreeGridItemCollection()
        hdr = forms.TreeGridItem()
        hdr.Values = ["PREVIEW — object counts (no areas computed)", "", 10]
        collection.Add(hdr)

        for level in sorted(pv["tree"], key=_lp.level_sort_key):
            cats = pv["tree"][level]
            level_item = forms.TreeGridItem()
            level_item.Values = [f"▸ {level}",
                                 str(sum(cats.values())).rjust(10), 10]
            level_item.Expanded = True
            for i, cat in enumerate(sorted(cats)):
                child = forms.TreeGridItem()
                child.Values = [f"  {cat}", str(cats[cat]).rjust(10), i % 2]
                level_item.Children.Add(child)
            collection.Add(level_item)

        if pv["ignored"]:
            ign_item = forms.TreeGridItem()
            ign_item.Values = [f"▸ IGNORED (prefix '{prefix}')",
                               str(len(pv["ignored"])).rjust(10), 10]
            ign_item.Expanded = True
            for i, lp_full in enumerate(pv["ignored"]):
                child = forms.TreeGridItem()
                child.Values = [f"  {lp_full}", "", 40]
                ign_item.Children.Add(child)
            collection.Add(ign_item)

        if pv["parent_direct_objects"]:
            w = forms.TreeGridItem()
            w.Values = [
                f"  [!] {pv['parent_direct_objects']} object(s) directly on "
                "the parent layer — not measured", "", 40]
            collection.Add(w)
        self.results_s3_breakdown_grid.DataStore = collection

    # ------------------------------------------------------------------
    # Per-scenario runner — S4
    # ------------------------------------------------------------------

    def _run_s4(self, unit):
        parent = self._selected_layer(self.parent_layer_s4_dd)
        if not parent:
            self.status_label.Text = "Please select a parent layer on the S4 tab."
            self.status_label.TextColor = _t.TEXT_ERROR
            return

        dims = self._s4_dims()
        if not dims:
            self.status_label.Text = "S4 requires at least one dimension."
            self.status_label.TextColor = _t.TEXT_WARN
            return

        prefix = self._ignore_prefix()
        data = calc_s4(parent, dims, prefix,
                       float(self.z_height_tol_stepper.Value))

        if not data["tree"]:
            self.results_s4_grid.DataStore = forms.TreeGridItemCollection()
            self.status_label.Text = (
                f"No measurable objects under '{short_name(parent)}' "
                f"(ignore prefix '{prefix}')."
            )
            self.status_label.TextColor = _t.TEXT_WARN
            return

        self._populate_s4_grid(data)
        self._last_s4 = data
        self._export_data = {
            "scenario": 4, "unit": unit,
            "params": {"parent": parent, "dims": dims,
                       "dim_labels": [_dim_label(d) for d in dims],
                       "ignore_prefix": prefix},
            "tree": data["tree"],
            "overall_total": data["overall_total"],
        }
        n_warn = len(data["warnings"])
        self.status_label.Text = (
            f"S4  —  '{short_name(parent)}'  |  "
            + " > ".join(_dim_label(d) for d in dims)
            + f"  |  Overall total: {_fmt(data['overall_total'])} {unit}"
            + (f"  |  {n_warn} warning(s)" if n_warn else "")
        )
        self.status_label.TextColor = (
            _t.TEXT_WARN if n_warn else _t.TEXT_OK
        )

    # ------------------------------------------------------------------
    # Per-scenario runners — R1 / R2
    # ------------------------------------------------------------------

    def _run_r(self, unit, tag, dim_dd, key_cb, chart, warn_lbl, store):
        """Shared R1/R2 runner. store: ('_r1'|'_r2') state prefix."""
        parent = self._selected_layer(self.parent_layer_dd)
        if not parent:
            self.status_label.Text = (
                f"{tag} uses the S3 Parent Layer — select one on the S3 tab.")
            self.status_label.TextColor = _t.TEXT_ERROR
            return

        dim = self._r_dim(dim_dd, key_cb)
        if dim is None:
            self.status_label.Text = (
                f"{tag}: choose a user text key (or another dimension).")
            self.status_label.TextColor = _t.TEXT_WARN
            return

        targets, n_invalid = self._targets_dict()
        tol  = self.tolerance_stepper.Value / 100.0
        data = calc_r(parent, dim, targets, self._ignore_prefix(),
                      float(self.z_height_tol_stepper.Value))

        if not data["entries"]:
            self.status_label.Text = (
                "No measurable objects found — check the S3 parent layer "
                "and the ignore prefix.")
            self.status_label.TextColor = _t.TEXT_WARN
            return

        warnings = list(data["warnings"])
        if n_invalid:
            warnings.append(
                f"[!] {n_invalid} target value(s) are not numbers — ignored.")

        setattr(self, store + "_entries", data["entries"])
        setattr(self, store + "_tol", tol)
        setattr(self, store + "_unit", unit)
        n = len(data["entries"])
        chart.Size = drawing.Size(
            max(400, chart.Width),
            max(10, n * _CHART_ROW_H + 20)
        )
        chart.Invalidate()

        if warnings:
            warn_lbl.Text    = "\n".join(warnings)
            warn_lbl.Visible = True
        else:
            warn_lbl.Visible = False

        self.status_label.Text = (
            f"{tag}  [{_dim_label(dim)}]  —  {n} entr{'y' if n == 1 else 'ies'}"
            f"  |  Tolerance: {tol*100:.1f}%"
            + (f"  |  {len(warnings)} warning(s)" if warnings else "")
        )
        self.status_label.TextColor = (
            _t.TEXT_WARN if warnings else _t.TEXT_OK
        )

    def _run_r1(self, unit):
        self._run_r(unit, "R1", self.r1_dim_dd, self.r1_key_cb,
                    self.chart_r1, self.warn_r1, "_r1")

    def _run_r2(self, unit):
        self._run_r(unit, "R2", self.r2_dim_dd, self.r2_key_cb,
                    self.chart_r2, self.warn_r2, "_r2")

    # ------------------------------------------------------------------
    # Bullet chart paint handlers
    # ------------------------------------------------------------------

    def _paint_r1(self, sender, e):
        g = e.Graphics
        w = sender.Width
        for i, entry in enumerate(self._r1_entries):
            _draw_bullet_row(g, i, entry, self._r1_tol, self._r1_unit, w)

    def _paint_r2(self, sender, e):
        g = e.Graphics
        w = sender.Width
        for i, entry in enumerate(self._r2_entries):
            _draw_bullet_row(g, i, entry, self._r2_tol, self._r2_unit, w)


# ══════════════════════════════════════════════════════════════════════════════
# Excel export
# ══════════════════════════════════════════════════════════════════════════════

# Styles -------------------------------------------------------------------

_HDR_FONT  = Font(bold=True, color="FFFFFF")
_HDR_FILL  = PatternFill(fill_type="solid", fgColor="2F5496")   # dark blue
_HDR_ALIGN = Alignment(horizontal="center", vertical="center")

_SEC_FONT  = Font(bold=True, color="1F3864")
_SEC_FILL  = PatternFill(fill_type="solid", fgColor="D9E1F2")   # light blue

_TOT_FONT  = Font(bold=True)
_TOT_FILL  = PatternFill(fill_type="solid", fgColor="F2F2F2")   # light grey

_WARN_FONT = Font(bold=True, color="7F6000")
_WARN_FILL = PatternFill(fill_type="solid", fgColor="FFE699")   # amber

_AREA_FMT  = "#,##0.0000"


def _hdr(ws, row, cols):
    """Write a styled header row. Returns the row index + 1."""
    for ci, text in enumerate(cols, 1):
        c = ws.cell(row=row, column=ci, value=text)
        c.font      = _HDR_FONT
        c.fill      = _HDR_FILL
        c.alignment = _HDR_ALIGN
    return row + 1


def _sec(ws, row, text, n_cols=2):
    """Write a section-label row."""
    c = ws.cell(row=row, column=1, value=text)
    c.font = _SEC_FONT
    c.fill = _SEC_FILL
    for ci in range(2, n_cols + 1):
        ws.cell(row=row, column=ci).fill = _SEC_FILL
    return row + 1


def _tot(ws, row, label, value):
    """Write a bold total row with area formatting."""
    lc = ws.cell(row=row, column=1, value=label)
    lc.font = _TOT_FONT
    lc.fill = _TOT_FILL
    vc = ws.cell(row=row, column=2, value=value)
    vc.font          = _TOT_FONT
    vc.fill          = _TOT_FILL
    vc.number_format = _AREA_FMT
    return row + 1


def _warn(ws, row, label, value=None, n_cols=2):
    """Write an amber-highlighted warning row."""
    lc = ws.cell(row=row, column=1, value=label)
    lc.font = _WARN_FONT
    lc.fill = _WARN_FILL
    for ci in range(2, n_cols + 1):
        ws.cell(row=row, column=ci).fill = _WARN_FILL
    if value is not None:
        vc = ws.cell(row=row, column=2, value=value)
        vc.font          = _WARN_FONT
        vc.fill          = _WARN_FILL
        vc.number_format = _AREA_FMT
    return row + 1


def _col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# Per-scenario writers -----------------------------------------------------

def _xl_s1(wb, data, unit):
    name_col = data["params"]["name_key"] or "Object Name"

    ws = wb.create_sheet("Objects")
    _col_widths(ws, [38, 32, 20])
    row = _hdr(ws, 1, ["GUID", name_col, f"Footprint Area ({unit})"])
    for obj in data["objects"]:
        ws.cell(row, 1, obj["guid"])
        ws.cell(row, 2, obj["name"])
        ws.cell(row, 3, obj["area"]).number_format = _AREA_FMT
        row += 1

    ws2 = wb.create_sheet("Summary")
    _col_widths(ws2, [36, 22])
    row = _hdr(ws2, 1, ["S1 — Selected Objects", f"[{unit}]"])
    pairs = [
        ("Name Key used",        data["params"]["name_key"] or "(object name / GUID)"),
        ("Object count",         len(data["objects"])),
        ("Total footprint area", sum(o["area"] for o in data["objects"])),
    ]
    for label, value in pairs:
        lc = ws2.cell(row, 1, label)
        lc.font = Font(bold=True)
        vc = ws2.cell(row, 2, value)
        if isinstance(value, float):
            vc.number_format = _AREA_FMT
        row += 1


def _xl_s2(wb, data, unit):
    obj_col = data["params"]["obj_key"] or "Object Name"
    layer   = data["params"]["layer_name"]
    raw_sum = sum(o["area"] for o in data["objects"])

    ws = wb.create_sheet("Objects")
    _col_widths(ws, [38, 30, 32, 20])
    row = _hdr(ws, 1, ["GUID", "Layer", obj_col, f"Footprint Area ({unit})"])
    for obj in data["objects"]:
        ws.cell(row, 1, obj["guid"])
        ws.cell(row, 2, layer)
        ws.cell(row, 3, obj["name"])
        ws.cell(row, 4, obj["area"]).number_format = _AREA_FMT
        row += 1

    ws2 = wb.create_sheet("Summary")
    _col_widths(ws2, [38, 22])
    row = _hdr(ws2, 1, ["S2 — By Layer", f"[{unit}]"])
    kv_rows = [
        ("Layer",                              layer),
        ("Object Key used",                    data["params"]["obj_key"] or "(object name / GUID)"),
        ("Object count",                       len(data["objects"])),
        (f"Sum of individual areas ({unit})",  raw_sum),
        (f"Combined footprint ({unit})",       data["total"]),
        ("Boolean Union succeeded",            "Yes" if data["union_ok"] else "No — sum shown"),
    ]
    for label, value in kv_rows:
        lc = ws2.cell(row, 1, label)
        lc.font = Font(bold=True)
        vc = ws2.cell(row, 2, value)
        if isinstance(value, float):
            vc.number_format = _AREA_FMT
        row += 1

    overlap = raw_sum - data["total"]
    if overlap > 1e-6:
        row += 1
        row = _sec(ws2, row, "[!] Overlap Warning", n_cols=2)
        row = _warn(ws2, row, "Overlapping area (sum - total)", overlap)
        row = _warn(ws2, row, "Some objects share footprint area. Verify whether double-counting is intentional.")


def _xl_s3(wb, data, unit):
    params  = data["params"]
    obj_col = params["obj_key"] or "Object Name"
    parent  = params["parent"]
    levels  = data["levels"]

    # ── Sheet 1: Objects — flat, pivot-ready ─────────────────────────
    ws = wb.create_sheet("Objects")
    _col_widths(ws, [38, 28, 22, 28, 30, 20])
    row = _hdr(ws, 1, ["GUID", "Parent Layer", "Level", "Category", obj_col,
                       f"Footprint Area ({unit})"])
    for level, lv in levels.items():
        for obj in lv["objects"]:
            ws.cell(row, 1, obj["guid"])
            ws.cell(row, 2, parent)
            ws.cell(row, 3, level)
            ws.cell(row, 4, obj["category"])
            ws.cell(row, 5, obj["name"])
            ws.cell(row, 6, obj["area"]).number_format = _AREA_FMT
            row += 1

    # ── Sheet 2: Summary — Level × Category matrix ───────────────────
    cats = sorted({c for lv in levels.values() for c in lv["category_totals"]})
    n_cols = len(cats) + 2  # Level | categories… | Level Total

    ws2 = wb.create_sheet("Summary")
    _col_widths(ws2, [26] + [20] * (n_cols - 1))
    row = _hdr(ws2, 1, ["S3 — Layer Hierarchy", f"[{unit}]"])

    for label, value in [
        ("Parent Layer",  parent),
        ("Object Key",    params["obj_key"] or "(object name / GUID)"),
        ("Ignore Prefix", params.get("ignore_prefix", "")),
    ]:
        ws2.cell(row, 1, label).font = Font(bold=True)
        ws2.cell(row, 2, value)
        row += 1
    row += 1

    row = _sec(ws2, row, "Combined Footprint — Level × Category", n_cols=n_cols)
    ws2.cell(row, 1, "Level").font = Font(bold=True)
    for ci, cat in enumerate(cats, 2):
        ws2.cell(row, ci, cat).font = Font(bold=True)
    ws2.cell(row, n_cols, "Level Total").font = Font(bold=True)
    row += 1

    for level, lv in levels.items():
        note = "" if lv["union_ok"] else " [union failed]"
        ws2.cell(row, 1, level + note)
        for ci, cat in enumerate(cats, 2):
            if cat in lv["category_totals"]:
                c = ws2.cell(row, ci, lv["category_totals"][cat])
                c.number_format = _AREA_FMT
        tc = ws2.cell(row, n_cols, lv["total"])
        tc.font          = _TOT_FONT
        tc.number_format = _AREA_FMT
        row += 1

    # Bottom row: per-category building totals + grand total
    cat_overall = data.get("category_overall", {})
    lc = ws2.cell(row, 1, "CATEGORY TOTAL (all levels)")
    lc.font = _TOT_FONT
    lc.fill = _TOT_FILL
    for ci, cat in enumerate(cats, 2):
        c = ws2.cell(row, ci, cat_overall.get(cat, 0.0))
        c.font          = _TOT_FONT
        c.fill          = _TOT_FILL
        c.number_format = _AREA_FMT
    gc = ws2.cell(row, n_cols, data["overall_total"])
    gc.font          = _TOT_FONT
    gc.fill          = _TOT_FILL
    gc.number_format = _AREA_FMT
    row += 2

    # Note: Σ(category totals) ≥ level total when categories overlap —
    # the matrix keeps both readings visible.
    warn_rows = []
    for level, lv in levels.items():
        for w_text in lv["warnings"]:
            warn_rows.append(f"{level}: {w_text}")
    warn_rows += list(data.get("warnings", []))

    if warn_rows:
        row = _sec(ws2, row, "[!] Warnings", n_cols=n_cols)
        for w_text in warn_rows:
            row = _warn(ws2, row, w_text, n_cols=n_cols)


def _xl_s4(wb, data, unit):
    """
    S4 Custom Aggregation Excel export.
    Sheet 'Leaf Data': flat table — one row per unique key path (leaf), with
    one column per key level plus an area column. Useful for pivot tables.
    Sheet 'Tree Summary': indented hierarchy showing subtotals at each node.
    """
    params   = data["params"]
    parent   = params["parent"]
    # v2 stores dimension labels; fall back to the v1 key list if present.
    key_seq  = params.get("dim_labels") or params.get("key_sequence") or []
    depth    = len(key_seq)
    key_path = " > ".join(key_seq)

    # ── Sheet 1: Leaf Data ───────────────────────────────────────────
    ws = wb.create_sheet("Leaf Data")
    col_widths = [22] * depth + [20]
    _col_widths(ws, col_widths)
    header_cols = list(key_seq) + [f"Area ({unit})"]
    row = _hdr(ws, 1, header_cols)

    def write_leaves(node, path_so_far):
        nonlocal row
        for val in sorted(node):
            entry = node[val]
            current_path = path_so_far + [val]
            if entry["children"]:
                write_leaves(entry["children"], current_path)
            else:
                # Leaf row: fill each key column
                for ci, pv in enumerate(current_path, 1):
                    ws.cell(row, ci, pv)
                ws.cell(row, depth + 1, entry["area"]).number_format = _AREA_FMT
                row += 1

    write_leaves(data["tree"], [])

    # ── Sheet 2: Tree Summary ────────────────────────────────────────
    ws2 = wb.create_sheet("Tree Summary")
    _col_widths(ws2, [52, 22])
    row2 = _hdr(ws2, 1, ["S4 — Custom Aggregation", f"Area ({unit})"])

    # Header info rows
    for lbl, val in [("Parent Layer", parent), ("Key Hierarchy", key_path)]:
        ws2.cell(row2, 1, lbl).font = Font(bold=True)
        ws2.cell(row2, 2, val)
        row2 += 1
    row2 += 1

    def write_tree_rows(node, level):
        nonlocal row2
        for val in sorted(node):
            entry = node[val]
            indent = "    " * level
            is_leaf = not entry["children"]
            label = f"{indent}{'▸ ' if not is_leaf else '  '}{val}"
            lc = ws2.cell(row2, 1, label)
            vc = ws2.cell(row2, 2, entry["area"])
            vc.number_format = _AREA_FMT
            if not is_leaf:
                lc.font = Font(bold=True)
                lc.fill = _SEC_FILL
                vc.font = Font(bold=True)
                vc.fill = _SEC_FILL
            row2 += 1
            if entry["children"]:
                write_tree_rows(entry["children"], level + 1)

    write_tree_rows(data["tree"], 0)
    row2 += 1

    # Overall total
    lc = ws2.cell(row2, 1, "OVERALL TOTAL")
    lc.font = _TOT_FONT
    lc.fill = _TOT_FILL
    vc = ws2.cell(row2, 2, data["overall_total"])
    vc.font          = _TOT_FONT
    vc.fill          = _TOT_FILL
    vc.number_format = _AREA_FMT


def export_to_excel(data, filepath):
    """Write calculation results to an Excel workbook with Objects + Summary sheets."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    s    = data["scenario"]
    unit = data["unit"]

    if s == 1:
        _xl_s1(wb, data, unit)
    elif s == 2:
        _xl_s2(wb, data, unit)
    elif s == 3:
        _xl_s3(wb, data, unit)
    else:
        _xl_s4(wb, data, unit)

    wb.save(filepath)


# ══════════════════════════════════════════════════════════════════════════════
# Entry point
# ══════════════════════════════════════════════════════════════════════════════

def main():
    form = LinderoForm()
    form.Show()


if __name__ == "__main__":
    main()
