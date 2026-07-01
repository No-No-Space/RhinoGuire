#! python3
# ─────────────────────────────────────────────────────────────────────────────
# debug_merge.py — diagnostic for Lindero's combined_area()
#
# Usage: select the objects of ONE layer (e.g. the four Diagnostik objects),
# then RunPythonScript → this file. Output goes to the Rhino command history
# (press F2 to see it all) and to debug_merge_report.txt next to this script.
#
# It runs the exact same pipeline as Lindero v0.6 and reports every
# intermediate value, so we can see where the merged total diverges.
# ─────────────────────────────────────────────────────────────────────────────

import os
import rhinoscriptsyntax as rs
import scriptcontext as sc
import Rhino.Geometry as rg

import sys as _sys, os as _os
_rg_root = _os.path.normpath(_os.path.join(_os.path.dirname(_os.path.abspath(__file__)), ".."))
if _rg_root not in _sys.path:
    _sys.path.insert(0, _rg_root)

# Import the live Lindero module functions (reload to pick up edits)
import importlib
import AreaMeasurer.Lindero as L
importlib.reload(L)

OUT = []
def say(msg=""):
    OUT.append(str(msg))
    print(msg)

def main():
    doc_tol = sc.doc.ModelAbsoluteTolerance
    tol = L._bool_tol()  # clamped tolerance actually used by Lindero
    say("═" * 70)
    say("Lindero merge diagnostic")
    say("ModelAbsoluteTolerance: %r   effective 2D tolerance: %r   Unit system: %s"
        % (doc_tol, tol, sc.doc.ModelUnitSystem))
    guids = rs.SelectedObjects() or []
    say("Selected objects: %d" % len(guids))
    if not guids:
        say("Nothing selected — select the objects of one layer and rerun.")
        return

    all_curves = []
    per_object = []  # (label, breps, loops)
    say("\n── Per object ────────────────────────────────────────────────")
    for g in guids:
        loops = L.get_footprint_curves(g)
        layer = rs.ObjectLayer(g) or "?"
        label = str(g)[:8]
        say("Object %s  layer=%s  loops=%d" % (label, layer, len(loops)))
        for i, c in enumerate(loops):
            say("   loop %d: type=%s closed=%s area=%.4f"
                % (i, type(c).__name__, c.IsClosed, L.curve_area(c)))
        breps = None
        try:
            breps = rg.Brep.CreatePlanarBreps(loops, tol)
        except Exception as ex:
            say("   CreatePlanarBreps EXCEPTION: %s" % ex)
        if breps:
            say("   planar breps: %d, total area %.4f (holes resolved)"
                % (len(breps), sum(b.GetArea() for b in breps)))
        else:
            say("   planar breps: NONE — object falls back to loop parity")
        say("   _region_area (shown in S1 list): %.4f" % L._region_area(loops))
        all_curves.extend(loops)
        per_object.append((label, list(breps) if breps else None, loops))

    say("\n── Arrangement (Curve.CreateBooleanRegions) ─────────────────")
    boolean_regions = getattr(rg.Curve, "CreateBooleanRegions", None)
    if boolean_regions is None:
        say("Curve.CreateBooleanRegions NOT AVAILABLE → Lindero uses the")
        say("hole-blind Curve.CreateBooleanUnion fallback. That is the bug.")
        return
    try:
        arr = boolean_regions(all_curves, rg.Plane.WorldXY, False, tol)
    except Exception as ex:
        say("CreateBooleanRegions EXCEPTION: %s → fallback path runs." % ex)
        return
    if not arr or arr.RegionCount == 0:
        say("Arrangement empty (RegionCount=0) → fallback path runs.")
        return
    say("RegionCount: %d" % arr.RegionCount)

    total = 0.0
    covered_area = 0.0
    cov_per_obj = {label: 0.0 for label, _, _ in per_object}
    say("\n idx |     area | int.pt found | covered by            | counted")
    say(" " + "-" * 66)
    for i in range(arr.RegionCount):
        loops = list(arr.RegionCurves(i) or [])
        if not loops:
            say(" %3d | (no curves)" % i)
            continue
        try:
            breps = rg.Brep.CreatePlanarBreps(loops, tol)
        except Exception as ex:
            say(" %3d | CreatePlanarBreps EXCEPTION: %s" % (i, ex))
            continue
        if not breps:
            say(" %3d | region breps FAILED, raw loop area %.4f"
                % (i, L._region_area(loops)))
            continue
        for b in breps:
            a = b.GetArea()
            pt = L._interior_point(b)
            if pt is None:
                say(" %3d | %8.2f | NO INTERIOR POINT — region skipped as unresolved!" % (i, a))
                continue
            covers = []
            for label, obreps, oloops in per_object:
                if L._covered_by_object(pt, obreps, oloops, tol):
                    covers.append(label)
                    cov_per_obj[label] += a
            counted = bool(covers)
            if counted:
                covered_area += a
            say(" %3d | %8.2f | yes          | %-21s | %s"
                % (i, a, ",".join(covers) if covers else "—", "YES" if counted else "no"))
            total += a

    say("\n── Conservation check (arrangement validity) ────────────────")
    say("Regions covered by each object must sum back to its own area:")
    conserved = True
    for label, obreps, oloops in per_object:
        own = (sum(b.GetArea() for b in obreps) if obreps
               else L._region_area(oloops))
        cov = cov_per_obj[label]
        ok = abs(cov - own) <= max(0.001 * own, 1.0)
        if not ok:
            conserved = False
        say("  %s  own=%10.2f   covered=%10.2f   %s"
            % (label, own, cov, "ok" if ok else "VIOLATED"))
    say("Arrangement %s" % ("passes — trusted."
                            if conserved else "REJECTED — corrupt."))

    say("\n── Pairwise inclusion–exclusion (path 2) ────────────────────")
    pw_regions = [(obreps, oloops) for _, obreps, oloops in per_object]
    pw, pw_ok = L._pairwise_union_area(pw_regions, tol, (tol * 10.0) ** 2)
    say("pairwise union area: %.4f   (ok=%s)" % (pw, pw_ok))

    say("\n── Summary ──────────────────────────────────────────────────")
    say("Sum of all arrangement regions (gross plane coverage): %.4f" % total)
    say("Arrangement covered total:                             %.4f" % covered_area)
    say("Pairwise inclusion–exclusion total:                    %.4f" % pw)
    say("Lindero combined_area() (what the UI uses):            %.4f"
        % L.combined_area(guids)[0])

    path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "debug_merge_report.txt")
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(OUT))
        say("\nReport written to: %s" % path)
    except Exception as ex:
        say("Could not write report file: %s" % ex)

main()
