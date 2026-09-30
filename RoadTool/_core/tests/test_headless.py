#! python3
# -*- coding: utf-8 -*-
"""Headless tests for the RhinoCommon-free parts of Trocha (config.py only —
geometry.py/state.py/junctions.py import RhinoCommon at module scope and can
only be exercised inside Rhino; see RoadTool/README.md's smoke-test
checklist for those).

Run with plain CPython (no Rhino needed):

    python RoadTool/_core/tests/test_headless.py
"""

import os
import sys

# Make the repo root importable so `RoadTool._core` resolves.
_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.normpath(os.path.join(_HERE, "..", "..", ".."))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from RoadTool._core import config


_failures = []


def check(name, cond):
    status = "ok  " if cond else "FAIL"
    print("[%s] %s" % (status, name))
    if not cond:
        _failures.append(name)


def approx(a, b, tol=1e-9):
    return abs(a - b) <= tol


# ── config.py ────────────────────────────────────────────────────────────
#
# 2026-07-15 bug: fit_tol/top_rise/margin used to scale off `tolerance`
# (ModelAbsoluteTolerance) directly. On a loose-tolerance file (0.1+, common
# on large site models) that silently produced fit_tol >= 1m and
# top_rise/margin >= 0.5m, inflating a requested 0.1m thickness to 1.316m and
# floating the slab off the terrain. Defaults now derive from `sample_step`
# (a physical, document-scaled knob the user already sets, like width) -
# `tolerance` is only a tiny floor plus its real remaining use: the Brep
# offset/Boolean `tol` argument in geometry.py.

def test_config_defaults_derive_from_sample_step():
    cfg = config.TrochaConfig(tolerance=0.01, sample_step=2.0)
    check("fit_tol defaults to 5% of sample_step", approx(cfg.fit_tol, 0.1))
    check("top_rise defaults to 2% of sample_step", approx(cfg.top_rise, 0.04))
    check("margin defaults to 2% of sample_step", approx(cfg.margin, 0.04))
    check("sample_step defaults to DEFAULT_SAMPLE_STEP when omitted",
          approx(config.TrochaConfig(tolerance=0.01).sample_step, config.DEFAULT_SAMPLE_STEP))
    check("smooth_center defaults on", cfg.smooth_center is True)


def test_config_explicit_overrides_win():
    cfg = config.TrochaConfig(tolerance=0.01, sample_step=1.5, fit_tol=0.02,
                               top_rise=0.03, margin=0.04, smooth_center=False)
    check("sample_step override", approx(cfg.sample_step, 1.5))
    check("fit_tol override", approx(cfg.fit_tol, 0.02))
    check("top_rise override", approx(cfg.top_rise, 0.03))
    check("margin override", approx(cfg.margin, 0.04))
    check("smooth_center override", cfg.smooth_center is False)


def test_loose_tolerance_no_longer_inflates_defaults():
    # Regression test for the 2026-07-15 bug: a loose document tolerance
    # (0.1 - the exact value from the bug report) must not meaningfully
    # inflate fit_tol/top_rise/margin above the sample_step-derived value.
    # (The old tolerance*10 / *5 formula gave fit_tol=1.0, top_rise=0.5 here
    # - a 10x/12x blowup; this asserts it stays within a modest 1.5x.)
    fine = config.TrochaConfig(tolerance=0.001, sample_step=2.0)
    loose = config.TrochaConfig(tolerance=0.1, sample_step=2.0)
    check("loose tolerance keeps fit_tol close to the sample_step-derived value",
          loose.fit_tol <= fine.fit_tol * 1.5)
    check("loose tolerance keeps top_rise close to the sample_step-derived value",
          loose.top_rise <= fine.top_rise * 1.5)
    check("loose tolerance keeps margin close to the sample_step-derived value",
          loose.margin <= fine.margin * 1.5)


def test_tolerance_floor_still_applies():
    # If sample_step is set unreasonably small, the tiny tolerance*0.1 floor
    # keeps the knobs numerically sane instead of collapsing toward zero.
    cfg = config.TrochaConfig(tolerance=0.1, sample_step=0.001)
    check("tolerance floor bounds fit_tol", approx(cfg.fit_tol, 0.01))
    check("tolerance floor bounds top_rise", approx(cfg.top_rise, 0.01))
    check("tolerance floor bounds margin", approx(cfg.margin, 0.01))


def test_tag_keys_are_unique_and_prefixed():
    keys = [config.TAG_ROAD, config.TAG_WIDTH, config.TAG_THICK, config.TAG_TERRAIN,
            config.TAG_CHILD, config.TAG_PARENT, config.TAG_MERGED]
    check("all RG_ROAD tag keys unique", len(keys) == len(set(keys)))
    check("all RG_ROAD tag keys share the RG_ROAD prefix",
          all(k.startswith("RG_ROAD") for k in keys))


# 2026-09-30 bug: an unchecked centerline Rebuild drifted up to 3.65m off the
# drawn line. The drift cap scales with road width unless set explicitly.

def test_max_drift_scales_with_width():
    cfg = config.TrochaConfig(tolerance=0.01)
    check("max_drift defaults to None (resolved per road)", cfg.max_drift is None)
    check("resolved max_drift is 5% of width",
          approx(cfg.resolved_max_drift(6.0), 6.0 * config.MAX_DRIFT_WIDTH_FACTOR))
    check("resolved max_drift never below tolerance",
          approx(config.TrochaConfig(tolerance=0.1).resolved_max_drift(0.5), 0.1))
    check("explicit max_drift wins",
          approx(config.TrochaConfig(tolerance=0.01, max_drift=1.0).resolved_max_drift(6.0), 1.0))


if __name__ == "__main__":
    test_config_defaults_derive_from_sample_step()
    test_config_explicit_overrides_win()
    test_loose_tolerance_no_longer_inflates_defaults()
    test_tolerance_floor_still_applies()
    test_tag_keys_are_unique_and_prefixed()
    test_max_drift_scales_with_width()

    print()
    if _failures:
        print("%d check(s) FAILED: %s" % (len(_failures), ", ".join(_failures)))
        sys.exit(1)
    print("All checks passed.")
