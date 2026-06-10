#! python3
# -*- coding: utf-8 -*-
"""Slope unit conversions for TerrainTools (see DECISIONS D2).

Everything is stored internally as a single canonical gradient::

    m = rise / run = V / H        (vertical over horizontal, always > 0)

The civil ratio convention used here is **H:V** — e.g. ``2:1`` means
2 horizontal to 1 vertical = 50% = 26.57 degrees.

Pure stdlib (math only) so it can be unit-tested headless.
"""

import math

UNITS = ("ratio_hv", "percent", "degrees")

# Human labels for each unit (for dropdowns / status text).
UNIT_LABELS = {
    "ratio_hv": "H:V",
    "percent":  "%",
    "degrees":  "°",
}


def _parse_ratio(value):
    """Return (H, V) from a ratio input.

    Accepts a float/int ``H`` (interpreted as ``H:1``) or a string ``"H:V"``
    (or ``"H"``). Raises ValueError on non-positive components.
    """
    if isinstance(value, str):
        s = value.strip()
        if ":" in s:
            h_str, v_str = s.split(":", 1)
            h, v = float(h_str), float(v_str)
        else:
            h, v = float(s), 1.0
    else:
        h, v = float(value), 1.0
    if h <= 0.0 or v <= 0.0:
        raise ValueError("Ratio H:V components must both be > 0")
    return h, v


def to_gradient(value, unit):
    """Convert a slope expressed in *unit* to canonical gradient ``m = V/H`` (> 0).

    unit in {'ratio_hv', 'percent', 'degrees'}.
      ratio_hv : float H (=> H:1) or string "H:V"   -> m = V / H
      percent  : p                                  -> m = p / 100
      degrees  : theta in (0, 90)                   -> m = tan(theta)

    Raises ValueError for non-positive / out-of-range input or unknown unit.
    """
    if unit == "ratio_hv":
        h, v = _parse_ratio(value)
        return v / h
    if unit == "percent":
        p = float(value)
        if p <= 0.0:
            raise ValueError("Percent slope must be > 0")
        return p / 100.0
    if unit == "degrees":
        d = float(value)
        if d <= 0.0 or d >= 90.0:
            raise ValueError("Degrees must be in the open range (0, 90)")
        return math.tan(math.radians(d))
    raise ValueError("Unknown slope unit: %r" % (unit,))


def from_gradient(m, unit):
    """Convert a canonical gradient *m* back to a display value in *unit*.

    ratio_hv -> H (the run per 1 unit of rise, i.e. the ``H`` in ``H:1``)
    percent  -> 100 * m
    degrees  -> atan(m) in degrees
    """
    if m <= 0.0:
        raise ValueError("Gradient m must be > 0")
    if unit == "ratio_hv":
        return 1.0 / m
    if unit == "percent":
        return 100.0 * m
    if unit == "degrees":
        return math.degrees(math.atan(m))
    raise ValueError("Unknown slope unit: %r" % (unit,))


def _fmt(x):
    """Trim trailing zeros from a float for compact labels."""
    return ("%.4f" % x).rstrip("0").rstrip(".")


def format_slope(m, unit):
    """Human label for gradient *m* in *unit*, e.g. '2:1', '50%', '26.57°'."""
    if m is None or m <= 0.0:
        return "flat"
    if unit == "ratio_hv":
        return "%s:1" % _fmt(1.0 / m)
    if unit == "percent":
        return "%s%%" % _fmt(100.0 * m)
    if unit == "degrees":
        return "%s°" % _fmt(math.degrees(math.atan(m)))
    return _fmt(m)


def describe(m):
    """Return all three unit equivalents, e.g. '2:1 · 50% · 26.57°'."""
    if m is None or m <= 0.0:
        return "flat"
    return " · ".join((
        format_slope(m, "ratio_hv"),
        format_slope(m, "percent"),
        format_slope(m, "degrees"),
    ))
