#!/usr/bin/env python3
"""Build RhinoGuire.UI's Lucide icon set.

Reads   src/RhinoGuire.UI/Icons/lucide/*.svg   (vendored Lucide icons, ISC, see LICENSE there)
Writes  src/RhinoGuire.UI/Icons/LucideIcons.g.cs

Each SVG (24 x 24 grid, stroked, no fill) becomes one WPF path-markup string: <path>,
<circle>, <rect> (with rx/ry), <line>, <polyline> and <polygon> are converted, and every
path is re-tokenised so WPF's parser accepts it (SVG allows "0 012 2" arc flags and
".5.5" number runs that WPF reads differently). The file name becomes the IconKind
member: "triangle-alert.svg" -> IconKind.TriangleAlert.

To add an icon: copy its SVG from https://lucide.dev (lucide-static package) into the
folder and run this script.

Usage:
    python tools/build_lucide_icons.py           # write LucideIcons.g.cs
    python tools/build_lucide_icons.py --check   # exit 1 if LucideIcons.g.cs is stale

Standard library only.
"""

import argparse
import os
import re
import sys
import xml.etree.ElementTree as ET

ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
ICON_DIR = os.path.join(ROOT, "src", "RhinoGuire.UI", "Icons", "lucide")
OUT_PATH = os.path.join(ROOT, "src", "RhinoGuire.UI", "Icons", "LucideIcons.g.cs")
SVG_NS = "{http://www.w3.org/2000/svg}"

_ARGS = {"M": 2, "L": 2, "H": 1, "V": 1, "C": 6, "S": 4, "Q": 4, "T": 2, "A": 7, "Z": 0}
_NUM = re.compile(r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?")


def num(x):
    x = float(x)
    if x == int(x):
        return str(int(x))
    return ("%.4f" % x).rstrip("0").rstrip(".")


def tokenize_path(d):
    """Yields (command, [arguments], explicit) per argument group; arc flags are read as
    single digits. explicit is False for a group that repeats the previous command."""
    i, n = 0, len(d)
    cmd = None
    while i < n:
        ch = d[i]
        if ch.isspace() or ch == ",":
            i += 1
            continue
        explicit = ch.isalpha()
        if explicit:
            cmd = ch
            i += 1
            if _ARGS[cmd.upper()] == 0:
                yield cmd, [], True
                continue
        elif cmd is None or _ARGS[cmd.upper()] == 0:
            raise ValueError(f"unexpected number in path data {d!r}")
        count = _ARGS[cmd.upper()]
        args = []
        while len(args) < count:
            while i < n and (d[i].isspace() or d[i] == ","):
                i += 1
            if cmd.upper() == "A" and len(args) in (3, 4):
                if i >= n or d[i] not in "01":
                    raise ValueError(f"bad arc flag in {d!r}")
                args.append(d[i])
                i += 1
                continue
            m = _NUM.match(d, i)
            if not m:
                raise ValueError(f"expected a number at {i} in {d!r}")
            args.append(m.group(0))
            i = m.end()
        yield cmd, args, explicit


def normalise_path(d):
    """Re-emits SVG path data in a form WPF parses identically. The leading moveto becomes
    absolute so paths can be concatenated into one geometry, and the implicit lineto after
    a moveto is written out."""
    out = []
    for index, (cmd, args, explicit) in enumerate(tokenize_path(d)):
        if index == 0:
            if cmd not in "Mm":
                raise ValueError(f"path must start with a moveto: {d!r}")
            emit = "M"
        elif not explicit and cmd in "Mm":
            emit = "L" if cmd == "M" else "l"
        else:
            emit = cmd
        is_arc = cmd.upper() == "A"
        text = " ".join(a if is_arc and k in (3, 4) else num(a) for k, a in enumerate(args))
        out.append(emit + text)
    return " ".join(out)


def circle(cx, cy, r):
    cx, cy, r = float(cx), float(cy), float(r)
    return (f"M{num(cx - r)} {num(cy)} A{num(r)} {num(r)} 0 1 0 {num(cx + r)} {num(cy)} "
            f"A{num(r)} {num(r)} 0 1 0 {num(cx - r)} {num(cy)} Z")


def rect(x, y, w, h, rx, ry):
    x, y, w, h = (float(v) for v in (x, y, w, h))
    rx = float(rx) if rx is not None else (float(ry) if ry is not None else 0.0)
    ry = float(ry) if ry is not None else rx
    rx, ry = min(rx, w / 2), min(ry, h / 2)
    if rx <= 0 or ry <= 0:
        return f"M{num(x)} {num(y)} H{num(x + w)} V{num(y + h)} H{num(x)} Z"
    a = f"A{num(rx)} {num(ry)} 0 0 1"
    return (f"M{num(x + rx)} {num(y)} H{num(x + w - rx)} {a} {num(x + w)} {num(y + ry)} "
            f"V{num(y + h - ry)} {a} {num(x + w - rx)} {num(y + h)} H{num(x + rx)} {a} {num(x)} {num(y + h - ry)} "
            f"V{num(y + ry)} {a} {num(x + rx)} {num(y)} Z")


def points(spec, close):
    vals = _NUM.findall(spec)
    pairs = [(vals[i], vals[i + 1]) for i in range(0, len(vals) - 1, 2)]
    s = "M" + " L".join(f"{num(px)} {num(py)}" for px, py in pairs)
    return s + (" Z" if close else "")


def svg_to_path(path):
    tree = ET.parse(path)
    parts = []
    for el in tree.getroot():
        tag = el.tag.replace(SVG_NS, "")
        a = el.attrib
        if tag == "path":
            parts.append(normalise_path(a["d"]))
        elif tag == "circle":
            parts.append(circle(a["cx"], a["cy"], a["r"]))
        elif tag == "ellipse":
            cx, cy, rx, ry = (float(a[k]) for k in ("cx", "cy", "rx", "ry"))
            parts.append(f"M{num(cx - rx)} {num(cy)} A{num(rx)} {num(ry)} 0 1 0 {num(cx + rx)} {num(cy)} "
                         f"A{num(rx)} {num(ry)} 0 1 0 {num(cx - rx)} {num(cy)} Z")
        elif tag == "rect":
            parts.append(rect(a.get("x", 0), a.get("y", 0), a["width"], a["height"], a.get("rx"), a.get("ry")))
        elif tag == "line":
            parts.append(f"M{num(a['x1'])} {num(a['y1'])} L{num(a['x2'])} {num(a['y2'])}")
        elif tag == "polyline":
            parts.append(points(a["points"], close=False))
        elif tag == "polygon":
            parts.append(points(a["points"], close=True))
        else:
            raise ValueError(f"{os.path.basename(path)}: unsupported element <{tag}>")
    return " ".join(parts)


def pascal(name):
    return "".join(p[:1].upper() + p[1:] for p in re.split(r"[-_]", name) if p)


def build():
    files = sorted(f for f in os.listdir(ICON_DIR) if f.endswith(".svg"))
    version = "unknown"
    entries = []
    for f in files:
        full = os.path.join(ICON_DIR, f)
        with open(full, encoding="utf-8") as fh:
            m = re.search(r"lucide-static v([0-9.]+)", fh.read())
            if m:
                version = m.group(1)
        entries.append((pascal(f[:-4]), f[:-4], svg_to_path(full)))

    lines = [
        "// <auto-generated>",
        "// GENERATED FILE, DO NOT EDIT. Built by tools/build_lucide_icons.py from",
        f"// src/RhinoGuire.UI/Icons/lucide/*.svg (Lucide v{version}, ISC licence: Icons/lucide/LICENSE).",
        "// </auto-generated>",
        "",
        "namespace RhinoGuire.UI.Controls;",
        "",
        "/// <summary>The Lucide icons shipped with RhinoGuire.UI. Draw one with <see cref=\"Icon\"/>.</summary>",
        "public enum IconKind",
        "{",
        "    /// <summary>No icon: draws nothing.</summary>",
        "    None = 0,",
    ]
    for pas, kebab, _ in entries:
        lines.append(f"    /// <summary>Lucide <c>{kebab}</c>.</summary>")
        lines.append(f"    {pas},")
    lines += [
        "}",
        "",
        "public static partial class LucideIcons",
        "{",
        f"    /// <summary>Lucide release the SVGs were taken from.</summary>",
        f"    public const string LucideVersion = \"{version}\";",
        "",
        "    /// <summary>Path markup on Lucide's 24 x 24 grid, indexed by <see cref=\"IconKind\"/>.</summary>",
        "    private static readonly string[] PathData =",
        "    {",
        "        \"\", // None",
    ]
    for pas, _, data in entries:
        lines.append(f"        \"{data}\", // {pas}")
    lines += ["    };", "}", ""]
    return "\n".join(lines)


def main(argv=None):
    parser = argparse.ArgumentParser(description="Build Icons/LucideIcons.g.cs from the vendored Lucide SVGs.")
    parser.add_argument("--check", action="store_true", help="fail if the generated file is stale")
    args = parser.parse_args(argv)
    text = build()
    current = open(OUT_PATH, encoding="utf-8").read() if os.path.exists(OUT_PATH) else None
    if args.check:
        if current != text:
            print("Out of date, run `python tools/build_lucide_icons.py`: Icons/LucideIcons.g.cs")
            return 1
        print("Lucide icons are up to date.")
        return 0
    if current != text:
        with open(OUT_PATH, "w", encoding="utf-8", newline="\n") as f:
            f.write(text)
        print("Wrote Icons/LucideIcons.g.cs")
    else:
        print("Icons/LucideIcons.g.cs already up to date")
    return 0


if __name__ == "__main__":
    sys.exit(main())
