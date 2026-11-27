#!/usr/bin/env python3
"""Build the AAJ skin token dictionaries for the RhinoGuire WPF plugin.

Reads   src/RhinoGuire.UI/Skins/AAJ/tokens.json   (vendored AAJ Design System tokens)
        src/RhinoGuire.UI/Skins/AAJ/skin-map.json (skin contract key -> AAJ token or literal)
Writes  src/RhinoGuire.UI/Skins/AAJ/Tokens.Light.xaml
        src/RhinoGuire.UI/Skins/AAJ/Tokens.Dark.xaml
        src/RhinoGuire.UI/Skins/AAJ/Tokens.Shared.xaml
        src/RhinoGuire.UI/Skins/AAJ/Tokens.<Variant>.xaml  (one per entry in skin-map "variants")

It also computes the WCAG contrast ratio for every pair listed under "contrastWith"
and records it in skin-map.json under "contrast", so the map always shows the
numbers behind each proposed value.

Usage:
    python tools/build_aaj_skin.py            # write the XAML and refresh the contrast numbers
    python tools/build_aaj_skin.py --check    # exit 1 if anything generated is out of date
    python tools/build_aaj_skin.py --report   # print the proposed entries as a Markdown table

Standard library only.
"""

import argparse
import json
import os
import re
import sys

ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
SKIN_DIR = os.path.join(ROOT, "src", "RhinoGuire.UI", "Skins", "AAJ")
MAP_PATH = os.path.join(SKIN_DIR, "skin-map.json")
THEMES = ("light", "dark")

HEADER = (
    "<!-- GENERATED FILE, DO NOT EDIT.\n"
    "     Built by tools/build_aaj_skin.py from Skins/AAJ/tokens.json and Skins/AAJ/skin-map.json.\n"
    "     Change the map (or the vendored tokens) and run the script again.\n"
    "     The keys are specified in docs/port/SKIN_CONTRACT.md. -->\n"
)


# ----------------------------------------------------------------- tokens --

def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


class Tokens:
    """Resolves AAJ token names (with {alias} references) to values."""

    def __init__(self, data):
        self.colors = {t["name"]: t["value"] for t in data["color"]["tokens"]}
        self.dims = {}
        for group in ("spacing", "radius", "borderWidth", "fontSize", "lineHeight", "letterSpacing"):
            for t in data.get(group, {}).get("tokens", []):
                self.dims[t["name"]] = t["value"]

    def color(self, name, seen=()):
        if name not in self.colors:
            raise KeyError(f"unknown AAJ colour token '{name}'")
        value = self.colors[name]
        m = re.fullmatch(r"\{([^}]+)\}", value.strip())
        if m:
            if m.group(1) in seen:
                raise ValueError(f"alias loop at '{name}'")
            return self.color(m.group(1), seen + (name,))
        return normalise_hex(value)

    def dim(self, name):
        """A dimension token in px (rem converted at 16px)."""
        if name not in self.dims:
            raise KeyError(f"unknown AAJ dimension token '{name}'")
        value = str(self.dims[name]).strip()
        m = re.fullmatch(r"(-?[0-9.]+)(px|rem)?", value)
        if not m:
            raise ValueError(f"token '{name}' = '{value}' is not a plain px/rem value")
        number = float(m.group(1))
        return number * 16.0 if m.group(2) == "rem" else number


def normalise_hex(value):
    v = value.strip().lstrip("#")
    if len(v) == 3:
        v = "".join(c * 2 for c in v)
    if not re.fullmatch(r"[0-9a-fA-F]{6}", v):
        raise ValueError(f"not an opaque #RRGGBB colour: '{value}'")
    return "#" + v.upper()


def resolve_color(tokens, spec):
    return normalise_hex(spec) if spec.startswith("#") else tokens.color(spec)


# --------------------------------------------------------------- contrast --

def _linear(c):
    c /= 255.0
    return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4


def luminance(hex_color):
    h = hex_color.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * _linear(r) + 0.7152 * _linear(g) + 0.0722 * _linear(b)


def contrast(a, b):
    la, lb = sorted((luminance(a), luminance(b)), reverse=True)
    return round((la + 0.05) / (lb + 0.05), 2)


# ------------------------------------------------------------------ build --

def xml_comment(text):
    """XML comments may not contain '--'."""
    while "--" in text:
        text = text.replace("--", "- -")
    return text


def fmt_number(x):
    x = float(x)
    return str(int(x)) if x == int(x) else repr(round(x, 4))


def fmt_thickness(value):
    parts = [p.strip() for p in str(value).split(",")]
    return ",".join(fmt_number(p) for p in parts)


def shared_value(tokens, key, entry, override=None):
    """Returns the XAML text for a shared metric."""
    kind = entry["type"]
    if override is not None:
        raw = override
    elif "value" in entry:
        raw = entry["value"]
    else:
        raw = tokens.dim(entry["token"])
    if kind == "Double":
        return fmt_number(raw)
    if kind == "Thickness":
        return fmt_thickness(raw)
    if kind == "CornerRadius":
        return fmt_number(raw)
    if kind in ("FontFamily", "FontWeight", "Visibility"):
        return str(raw)
    raise ValueError(f"{key}: unsupported type '{kind}'")


XAML_TAG = {
    "Double": "sys:Double",
    "Thickness": "Thickness",
    "CornerRadius": "CornerRadius",
    "FontFamily": "FontFamily",
    "FontWeight": "FontWeight",
    "Visibility": "Visibility",
}


def dictionary(body_lines, needs_sys):
    ns = [
        '<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"',
        '                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"',
    ]
    if needs_sys:
        ns.append('                    xmlns:sys="clr-namespace:System;assembly=mscorlib"')
    ns[-1] += ">"
    return HEADER + "\n".join(ns) + "\n" + "\n".join(body_lines) + "\n</ResourceDictionary>\n"


def color_lines(tokens, section, theme, title):
    lines = [f"\n    <!-- {title} -->"]
    for key, entry in section.items():
        spec = entry[theme]
        value = resolve_color(tokens, spec)
        source = "literal" if spec.startswith("#") else spec
        flag = ", RhinoGuire value (not in AAJ), approved at Gate 1A" if entry.get("proposed") else ""
        lines.append(f"    <!-- {key}: {source}{flag} -->")
        lines.append(f'    <Color x:Key="{key}Color">{value}</Color>')
        lines.append(f'    <SolidColorBrush x:Key="{key}Brush" Color="{value}" />')
    return lines


def build_theme(tokens, skin_map, theme):
    lines = color_lines(tokens, skin_map["colors"], theme, "Contract colours (docs/port/SKIN_CONTRACT.md)")
    lines += color_lines(tokens, skin_map.get("skinColors", {}), theme, "AAJ-only colours, used by Skins/AAJ/Controls.xaml only")
    return dictionary(lines, needs_sys=False)


def shared_lines(tokens, section, title, overrides=None):
    lines = [f"\n    <!-- {title} -->"]
    for key, entry in section.items():
        if overrides is not None and key not in overrides:
            continue
        value = shared_value(tokens, key, entry, None if overrides is None else overrides[key])
        tag = XAML_TAG[entry["type"]]
        source = entry.get("token", "literal")
        notes = [source] + ([entry["note"]] if entry.get("note") else [])
        if entry.get("proposed"):
            notes.append("RhinoGuire value (not in AAJ), approved at Gate 1A")
        lines.append(f"    <!-- {key}: {xml_comment('; '.join(notes))} -->")
        lines.append(f'    <{tag} x:Key="{key}">{value}</{tag}>')
    return lines


def build_shared(tokens, skin_map):
    lines = shared_lines(tokens, skin_map["shared"], "Contract metrics (docs/port/SKIN_CONTRACT.md)")
    lines += shared_lines(tokens, skin_map.get("skinShared", {}), "AAJ-only metrics, used by Skins/AAJ/Controls.xaml only")
    return dictionary(lines, needs_sys=True)


def build_variant(tokens, skin_map, name, overrides):
    overrides = {k: v for k, v in overrides.items() if not k.startswith("$")}
    all_shared = dict(skin_map["shared"])
    all_shared.update(skin_map.get("skinShared", {}))
    unknown = [k for k in overrides if k not in all_shared]
    if unknown:
        raise KeyError(f"variant {name}: unknown keys {unknown}")
    section = {k: all_shared[k] for k in overrides}
    comment = skin_map["variants"][name].get("$comment", "")
    lines = [f"\n    <!-- Variant '{name}': merged after Tokens.Shared.xaml. {comment} -->"]
    lines += shared_lines(tokens, section, f"Overrides for {name}", overrides)
    return dictionary(lines, needs_sys=True)


def compute_contrast(tokens, skin_map):
    """Fills entry['contrast'] = {theme: {other: ratio}} for every contrastWith list."""
    palette = {}
    for section in ("colors", "skinColors"):
        for key, entry in skin_map.get(section, {}).items():
            palette[key] = {t: resolve_color(tokens, entry[t]) for t in THEMES}
    for section in ("colors", "skinColors"):
        for key, entry in skin_map.get(section, {}).items():
            others = entry.get("contrastWith")
            if not others:
                entry.pop("contrast", None)
                continue
            result = {}
            for theme in THEMES:
                result[theme] = {o: contrast(palette[key][theme], palette[o][theme]) for o in others}
            entry["contrast"] = result
    return palette


def dump_map(skin_map):
    """JSON with one line per entry, so the map stays readable and diffs stay small."""
    out = ["{"]
    top = list(skin_map.items())
    for i, (k, v) in enumerate(top):
        comma = "," if i < len(top) - 1 else ""
        if isinstance(v, dict) and v and all(isinstance(x, dict) for x in v.values()):
            out.append(f"  {json.dumps(k)}: {{")
            items = list(v.items())
            for j, (ek, ev) in enumerate(items):
                c2 = "," if j < len(items) - 1 else ""
                out.append(f"    {json.dumps(ek)}: {json.dumps(ev, ensure_ascii=False)}{c2}")
            out.append(f"  }}{comma}")
        else:
            out.append(f"  {json.dumps(k)}: {json.dumps(v, ensure_ascii=False)}{comma}")
    out.append("}")
    return "\n".join(out) + "\n"


def outputs(tokens, skin_map):
    files = {
        "Tokens.Light.xaml": build_theme(tokens, skin_map, "light"),
        "Tokens.Dark.xaml": build_theme(tokens, skin_map, "dark"),
        "Tokens.Shared.xaml": build_shared(tokens, skin_map),
    }
    for name, overrides in skin_map.get("variants", {}).items():
        files[f"Tokens.{name}.xaml"] = build_variant(tokens, skin_map, name, overrides)
    return files


def report(tokens, skin_map):
    print("| Key | Light | Dark | Contrast (light / dark) | Note |")
    print("| --- | --- | --- | --- | --- |")
    for section in ("colors", "skinColors"):
        for key, entry in skin_map.get(section, {}).items():
            if not entry.get("proposed"):
                continue
            vals = []
            for theme in THEMES:
                spec = entry[theme]
                hexv = resolve_color(tokens, spec)
                vals.append(f"`{hexv}`" + ("" if spec.startswith("#") else f" ({spec})"))
            c = entry.get("contrast", {})
            pairs = []
            for other in entry.get("contrastWith", []):
                pairs.append(f"{other} {c['light'][other]:.2f} / {c['dark'][other]:.2f}")
            print(f"| `{key}` | {vals[0]} | {vals[1]} | {'; '.join(pairs) or '-'} | {entry.get('note', '')} |")
    for section in ("shared", "skinShared"):
        for key, entry in skin_map.get(section, {}).items():
            if entry.get("proposed"):
                print(f"| `{key}` | {shared_value(tokens, key, entry)} | same | - | {entry.get('note', '')} |")


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--check", action="store_true", help="fail if generated files or contrast numbers are stale")
    parser.add_argument("--report", action="store_true", help="print the proposed entries as Markdown")
    args = parser.parse_args(argv)

    skin_map = load_json(MAP_PATH)
    tokens = Tokens(load_json(os.path.join(SKIN_DIR, skin_map.get("source", "tokens.json"))))
    compute_contrast(tokens, skin_map)
    files = outputs(tokens, skin_map)
    map_text = dump_map(skin_map)

    if args.report:
        report(tokens, skin_map)
        return 0

    stale = []
    for name, text in files.items():
        path = os.path.join(SKIN_DIR, name)
        current = open(path, encoding="utf-8").read() if os.path.exists(path) else None
        if current != text:
            stale.append(name)
            if not args.check:
                with open(path, "w", encoding="utf-8", newline="\n") as f:
                    f.write(text)
    current_map = open(MAP_PATH, encoding="utf-8").read()
    if current_map != map_text:
        stale.append("skin-map.json (contrast numbers / formatting)")
        if not args.check:
            with open(MAP_PATH, "w", encoding="utf-8", newline="\n") as f:
                f.write(map_text)

    if args.check:
        if stale:
            print("Out of date, run `python tools/build_aaj_skin.py`: " + ", ".join(stale))
            return 1
        print("AAJ skin tokens are up to date.")
        return 0
    print("Wrote: " + (", ".join(stale) if stale else "nothing (already up to date)"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
