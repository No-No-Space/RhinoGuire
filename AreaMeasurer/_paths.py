#! python3
# -*- coding: utf-8 -*-
"""
Layer-path logic for Lindero — pure Python, NO Rhino imports.

Encodes the v2 layer-structure contract (see AreaMeasurer/PLAN.md §2):

    parent (proposal layer, chosen in the UI)
    └── level            = direct child of the parent      e.g. Ebene_06
        └── ... any depth ...
            └── category = the object's own (lowest) layer e.g. Pflege

Rules:
- A layer whose short name starts with the ignore prefix (default "_") is
  excluded together with its ENTIRE subtree (cascade). An empty prefix
  disables ignoring.
- The prefix is only tested on segments BELOW the parent — the parent
  itself (and its ancestors) are the user's explicit choice.
- Objects directly on a level layer are measured with category PLACEHOLDER.
- Objects directly on the parent layer are not measured (the preview
  reports them).

Layer paths use Rhino's "::" separator. Everything here is pure string
logic, so this module is headless-testable:

    python AreaMeasurer/tests/test_paths.py
"""

import re

SEP = "::"
PLACEHOLDER = "—"
DEFAULT_IGNORE_PREFIX = "_"

# classify() kinds
OUTSIDE = "outside"   # not under the parent (or the parent itself)
PARENT  = "parent"    # the parent layer itself
IGNORED = "ignored"   # excluded by the prefix cascade
LEVEL   = "level"     # a level layer (direct child of the parent)
MEMBER  = "member"    # a measurable layer inside a level


def short_name(full_path):
    """Last segment of a layer full path."""
    return full_path.split(SEP)[-1]


def is_ignored_name(name, prefix):
    """True if a single segment name is excluded by the prefix. Empty prefix
    disables ignoring."""
    return bool(prefix) and name.startswith(prefix)


def rel_segments(parent, full_path):
    """Segments of full_path below parent.

    Returns [] when full_path IS the parent, None when full_path is not
    under the parent at all.
    """
    if full_path == parent:
        return []
    pre = parent + SEP
    if not full_path.startswith(pre):
        return None
    return full_path[len(pre):].split(SEP)


def classify(parent, full_path, prefix=DEFAULT_IGNORE_PREFIX):
    """Classify one layer path against the v2 contract.

    Returns (kind, level, category):
      (OUTSIDE, None,  None)        — not under the parent
      (PARENT,  None,  None)        — the parent itself
      (IGNORED, level_or_None, None)— prefix cascade (level None when the
                                       level segment itself is prefixed)
      (LEVEL,   level, PLACEHOLDER) — the level layer itself
      (MEMBER,  level, category)    — measurable layer inside a level;
                                       category = its own short name
    """
    rel = rel_segments(parent, full_path)
    if rel is None:
        return (OUTSIDE, None, None)
    if not rel:
        return (PARENT, None, None)
    if is_ignored_name(rel[0], prefix):
        return (IGNORED, None, None)
    if any(is_ignored_name(s, prefix) for s in rel[1:]):
        return (IGNORED, rel[0], None)
    level = rel[0]
    if len(rel) == 1:
        return (LEVEL, level, PLACEHOLDER)
    return (MEMBER, level, rel[-1])


def partition(parent, layer_paths, prefix=DEFAULT_IGNORE_PREFIX):
    """Split all layers under a parent into measure layers and ignored ones.

    Returns (included, ignored):
      included — [(full_path, level, category)] for LEVEL and MEMBER layers
                 (LEVEL rows carry category PLACEHOLDER: objects sitting
                 directly on a level layer are still measured)
      ignored  — sorted [full_path] excluded by the prefix cascade

    Layers outside the parent subtree, and the parent itself, appear in
    neither list.
    """
    included, ignored = [], []
    for lp in layer_paths:
        kind, level, category = classify(parent, lp, prefix)
        if kind in (LEVEL, MEMBER):
            included.append((lp, level, category))
        elif kind == IGNORED:
            ignored.append(lp)
    included.sort(key=lambda t: (level_sort_key(t[1]), t[2]))
    return included, sorted(ignored)


def build_tree(included):
    """{level: {category: [layer_full_paths]}} from partition() output."""
    tree = {}
    for lp, level, category in included:
        tree.setdefault(level, {}).setdefault(category, []).append(lp)
    return tree


def segment_at_depth(parent, full_path, depth):
    """Segment at a 1-based depth below the parent (1 = level).

    Returns PLACEHOLDER when the layer is not under the parent or not that
    deep. Used by the S4 'Layer @ depth' dimension.
    """
    rel = rel_segments(parent, full_path)
    if not rel or depth < 1 or depth > len(rel):
        return PLACEHOLDER
    return rel[depth - 1]


_NUM_RE = re.compile(r"-?\d+")


def level_sort_key(name):
    """Display ordering for level names — top floor first.

    Names containing a signed integer (Ebene_06, Ebene_-01, Ebene_00-Eingang)
    sort before purely textual ones (Ebene_KMA), descending by the LAST
    number found. Textual names follow, case-insensitive alphabetical.
    Affects display order only — never totals.
    """
    if name is None:
        return (2, 0, "")
    matches = _NUM_RE.findall(name)
    if matches:
        return (0, -int(matches[-1]), name.lower())
    return (1, 0, name.lower())
