# ui/theme.py
# RhinoGuire — shared design system (neo-brutalist, soft palette)
#
# Usage in any tool:
#   import sys, os
#   _root = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
#   if _root not in sys.path: sys.path.insert(0, _root)
#   from ui import theme

import Eto.Drawing as drawing
import Eto.Forms  as forms

# ── Palette (AAJ Design System) ───────────────────────────────────────────────

# Anchors
BG           = drawing.Color.FromArgb(247, 243, 236)   # Hueso-based page bg (#F7F3EC)
PANEL        = drawing.Color.FromArgb(242, 237, 230)   # Hueso surface (#F2EDE6)
BORDER       = drawing.Color.FromArgb(42,  39,  37)    # Carbón near-black (#2A2725)
TEXT         = drawing.Color.FromArgb(42,  39,  37)    # Carbón primary text (#2A2725)

# Accent colors
HEADER       = drawing.Color.FromArgb(42,  139, 156)   # Mar Caribe teal (#2A8B9C)
TAB_ACTIVE   = drawing.Color.FromArgb(42,  139, 156)   # Mar Caribe teal (#2A8B9C)
TAB_INACTIVE = drawing.Color.FromArgb(230, 225, 216)   # Muted warm grey (#E6E1D8)
ROW_ALT      = drawing.Color.FromArgb(234, 229, 222)   # Subtle alternate row tint (#EAE5DE)
TOTAL_BG     = drawing.Color.FromArgb(213, 234, 223)   # Soft success green (based on #5BA67A)

# Buttons
BTN_CALC     = drawing.Color.FromArgb(42,  139, 156)   # Mar Caribe teal (#2A8B9C)
BTN_CLEAR    = drawing.Color.FromArgb(224, 115, 92)    # Salmon red-orange (#E0735C)
BTN_DEFAULT  = drawing.Color.FromArgb(242, 237, 230)   # Hueso neutral buttons (#F2EDE6)

# Semantic Text
TEXT_MUTED   = drawing.Color.FromArgb(138, 133, 128)   # Tierra warm grey (#8A8580)
TEXT_OK      = drawing.Color.FromArgb(91,  166, 122)   # Success Green (#5BA67A)
TEXT_WARN    = drawing.Color.FromArgb(229, 162, 59)    # Warning Yellow (#E5A23B)
TEXT_ERROR   = drawing.Color.FromArgb(196, 74,  63)    # Error Red (#C44A3F)

# Chart-specific (bullet chart — R1/R2)
CHART_BG     = drawing.Color.FromArgb(216, 211, 202)   # Lighter Tierra (#D8D3CA)
CHART_LOW    = drawing.Color.FromArgb(229, 176, 86)    # Sol de Maíz yellow (#E5B056)
CHART_HIGH   = drawing.Color.FromArgb(224, 115, 92)    # Salmon red-orange (#E0735C)
CHART_BAR    = drawing.Color.FromArgb(42,  139, 156)   # Mar Caribe teal (#2A8B9C)
CHART_GOAL   = drawing.Color.FromArgb(42,  39,  37)    # Carbón near-black (#2A2725)
CHART_TOL    = drawing.Color.FromArgb(138, 133, 128)   # Tierra warm grey (#8A8580)
CHART_LABEL  = drawing.Color.FromArgb(42,  39,  37)    # Carbón primary text (#2A2725)
CHART_DELTA  = drawing.Color.FromArgb(138, 133, 128)   # Tierra warm grey (#8A8580)
CHART_NOTGT  = drawing.Color.FromArgb(196, 74,  63)    # Error Red (#C44A3F)

# ── Spacing (8pt Base) ────────────────────────────────────────────────────────

SPACE_1 = 4
SPACE_2 = 8
SPACE_3 = 12
SPACE_4 = 16
SPACE_6 = 24
SPACE_8 = 32

# ── Fonts ─────────────────────────────────────────────────────────────────────

# Font stacks with dynamic system fallbacks
FONT_DISPLAY = ["Knile", "Fraunces", "Georgia", "Times New Roman"]
FONT_HEADING = ["Geomanist", "Manrope", "Segoe UI Semibold", "Segoe UI", "Arial"]
FONT_BODY    = ["Archia", "Inter", "Segoe UI", "Arial"]
FONT_MONO    = ["Silka Mono", "JetBrains Mono", "Consolas", "Courier New"]

def _get_font(families, size, style=None):
    """Try to construct a font using the first available family name in Eto."""
    try:
        available = {f.Name.lower() for f in drawing.Fonts.AvailableFontFamilies}
    except Exception:
        available = set()

    for family in families:
        if not available or family.lower() in available:
            try:
                if style is not None:
                    return drawing.Font(family, size, style)
                else:
                    return drawing.Font(family, size)
            except Exception:
                pass
    
    # Direct fallback generation if AvailableFontFamilies isn't accessible or is empty
    for family in families:
        try:
            if style is not None:
                return drawing.Font(family, size, style)
            else:
                return drawing.Font(family, size)
        except Exception:
            pass
            
    try:
        if style is not None:
            return drawing.SystemFonts.Default(size, style)
        else:
            return drawing.SystemFonts.Default(size)
    except Exception:
        pass

    try:
        if style == drawing.FontStyle.Bold:
            return drawing.SystemFonts.Bold()
        return drawing.SystemFonts.Default()
    except Exception:
        return None

F_MONO   = _get_font(FONT_MONO, 9.0)
F_MONO_B = _get_font(FONT_MONO, 9.5, drawing.FontStyle.Bold)
F_SANS   = _get_font(FONT_BODY, 9.0)
F_SANS_B = _get_font(FONT_HEADING, 9.0, drawing.FontStyle.Bold)
F_SANS_S = _get_font(FONT_BODY, 8.0)
F_HEAD   = _get_font(FONT_HEADING, 10.0, drawing.FontStyle.Bold)   # section headers (slightly larger)

# ── Builder helpers ───────────────────────────────────────────────────────────

def lbl(text, font=None, color=None, align=None):
    """Create a styled Label."""
    w = forms.Label()
    w.Text = text
    if font:  w.Font          = font  or F_SANS
    if color: w.TextColor     = color
    if align: w.TextAlignment = align
    return w


def btn(text, bg=None):
    """Create a styled Button."""
    w = forms.Button()
    w.Text            = text
    w.Font            = F_SANS_B
    w.BackgroundColor = bg or BTN_DEFAULT
    return w


def pad(content, bg=None, h=8, v=6):
    """Wrap a control in a Panel with background color and padding."""
    p = forms.Panel()
    p.BackgroundColor = bg or BG
    p.Padding         = drawing.Padding(h, v)
    p.Content         = content
    return p


def trow(ctrl, scale=False):
    """Wrap a control in a TableRow with a single TableCell."""
    row = forms.TableRow(forms.TableCell(ctrl))
    row.ScaleHeight = scale
    return row


def section_header(text):
    """Bold label used as a settings/section divider."""
    return lbl(text, F_HEAD, TEXT)


def hint(text):
    """Muted small label used for instructional text."""
    return lbl(text, F_SANS_S, TEXT_MUTED)


def bind_key_search(combo, all_keys):
    """Attach real-time search filtering to a ComboBox.

    As the user types, DataStore is narrowed to keys whose names contain the
    query (case-insensitive substring). Clearing the text or selecting an exact
    match restores the full list so the user can still scroll everything.

    Returns an update(new_keys) callable — call it when the model's available
    key list changes (e.g. after a Refresh Keys action).
    """
    _store = list(all_keys)
    _busy  = [False]

    def _refresh():
        if _busy[0]:
            return
        _busy[0] = True
        try:
            text = (combo.Text or "").lower()
            if not text or any(k.lower() == text for k in _store):
                filtered = list(_store)
            else:
                filtered = [k for k in _store if text in k.lower()]
            current = combo.Text
            combo.DataStore = filtered
            combo.Text = current
        finally:
            _busy[0] = False

    def update(new_keys):
        _store[:] = new_keys
        _refresh()

    def _on_text_changed(s, e):
        _refresh()

    combo.TextChanged += _on_text_changed
    return update


def status_color(state):
    """Return the TextColor for a given status state string.

    States: 'ok', 'warn', 'error', 'info'
    """
    return {
        "ok":    TEXT_OK,
        "warn":  TEXT_WARN,
        "error": TEXT_ERROR,
        "info":  TEXT_MUTED,
    }.get(state, TEXT_MUTED)


_PREFS_PATH = None

def _get_prefs_path():
    global _PREFS_PATH
    if _PREFS_PATH is None:
        import os
        _PREFS_PATH = os.path.normpath(
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "_prefs.json")
        )
    return _PREFS_PATH

def prefs_get(key, fallback=None):
    """Return the last-used folder for *key*, or *fallback* if not set / gone."""
    import os, json
    try:
        path = _get_prefs_path()
        with open(path, 'r') as f:
            folder = json.load(f).get(key)
        if folder and os.path.isdir(folder):
            return folder
    except Exception:
        pass
    return fallback

def prefs_set(key, file_path):
    """Save the directory of *file_path* as the last-used folder for *key*."""
    import os, json
    try:
        path = _get_prefs_path()
        try:
            with open(path, 'r') as f:
                data = json.load(f)
        except Exception:
            data = {}
        data[key] = os.path.dirname(file_path)
        with open(path, 'w') as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass

