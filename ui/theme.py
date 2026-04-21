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

# ── Palette ───────────────────────────────────────────────────────────────────

BG           = drawing.Color.FromArgb(248, 244, 235)   # warm cream — form background
PANEL        = drawing.Color.FromArgb(255, 255, 255)   # white — content panels
HEADER       = drawing.Color.FromArgb(255, 240, 120)   # soft yellow — section headers, active tab
ROW_ALT      = drawing.Color.FromArgb(241, 237, 227)   # subtle stripe — alternate table rows
TOTAL_BG     = drawing.Color.FromArgb(195, 231, 192)   # soft green — totals / success rows
BTN_CALC     = drawing.Color.FromArgb(168, 210, 255)   # soft blue — primary action (Calculate)
BTN_CLEAR    = drawing.Color.FromArgb(255, 188, 188)   # soft red — destructive action (Clear)
BTN_DEFAULT  = drawing.Color.FromArgb(218, 213, 203)   # warm gray — neutral buttons / status bar
TAB_ACTIVE   = drawing.Color.FromArgb(255, 240, 120)   # same as HEADER
TAB_INACTIVE = drawing.Color.FromArgb(232, 228, 219)   # muted warm — inactive tabs
BORDER       = drawing.Color.FromArgb(18,  18,  18)    # near-black — hard borders / gaps

TEXT         = drawing.Color.FromArgb(18,  18,  18)    # near-black — primary text
TEXT_MUTED   = drawing.Color.FromArgb(98,  93,  85)    # warm gray — hints / descriptions
TEXT_ERROR   = drawing.Color.FromArgb(180, 50,  50)    # soft red — errors
TEXT_WARN    = drawing.Color.FromArgb(180, 110, 20)    # amber — warnings
TEXT_OK      = drawing.Color.FromArgb(50,  130, 60)    # soft green — success

# Chart-specific (bullet chart — R1/R2)
CHART_BG     = drawing.Color.FromArgb(218, 213, 203)   # warm gray — bar background
CHART_LOW    = drawing.Color.FromArgb(255, 235, 130)   # soft yellow — below-target zone
CHART_HIGH   = drawing.Color.FromArgb(255, 195, 100)   # soft orange — above-target zone
CHART_BAR    = drawing.Color.FromArgb(100, 140, 190)   # muted blue — measured value bar
CHART_GOAL   = drawing.Color.FromArgb(18,  18,  18)    # near-black — goal line
CHART_TOL    = drawing.Color.FromArgb(140, 135, 128)   # warm gray — tolerance markers
CHART_LABEL  = drawing.Color.FromArgb(18,  18,  18)    # primary text — chart labels
CHART_DELTA  = drawing.Color.FromArgb(98,  93,  85)    # muted — delta % annotation
CHART_NOTGT  = drawing.Color.FromArgb(160, 80,  60)    # terracotta — no-target warning

# ── Fonts ─────────────────────────────────────────────────────────────────────

F_MONO   = drawing.Font("Courier New", 9.0)
F_MONO_B = drawing.Font("Courier New", 9.5, drawing.FontStyle.Bold)
F_SANS   = drawing.Font("Segoe UI",    9.0)
F_SANS_B = drawing.Font("Segoe UI",    9.0, drawing.FontStyle.Bold)
F_SANS_S = drawing.Font("Segoe UI",    8.0)
F_HEAD   = drawing.Font("Segoe UI",    9.0, drawing.FontStyle.Bold)   # section headers

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
