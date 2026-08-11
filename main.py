r"""
PSD Tools: Batch Text Placer + Website Checker (GUI, Windows)
==============================================================

Two tools, one window, one .exe -- switch between them with tabs at the top:

  TAB 1: Batch Text Placer
  ------------------------
  [ Choose Folder... ]  <folder path label>

  Extraction source:
    Saved HTML/MHTML file:  <path label>  [ Browse... ]
    (leave unset to reuse an existing textboxes.json in the folder)

  [ ] Resize before placing text     Width (px): [____]   DPI: [____]

  [ Run ]  [ Show details ]

  TAB 2: PSD <-> Website Checker
  ------------------------------
  [ Choose Folder... ]  <folder path label>
  [ ] Ignore case  [ ] Ignore whitespace  [ ] Ignore line breaks  [ ] Check formatting

  [ Run Check ]  [ Save CSV Report... ]

  <color-coded results table>

WORKFLOW
--------
The folder you choose only needs to contain PSD/PSB/TIFF files -- no HTML
or JSON file has to be saved there ahead of time.

To get textboxes.json:
  SAVED FILE (recommended for distributing to teammates): open the
     chapter in your own browser while logged in, save it as a single
     MHTML file after it's fully rendered, then Browse to that file here.
     No login, no network access, and no shared credentials are needed at
     extraction time -- the file is opened straight off disk. A real
     browser is still used to run the extraction script (it reads live
     layout via getBoundingClientRect()/getComputedStyle(), which only a
     rendering engine can compute), but it's your own already-installed
     Chrome/Edge, driven locally.

If unset and a textboxes.json already sits in the folder, that
existing JSON is used instead and nothing is opened at all.

NOTES ON THE SAVED-FILE OPTION
-------------------------------
Either "Webpage, Complete" (.html) or "Webpage, Single File" (.mhtml) works.
Chrome normally opens .mhtml through a sandboxed viewer that blocks all
script execution, which would break extraction -- so when a .mhtml/.mht file
is selected, it is first automatically unpacked (the MIME multipart message
is parsed and re-saved as a plain, self-contained .html file plus its
resource files, in a temp folder) before it's opened for extraction. That
unpacked copy runs with scripts fully enabled, just like a real .html save.
.html ("Webpage, Complete") still needs no unpacking step, so it's slightly
faster and is picked automatically over a .mhtml/.mht file when both are
present in the same folder.

REQUIREMENTS (only needed to run/build from source -- NOT needed by teammates
using the built .exe, see below)
------------------------------------------------------------------------------
    pip install pywin32          (Tab 1: talks to Photoshop)
    pip install playwright       (Tab 1: drives a dedicated Chrome/Edge profile
                                   to open the saved file and run the
                                   extraction script; no extra browser
                                   download required)
    pip install beautifulsoup4   (Tab 2: parses the saved website export)
    pip install psd-tools        (Tab 2: reads PSD text layers/fonts directly,
                                   no Photoshop needed for this tab)

BUILDING A STANDALONE .EXE
---------------------------
Easiest: on a Windows machine with Python installed, just double-click
build.bat (included alongside this script). It installs all of the above
build dependencies and runs PyInstaller for you, then drops the finished
PSD-Batch-Placer.exe in the "dist" folder -- ONE file, with both tabs
built in, ready to send to a teammate.

Manual equivalent:
    pip install pywin32 playwright beautifulsoup4 psd-tools pyinstaller
    pyinstaller --onefile --windowed --name "PSD-Batch-Placer" ^
        --icon icon.ico --add-data "icon.ico;." ^
        gui_batch_photoshop_automation.py

    (icon.ico should sit alongside this script. --icon sets the built
    .exe's file thumbnail; --add-data bundles the same file inside the
    .exe so the running app's window can also use it -- see
    resource_path()/iconbitmap() near the top of MainApp below.)

    Teammates running the built .exe/.app need NOTHING installed except:
      - Tab 1 (Batch Text Placer): Google Chrome or Microsoft Edge (only if
        using the "Saved HTML/MHTML file" extraction source -- this drives
        their existing browser via its devtools protocol, it does not
        bundle a browser), and Photoshop already open before clicking Run.
        Works on both Windows (via pywin32/COM) and macOS (via the
        built-in `osascript`/JXA bridge -- no extra install needed on Mac
        either, since osascript ships with macOS).
      - Tab 2 (PSD <-> Website Checker): nothing extra at all -- it reads
        the PSD file's text layers directly, no Photoshop needed for this tab.
    They do NOT need Python, pip, or any of the packages above installed
    themselves -- PyInstaller bundles all of that into the single .exe/.app.
"""

import csv
import difflib
import email
import glob
import io
import json
import math
import mimetypes
import os
import re
import struct
import subprocess
import sys
import tempfile
import threading
import traceback
from email import policy as email_policy
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext


def resource_path(filename):
    """
    Finds a bundled resource (e.g. icon.ico) both when running this
    script directly from source and when running as a PyInstaller
    --onefile .exe (where bundled files are unpacked to a temp folder
    referenced by sys._MEIPASS at runtime).
    """
    base_dir = getattr(sys, '_MEIPASS', None) or os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, filename)

try:
    import win32com.client
    import pythoncom
    import win32gui
    import win32con
    import win32api
except ImportError:
    win32com = None
    pythoncom = None
    win32gui = None
    win32con = None
    win32api = None

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
except ImportError:
    sync_playwright = None
    PlaywrightTimeoutError = Exception

try:
    from bs4 import BeautifulSoup, NavigableString, Tag
except ImportError:
    BeautifulSoup = None
    NavigableString = None
    Tag = None

try:
    from psd_tools import PSDImage
except ImportError:
    PSDImage = None

try:
    from updater import check_for_update, download_update, apply_update
except ImportError:
    check_for_update = None
    download_update = None
    apply_update = None

# Nho tang so nay moi lan ban tag + push ban moi (vd tag v1.0.1 -> "1.0.1")
APP_VERSION = "1.0.4"


IMAGE_EXTENSIONS = (".psd", ".psb", ".tif", ".tiff")
WEB_EXTENSIONS = (".html", ".htm", ".mhtml", ".mht")


# ══════════════════════════════════════════════════════════════════
#  THEME  --  white + blue + light gray, applied via ttk.Style
# ══════════════════════════════════════════════════════════════════

COLORS = {
    "bg":          "#F1F4F9",   # app / window background (light gray)
    "surface":     "#FFFFFF",   # card / panel background
    "border":      "#E2E8F2",   # hairline borders on cards & inputs
    "primary":     "#2F6FED",   # main blue (buttons, active states)
    "primary_dk":  "#255BC7",   # pressed/hover blue
    "primary_lt":  "#EAF1FE",   # pale blue (selected tab / highlight fill)
    "text":        "#1E2530",   # primary text
    "text_muted":  "#6B7686",   # secondary/help text
    "text_on_blue":"#FFFFFF",
    "success":     "#1FA971",
}


def setup_style(root):
    """Configures a light, white/blue/gray ttk theme for the whole app."""
    root.configure(bg=COLORS["bg"])

    style = ttk.Style(root)
    # 'clam' is the only built-in theme that reliably honors custom colors
    # across platforms (Windows' native theme ignores most color options).
    style.theme_use("clam")

    base_font = ("Segoe UI", 10)
    bold_font = ("Segoe UI", 10, "bold")
    heading_font = ("Segoe UI", 11, "bold")
    try:
        root.option_add("*Font", base_font)
    except tk.TclError:
        pass

    # ---- generic containers -------------------------------------------------
    style.configure("TFrame", background=COLORS["bg"])
    style.configure("App.TFrame", background=COLORS["bg"])
    style.configure("Card.TFrame", background=COLORS["surface"])

    style.configure("TLabel", background=COLORS["bg"], foreground=COLORS["text"], font=base_font)
    style.configure("Card.TLabel", background=COLORS["surface"], foreground=COLORS["text"], font=base_font)
    style.configure("Muted.TLabel", background=COLORS["bg"], foreground=COLORS["text_muted"], font=base_font)
    style.configure("CardMuted.TLabel", background=COLORS["surface"], foreground=COLORS["text_muted"], font=base_font)
    style.configure("Heading.TLabel", background=COLORS["bg"], foreground=COLORS["text"], font=heading_font)
    style.configure("Status.TLabel", background=COLORS["bg"], foreground=COLORS["primary_dk"], font=bold_font)

    style.configure(
        "Card.TLabelframe", background=COLORS["surface"], bordercolor=COLORS["border"],
        relief="solid", borderwidth=1,
    )
    style.configure(
        "Card.TLabelframe.Label", background=COLORS["surface"], foreground=COLORS["text"], font=bold_font,
    )

    # ---- buttons --------------------------------------------------------------
    style.configure(
        "TButton", background=COLORS["surface"], foreground=COLORS["text"],
        bordercolor=COLORS["border"], borderwidth=1, focusthickness=0,
        padding=(12, 7), font=base_font, relief="flat",
    )
    style.map(
        "TButton",
        background=[("active", COLORS["primary_lt"]), ("disabled", "#F3F4F6")],
        foreground=[("disabled", COLORS["text_muted"])],
    )

    style.configure(
        "Primary.TButton", background=COLORS["primary"], foreground=COLORS["text_on_blue"],
        bordercolor=COLORS["primary"], borderwidth=0, padding=(16, 8), font=bold_font, relief="flat",
    )
    style.map(
        "Primary.TButton",
        background=[("active", COLORS["primary_dk"]), ("disabled", "#B9C7EA")],
        foreground=[("disabled", "#F0F3FC")],
    )

    # ---- inputs -----------------------------------------------------------
    style.configure(
        "TEntry", fieldbackground=COLORS["surface"], background=COLORS["surface"],
        bordercolor=COLORS["border"], foreground=COLORS["text"], padding=6, relief="flat",
    )
    style.map("TEntry", bordercolor=[("focus", COLORS["primary"])])

    style.configure("TCheckbutton", background=COLORS["bg"], foreground=COLORS["text"], font=base_font)
    style.map("TCheckbutton", background=[("active", COLORS["bg"])])
    style.configure("Card.TCheckbutton", background=COLORS["surface"], foreground=COLORS["text"], font=base_font)
    style.map("Card.TCheckbutton", background=[("active", COLORS["surface"])])

    style.configure("TRadiobutton", background=COLORS["surface"], foreground=COLORS["text"], font=base_font)
    style.map("TRadiobutton", background=[("active", COLORS["surface"])])

    # ---- notebook (tabs) ----------------------------------------------------
    style.configure("TNotebook", background=COLORS["bg"], borderwidth=0, bordercolor=COLORS["bg"])
    style.layout("TNotebook", [("TNotebook.client", {"sticky": "nswe"})])
    style.configure(
        "TNotebook.Tab", background=COLORS["bg"], foreground=COLORS["text_muted"],
        padding=(18, 10), font=bold_font, borderwidth=0,
        bordercolor=COLORS["bg"], lightcolor=COLORS["bg"], darkcolor=COLORS["bg"],
    )
    style.map(
        "TNotebook.Tab",
        background=[("selected", COLORS["surface"])],
        foreground=[("selected", COLORS["primary"])],
        bordercolor=[("selected", COLORS["surface"])],
        lightcolor=[("selected", COLORS["surface"])],
        darkcolor=[("selected", COLORS["surface"])],
        expand=[("selected", (1, 1, 1, 0))],
    )
    style.layout("TNotebook.Tab", [
        ("Notebook.tab", {"sticky": "nswe", "children": [
            ("Notebook.padding", {"side": "top", "sticky": "nswe", "children": [
                ("Notebook.label", {"side": "top", "sticky": ""}),
            ]}),
        ]}),
    ])

    # ---- treeview (results table) -----------------------------------------
    style.configure(
        "Treeview", background=COLORS["surface"], fieldbackground=COLORS["surface"],
        foreground=COLORS["text"], bordercolor=COLORS["border"], borderwidth=1,
        rowheight=26, font=base_font,
    )
    style.configure(
        "Treeview.Heading", background=COLORS["primary_lt"], foreground=COLORS["primary_dk"],
        font=bold_font, relief="flat", borderwidth=1,
    )
    style.map("Treeview.Heading", background=[("active", COLORS["primary_lt"])])
    style.map("Treeview", background=[("selected", COLORS["primary"])], foreground=[("selected", "#FFFFFF")])

    style.configure("TScrollbar", background=COLORS["bg"], troughcolor=COLORS["bg"], bordercolor=COLORS["bg"])

    return style


# ══════════════════════════════════════════════════════════════════
#  READ PSD DIMENSIONS WITHOUT OPENING IN PHOTOSHOP
#  (fixed 26-byte header, see Adobe PSD/PSB spec)
# ══════════════════════════════════════════════════════════════════

def read_psd_dimensions(path):
    with open(path, "rb") as f:
        header = f.read(26)
    if len(header) < 26 or header[0:4] != b"8BPS":
        raise ValueError("Not a valid PSD/PSB file: " + path)
    (signature, version, reserved, channels, height, width,
     depth, mode) = struct.unpack(">4sH6sHIIHH", header)
    return {"width": width, "height": height}


def round_half_up(x):
    return math.floor(x + 0.5)


# Bumped every time the placement/plan logic changes -- printed at the top
# of every run so it's easy to tell, from the log alone, whether an old
# copy of this script is what actually ran.
PIPELINE_BUILD = "2026-07-08.design-scale-fix"


# ══════════════════════════════════════════════════════════════════
#  BUILD THE PER-FILE PLAN (identical logic to the CLI version)
# ══════════════════════════════════════════════════════════════════

def has_global_fields(data):
    for page in data["pages"]:
        for tb in page["textBoxes"]:
            if "xRatioGlobal" in tb:
                return True
    return False


def debug_scale_info(files_info, data):
    """
    Returns the raw numbers behind the multi-file target-assignment math,
    purely for logging/diagnosis -- lets us see web_total_h, orig_total_h,
    and design_scale directly instead of only the end result.
    """
    n_files = len(files_info)
    n_pages = len(data["pages"])
    if n_files == n_pages:
        return None
    web_pages = sorted(data["pages"], key=lambda p: p["nativeY0"])
    web_total_h = max(p["nativeY0"] + p["nativeHeight"] for p in web_pages)
    orig_total_h = sum(f["height"] for f in files_info)
    return {
        "web_total_h": web_total_h,
        "orig_total_h": orig_total_h,
        "design_scale": orig_total_h / web_total_h,
    }


def compute_resolved_dims(files_info, resize_width=None):
    """
    Predicts each file's post-resize (w, h) for DISPLAY/LOG purposes only.
    Not used for any placement math (see build_plan) -- position math is
    scale-invariant and works entirely off each file's original dimensions,
    so this prediction being off by a pixel or two from Photoshop's real
    result (which can happen after many independently-rounded files) can
    no longer throw text placement off.
    """
    resolved_dims = []
    for f in files_info:
        if resize_width:
            new_w = resize_width
            new_h = round_half_up(f["height"] * (resize_width / f["width"]))
        else:
            new_w, new_h = f["width"], f["height"]
        resolved_dims.append({"w": new_w, "h": new_h})
    return resolved_dims


def build_plan(files_info, data):
    """
    Figures out, for every text box, which file it belongs to and its
    x/y/font ratios WITHIN that file.

    Two coordinate spaces are involved: the webpage's native pixel space
    (from the extracted JSON) and the PSD files' own pixel space -- these
    are NOT necessarily 1:1 (a PSD's design resolution can differ from the
    page's native capture resolution), so a conversion factor between them
    is required. That factor -- design_scale below -- is computed ONCE
    from each file's ORIGINAL (pre-resize) height, summed exactly (integers
    straight from PSD/PSB headers, no rounding involved at all) against the
    webpage's total native height. It has nothing to do with any later GUI
    resize, so there's no drift and no per-file rounding to compound.

    Target-file assignment and each box's local x/y/font ratios are then
    computed entirely in that original PSD pixel space. Those ratios are
    fractions (scale-invariant): applying them to whatever a file's ACTUAL
    current width/height is (post-resize or not -- read live inside
    Photoshop at placement time) lands text in the same relative spot
    regardless of what any later resize does to the file.
    """
    n_files = len(files_info)
    n_pages = len(data["pages"])

    per_file_boxes = [[] for _ in range(n_files)]

    if n_files == n_pages:
        for i in range(n_files):
            page = data["pages"][i]
            for tb in page["textBoxes"]:
                per_file_boxes[i].append({
                    "tb": tb,
                    "xRatioLocal": tb["xRatio"],
                    "yRatioLocal": tb["yRatio"],
                    "fontRatioLocal": tb.get("fontSizeRatio", 0),
                    "pageLabel": "P" + str(i + 1),
                })
    else:
        if not has_global_fields(data):
            raise ValueError(
                "File count (%d) != JSON page count (%d), and this JSON has no "
                "xRatioGlobal/yRatioGlobal fields. Re-export textboxes.json with "
                "the v8+ browser script." % (n_files, n_pages)
            )

        web_pages = sorted(data["pages"], key=lambda p: p["nativeY0"])
        web_total_h = max(p["nativeY0"] + p["nativeHeight"] for p in web_pages)
        json_native_w = data["totalNativeWidth"]

        # Exact (no rounding) conversion between web-native pixels and each
        # file's ORIGINAL pixel space.
        orig_total_h = sum(f["height"] for f in files_info)
        design_scale = orig_total_h / web_total_h

        # File boundaries in each file's ORIGINAL (pre-resize) pixel space.
        orig_file_meta = []
        cum_y = 0
        for f in files_info:
            orig_file_meta.append({"y0": cum_y, "y1": cum_y + f["height"]})
            cum_y += f["height"]

        for page_idx, page in enumerate(data["pages"]):
            for tb in page["textBoxes"]:
                abs_native_y = tb["yRatioGlobal"] * web_total_h
                abs_native_x = tb["xRatioGlobal"] * json_native_w
                abs_orig_ps_y = abs_native_y * design_scale

                target_idx = n_files - 1
                for i, m in enumerate(orig_file_meta):
                    if m["y0"] <= abs_orig_ps_y < m["y1"]:
                        target_idx = i
                        break

                m = orig_file_meta[target_idx]
                orig_h = m["y1"] - m["y0"]

                # Scale-invariant ratios, all computed in original PSD space.
                x_ratio_local = abs_native_x / json_native_w
                y_ratio_local = (abs_orig_ps_y - m["y0"]) / orig_h
                font_native_px = tb.get("fontSizeRatioGlobal", 0) * web_total_h
                font_ratio_local = (font_native_px * design_scale) / orig_h

                per_file_boxes[target_idx].append({
                    "tb": tb,
                    "xRatioLocal": x_ratio_local,
                    "yRatioLocal": y_ratio_local,
                    "fontRatioLocal": font_ratio_local,
                    "pageLabel": "P" + str(page_idx + 1),
                })

    return per_file_boxes


# ══════════════════════════════════════════════════════════════════
#  EXTENDSCRIPT PAYLOAD (unchanged from the CLI version)
# ══════════════════════════════════════════════════════════════════

JSX_TEMPLATE = r"""
app.displayDialogs = DialogModes.NO;

function cssColorToSolidColor(cssColor) {
  var sc = new SolidColor();
  sc.rgb.red = 0; sc.rgb.green = 0; sc.rgb.blue = 0;
  if (!cssColor) return sc;
  var m = cssColor.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
  if (m) {
    sc.rgb.red   = parseInt(m[1], 10);
    sc.rgb.green = parseInt(m[2], 10);
    sc.rgb.blue  = parseInt(m[3], 10);
  }
  return sc;
}

// ── Box status color -> text color ──────────────────────────────────────
// On the review site, a box's background color indicates its status
// (yellow = approved/normal, red = "Proof", plus green/light-blue used
// for other statuses). The text itself always renders plain black on the
// site regardless of status -- but in Photoshop we want that status to
// stay visible at a glance, so every box EXCEPT yellow gets a text layer
// colored to match its box, instead of everything defaulting to black.
// Classified by hue (not exact RGB match) so small palette variations
// on the site don't fall through the cracks.
function rgbToHue(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  var max = Math.max(r, g, b), min = Math.min(r, g, b), diff = max - min;
  if (diff === 0) return 0;
  var hue;
  if (max === r)      hue = 60 * (((g - b) / diff) %% 6);
  else if (max === g) hue = 60 * (((b - r) / diff) + 2);
  else                 hue = 60 * (((r - g) / diff) + 4);
  if (hue < 0) hue += 360;
  return hue;
}

function textColorForBox(tb) {
  var bg = tb.bgColor;
  if (!bg || bg === 'unknown') return tb.color; // no bg info -- leave as extracted
  var m = bg.match(/rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)/i);
  if (!m) return tb.color;

  // Yellow (approved/normal) boxes keep the site's plain black text --
  // every other status (red "Proof", green, light-blue, ...) gets a text
  // layer colored to match its own box, using hue (not exact RGB match)
  // so small palette variations don't fall through the cracks. This
  // mirrors the same yellow-vs-everything-else rule already used for
  // the other supported site layout.
  var hue = rgbToHue(parseFloat(m[1]), parseFloat(m[2]), parseFloat(m[3]));
  var isYellow = hue >= 40 && hue < 65; // site's "approved" yellow sits ~52°
  return isYellow ? 'rgb(0,0,0)' : bg;
}

function pxToPt(px, dpi) { return px * 72 / dpi; }

// `styleScaleRatio` mirrors what Photoshop's own "Scale Styles" checkbox
// would have done to this stroke's size IF the stroke had already existed
// on the layer at resize time (newWidthPx / originalWidthPx, or 1 when no
// resize happens). We can't rely on scaleStyles for this stroke because
// it's created AFTER the resize (it's added while placing text, not part
// of the source file) -- so without this, the stroke would stay a flat
// 40px no matter how much the canvas shrinks or grows, making it look
// disproportionately thick on a downsized file and thin on an upsized
// one. Multiplying the base size by the same ratio keeps it visually
// consistent with the original design intent at any resize target.
function applyStrokeStyle(targetLayer, styleScaleRatio) {
  var ratio = (styleScaleRatio && styleScaleRatio > 0) ? styleScaleRatio : 1;
  var strokeSizePx = 40 * ratio;
  if (strokeSizePx < 1) strokeSizePx = 1; // never let it vanish entirely on a big downscale

  app.activeDocument.activeLayer = targetLayer;
  var idsetd = charIDToTypeID('setd');
  var desc = new ActionDescriptor();
  var ref = new ActionReference();
  ref.putEnumerated(charIDToTypeID('Lyr '), charIDToTypeID('Ordn'), charIDToTypeID('Trgt'));
  desc.putReference(charIDToTypeID('null'), ref);
  var layerDesc  = new ActionDescriptor();
  var fxDesc     = new ActionDescriptor();
  var strokeDesc = new ActionDescriptor();
  strokeDesc.putUnitDouble(charIDToTypeID('Sz  '), charIDToTypeID('#Pxl'), strokeSizePx);
  strokeDesc.putEnumerated(charIDToTypeID('Styl'), charIDToTypeID('FStl'), charIDToTypeID('OutS'));
  strokeDesc.putEnumerated(charIDToTypeID('Md  '), charIDToTypeID('BlnM'), charIDToTypeID('Nrml'));
  strokeDesc.putUnitDouble(charIDToTypeID('Opct'), charIDToTypeID('#Prc'), 70);
  strokeDesc.putEnumerated(charIDToTypeID('Fl  '), charIDToTypeID('FllT'), charIDToTypeID('SClr'));
  var colorDesc = new ActionDescriptor();
  colorDesc.putDouble(charIDToTypeID('Rd  '), 255);
  colorDesc.putDouble(charIDToTypeID('Grn '), 185);
  colorDesc.putDouble(charIDToTypeID('Bl  '), 147);
  strokeDesc.putObject(charIDToTypeID('Clr '), charIDToTypeID('RGBC'), colorDesc);
  strokeDesc.putBoolean(charIDToTypeID('enab'), true);
  fxDesc.putObject(charIDToTypeID('FrFX'), charIDToTypeID('FrFX'), strokeDesc);
  layerDesc.putObject(charIDToTypeID('Lefx'), charIDToTypeID('Lefx'), fxDesc);
  desc.putObject(charIDToTypeID('T   '), charIDToTypeID('Lyr '), layerDesc);
  executeAction(idsetd, desc, DialogModes.NO);
}

var MIN_FONT_PT = 6;

function placeOneBox(doc, tb, xRatio, yRatio, fontSizeRatio, pageLabel, group) {
  // docW/docH are read live, right here, every time this runs -- i.e.
  // AFTER the resize below, if one happened. xRatio/yRatio/fontSizeRatio
  // are plain fractions (scale-invariant), computed in Python purely from
  // each file's ORIGINAL dimensions -- so multiplying them by whatever
  // this doc's actual current size is always lands text in the right
  // relative spot, regardless of what the resize below does to it.
  var docW   = doc.width.as('px');
  var docH   = doc.height.as('px');
  var docDPI = doc.resolution;

  var targetX = xRatio * docW;
  var targetY = yRatio * docH;

  var fontPt;
  if (fontSizeRatio && fontSizeRatio > 0) {
    fontPt = pxToPt(fontSizeRatio * docH, docDPI);
    if (fontPt < MIN_FONT_PT) fontPt = MIN_FONT_PT;
  } else {
    fontPt = 13;
  }

  var fontPx    = fontPt * docDPI / 72;
  var yBaseline = targetY + fontPx * 0.72;

  var psText = tb.text.replace(/\n/g, '\r');

  var layer = doc.artLayers.add();
  layer.kind = LayerKind.TEXT;
  var tl = layer.textItem;
  tl.kind = TextType.POINTTEXT;

  var fontCandidates = ['MarkaziText-Regular', 'MarkaziText', 'ArialMT', 'Arial-BoldMT', 'Arial'];
  for (var fi = 0; fi < fontCandidates.length; fi++) {
    try { tl.font = fontCandidates[fi]; break; } catch(fe) {}
  }

  tl.size  = new UnitValue(fontPt, 'pt');
  tl.color = cssColorToSolidColor(textColorForBox(tb));
  tl.contents = psText;

  var align = (tb.textAlign || 'center').toLowerCase();
  if      (align === 'left')    { tl.justification = Justification.LEFT; }
  else if (align === 'right')   { tl.justification = Justification.RIGHT; }
  else if (align === 'justify') { tl.justification = Justification.FULLJUSTIFYLASTLEFT; }
  else                          { tl.justification = Justification.CENTER; }

  var cur = tl.position;
  layer.translate(
    new UnitValue(targetX   - cur[0].as('px'), 'px'),
    new UnitValue(yBaseline - cur[1].as('px'), 'px')
  );

  if (tb.opacity !== undefined && tb.opacity < 1) {
    layer.opacity = Math.round(tb.opacity * 100);
  }

  if (tb.layerName) {
    layer.name = tb.layerName;
  } else if (tb.mark) {
    layer.name = tb.mark + '_' + tb.text.replace(/\n/g, ' ').slice(0, 40);
  } else {
    layer.name = tb.text;
  }
  if (group) layer.move(group, ElementPlacement.PLACEATBEGINNING);
}

// ---- missing-linked-asset detection & auto-rasterize --------------------
// Photoshop flags a Smart Object layer with a small warning badge when its
// linked file (e.g. a Creative Cloud Library graphic) can't be found. That
// same flag is readable through Action Manager as smartObjectMore.warning.
// If a layer has it, we rasterize it (turns it into a normal pixel layer
// using its last known appearance) so the file can keep processing instead
// of erroring out on that layer.
function isSmartObjectLinkMissing(layer) {
  try {
    var ref = new ActionReference();
    ref.putProperty(charIDToTypeID('Prpr'), stringIDToTypeID('smartObjectMore'));
    ref.putIdentifier(charIDToTypeID('Lyr '), layer.id);
    var desc = executeActionGet(ref);
    var key = stringIDToTypeID('smartObjectMore');
    if (desc.hasKey(key)) {
      var soDesc = desc.getObjectValue(key);
      var warnKey = stringIDToTypeID('warning');
      if (soDesc.hasKey(warnKey)) {
        return soDesc.getBoolean(warnKey);
      }
    }
  } catch (e) { /* property unavailable on this PS version -- treat as not missing */ }
  return false;
}

function rasterizeMissingLinkedSmartObjects(doc, notes) {
  function walk(layerList) {
    for (var i = 0; i < layerList.length; i++) {
      var layer = layerList[i];
      var isGroup = false;
      try { isGroup = layer.typename === 'LayerSet'; } catch (eg) {}
      if (isGroup) { walk(layer.layers); continue; }

      var kind = null;
      try { kind = layer.kind; } catch (ek) {}
      if (kind === LayerKind.SMARTOBJECT && isSmartObjectLinkMissing(layer)) {
        var nm = layer.name;
        try {
          layer.rasterize(RasterizeType.ENTIRELAYER);
          notes.push('NOTE: rasterized Smart Object layer with a missing linked asset: "' + nm + '"');
        } catch (rasErr) {
          notes.push('WARNING: layer "' + nm + '" has a missing linked asset and could not be auto-rasterized: ' + rasErr.message);
        }
      }
    }
  }
  walk(doc.layers);
}

function rasterizeAllSmartObjects(doc, notes) {
  // Last-resort fallback for when isSmartObjectLinkMissing()'s detection
  // doesn't match on this Photoshop version -- rasterizes every Smart
  // Object in the doc (not just ones we could confirm are broken) so a
  // later operation (like resizeImage) can't still trip over the same
  // missing-link error.
  function walk(layerList) {
    for (var i = 0; i < layerList.length; i++) {
      var layer = layerList[i];
      var isGroup = false;
      try { isGroup = layer.typename === 'LayerSet'; } catch (eg) {}
      if (isGroup) { walk(layer.layers); continue; }
      var kind = null;
      try { kind = layer.kind; } catch (ek) {}
      if (kind === LayerKind.SMARTOBJECT) {
        var nm = layer.name;
        try {
          layer.rasterize(RasterizeType.ENTIRELAYER);
          notes.push('NOTE: rasterized Smart Object layer "' + nm + '" (fallback pass).');
        } catch (rasErr) { /* nothing more we can do about this one */ }
      }
    }
  }
  walk(doc.layers);
}

// doc.resizeImage(...) -- the simple DOM method used further below -- has
// NO way to also scale layer effects (Stroke, Drop Shadow, Inner Shadow,
// Outer Glow, Bevel & Emboss, etc). Without it, a 5px stroke stays a flat
// 5px no matter how much the canvas shrinks or grows, so it ends up looking
// far too thick (on a shrink) or too thin (on an enlarge) relative to the
// new image size -- exactly the "scale styles" checkbox in Photoshop's own
// Image Size dialog. That checkbox isn't exposed on the Document object at
// all; it only exists as a flag on the low-level Image Size ActionManager
// call, so we drive that action directly here to get the same result a
// human checking the box in the UI would get.
//
// `includeInterpolation` controls whether we also force the resample
// method (Bilinear) via the same call. Photoshop moved the resample-method
// enum from old 4-char codes to string IDs at some point, and passing the
// wrong one for a given Photoshop version can crash the ActionManager call
// at a level ExtendScript's try/catch can't intercept (it never reaches
// our catch block at all -- Photoshop just silently aborts the whole
// script). So resizeImageWithScaledStyles() is called with the resample
// method first; if that particular call is what's failing, the retry
// below drops `includeInterpolation` and just lets Photoshop use whatever
// resample method it was last set to -- scaleStyles/dimensions/resolution
// still get applied either way, only the exact resample algorithm is
// no longer guaranteed to be Bilinear specifically.
function resizeImageWithScaledStyles(doc, newWidthPx, newHeightPx, dpi, includeInterpolation) {
  var idImgS = charIDToTypeID('ImgS');
  var desc = new ActionDescriptor();
  desc.putUnitDouble(charIDToTypeID('Wdth'), charIDToTypeID('#Pxl'), newWidthPx);
  desc.putUnitDouble(charIDToTypeID('Hght'), charIDToTypeID('#Pxl'), newHeightPx);
  desc.putUnitDouble(charIDToTypeID('Rslt'), charIDToTypeID('#Rsl'), dpi);
  desc.putBoolean(stringIDToTypeID('scaleStyles'), true);   // <-- the important part
  desc.putBoolean(charIDToTypeID('CnsP'), true);            // constrain proportions
  if (includeInterpolation) {
    // Modern Photoshop identifies resample methods by string ID, not the
    // old 4-char codes (those were retired around when "Preserve Details"
    // upscaling was added) -- 'bilinear' is the current string ID.
    desc.putEnumerated(charIDToTypeID('Intr'), charIDToTypeID('Intp'), stringIDToTypeID('bilinear'));
  }
  executeAction(idImgS, desc, DialogModes.NO);
}

// `scaleStyles` (above) only scales layer EFFECTS (stroke, drop shadow,
// glow, bevel) -- it has no equivalent for Smart Filters (a real
// Gaussian Blur, Motion Blur, Smart Blur, etc. applied to a Smart
// Object's pixels, shown as "Smart Filters" under the layer in the
// Layers panel). Photoshop does not scale these on resize, ever --
// there's no checkbox for it, scripted or otherwise. A Motion Blur
// stored with e.g. a 125px distance stays a live filter that keeps
// re-rendering at that same absolute 125px on whatever canvas size the
// document currently is -- including after a resize -- so it ends up
// looking proportionally far stronger on a downsized file.
//
// The first thing we tried here was reading the Smart Filter's raw
// Action Manager descriptor (smartObject -> filterFX -> filterFXList ->
// Fltr -> Dstn/Rds ) and writing back a scaled number. That part of the
// descriptor schema is real (confirmed straight from a PSD's own binary
// data), and the write call completes without throwing -- but Smart
// Filters cache their own rendered result, and Photoshop does not
// reliably re-render that cache from a scripted property write; only
// its interactive filter dialog does that reliably. So the number can
// change internally while the visible blur stays exactly as it was --
// which is exactly what happened.
//
// What DOES reliably work is resizeImage()'s own pixel resampling: it
// uniformly rescales whatever raster pixels are on a layer, live filter
// or not. So instead of editing the filter's parameters, this rasterizes
// (bakes) any Smart Object layer that has a scalable filter -- Motion
// Blur distance or Gaussian/Surface/Smart Blur radius -- BEFORE the
// canvas resize happens. That freezes the filter at its
// currently-designed look as flat pixels; the resize immediately after
// then scales those pixels down/up along with everything else on the
// canvas, so the blur ends up proportionally correct automatically --
// the same way it would if you flattened the file and resized it by
// hand. Rather than rasterizing the Smart Object layer itself (which
// would permanently lose the live, re-editable filter), this
// duplicates the layer, rasterizes only the copy, and hides the
// original underneath -- so the resized/exported file still shows the
// correctly-scaled baked blur, but the original live Smart Object with
// its editable filter stays in the document (hidden) if anyone needs
// to go back to it later.
function rasterizeScalableFilterLayers(doc, notes) {
  var sID = stringIDToTypeID, cID = charIDToTypeID;
  var SO_KEY   = sID('smartObject');
  var FX_KEY   = sID('filterFX');
  var LIST_KEY = sID('filterFXList');
  var FLTR_KEY = sID('Fltr');
  var DSTN_KEY = cID('Dstn');   // Motion Blur / Radial Blur distance
  var RDS_KEY  = cID('Rds ');   // Gaussian / Surface / Smart Blur radius

  function hasScalableFilter(layer) {
    app.activeDocument.activeLayer = layer;
    var ref = new ActionReference();
    ref.putProperty(cID('Prpr'), SO_KEY);
    ref.putEnumerated(cID('Lyr '), cID('Ordn'), cID('Trgt'));
    var layerGet = executeActionGet(ref);
    if (!layerGet.hasKey(SO_KEY)) return false; // no smartObject data on this layer

    var soDesc = layerGet.getObjectValue(SO_KEY);
    if (!soDesc.hasKey(FX_KEY)) return false; // no Smart Filters at all

    var fxDesc = soDesc.getObjectValue(FX_KEY);
    if (!fxDesc.hasKey(LIST_KEY)) return false;

    var list = fxDesc.getList(LIST_KEY);
    for (var li = 0; li < list.count; li++) {
      var itemDesc = list.getObjectValue(li);
      if (!itemDesc.hasKey(FLTR_KEY)) continue;
      var fltrDesc = itemDesc.getObjectValue(FLTR_KEY);
      if (fltrDesc.hasKey(DSTN_KEY) || fltrDesc.hasKey(RDS_KEY)) return true;
    }
    return false;
  }

  // Duplicates `layer` directly above itself, hides + renames the
  // original underneath (so it's obviously the untouched, still-live
  // Smart Object if anyone opens the file later), and rasterizes only
  // the duplicate. The duplicate -- flat pixels, correctly baking in
  // whatever the blur currently looks like -- is what the canvas
  // resize right after this function scales along with everything
  // else. Nothing about the original Smart Object's filter is ever
  // touched, so re-running this on the same file later won't find a
  // "changed" filter to worry about; it'll just see the original
  // (still hidden, still live) and the already-rasterized duplicate
  // (no longer LayerKind.SMARTOBJECT, so it's skipped automatically).
  function duplicateHideAndRasterize(layer, nm) {
    var dup = layer.duplicate(layer, ElementPlacement.PLACEBEFORE);
    dup.name = nm;
    // Photoshop's duplicate() does not reliably mirror the source
    // layer's visibility onto the new copy -- explicitly force it on so
    // this never silently depends on that behavior.
    dup.visible = true;
    layer.visible = false;
    try { layer.name = nm + ' (original, hidden)'; } catch (eRename) {}
    // Re-assert after the rename -- belt and suspenders against the same
    // kind of visibility drift this whole fix is guarding against.
    layer.visible = false;
    dup.rasterize(RasterizeType.ENTIRELAYER);
    return dup;
  }

  // Pass 1: just collect Smart Object layer references. Duplicating
  // layers changes the length/order of layerList arrays as we go, so
  // walking and mutating in the same pass risks revisiting a layer
  // (e.g. re-processing the newly-hidden original a second time and
  // duplicating it again). A plain object reference in this array
  // stays valid and correct no matter how the stack is rearranged
  // afterward.

  // Matches the " (original, hidden)" suffix duplicateHideAndRasterize()
  // stamps onto a Smart Object it has already baked, e.g. "SFX (original,
  // hidden)".
  var ALREADY_PROCESSED_RE = / \(original, hidden\)$/;

  function collect(layerList, out) {
    for (var i = 0; i < layerList.length; i++) {
      var layer = layerList[i];
      var isGroup = false;
      try { isGroup = layer.typename === 'LayerSet'; } catch (eg) {}
      if (isGroup) { collect(layer.layers, out); continue; }

      // Skip a Smart Object this same function already processed on a
      // previous run of the tool (hidden + renamed here, rasterized copy
      // left visible above it). Re-processing it would duplicate it
      // *again* -- and Photoshop's duplicate() does not reliably keep the
      // new copy hidden to match a hidden source, so the extra duplicate
      // tends to come out visible, producing a doubled/ghosted layer on
      // top of the one baked in last time.
      var nmCheck = '';
      try { nmCheck = layer.name; } catch (enc) {}
      if (!layer.visible && ALREADY_PROCESSED_RE.test(nmCheck)) continue;

      var kind = null;
      try { kind = layer.kind; } catch (ek) {}
      if (kind === LayerKind.SMARTOBJECT) out.push(layer);
    }
  }

  var candidates = [];
  collect(doc.layers, candidates);

  for (var ci = 0; ci < candidates.length; ci++) {
    var layer = candidates[ci];
    var nm = '(unnamed)';
    try { nm = layer.name; } catch (enm) {}

    // Detection first (does this Smart Object have a Motion Blur /
    // Gaussian Blur / Surface Blur / Smart Blur filter?). On some
    // Smart Object subtypes (e.g. ones created via "Convert to Smart
    // Object" on a group, rather than from a placed file) this
    // Action Manager descriptor read throws a generic, unhelpful
    // Photoshop error ("General Photoshop error occurred...") even
    // though the layer is perfectly normal otherwise.
    var scalable = false;
    var detectFailed = false;
    try {
      scalable = hasScalableFilter(layer);
    } catch (eDetect) {
      detectFailed = true;
    }

    if (detectFailed) {
      // Can't tell whether this layer has a scalable Smart Filter --
      // err on the side of baking in a rasterized copy anyway. A
      // needless duplicate/rasterize just leaves an extra hidden
      // layer behind; skipping one that did have a blur would ship a
      // file where that blur silently doesn't scale.
      try {
        duplicateHideAndRasterize(layer, nm);
        notes.push('NOTE: duplicated Smart Object layer "' + nm + '", rasterized the copy, and hid the original underneath, as a precaution before resizing -- Photoshop would not let this script check it for Motion Blur/blur Smart Filters directly, so it was baked in on a copy rather than risk shipping an unscaled blur. The original live Smart Object is preserved in the file, just hidden.');
      } catch (eForce) {
        notes.push('NOTE: could not check, or duplicate/rasterize, Smart Filter layer "' + nm + '" (' + eForce.message + '). If this layer has a Motion Blur/Gaussian Blur/etc Smart Filter, it will NOT scale with this resize.');
      }
      continue;
    }

    if (scalable) {
      try {
        duplicateHideAndRasterize(layer, nm);
        notes.push('NOTE: duplicated Smart Object layer "' + nm + '" and rasterized the copy to bake in its Motion Blur/blur filter before resizing, so the blur scales proportionally with the canvas. The original live Smart Object is preserved in the file, just hidden underneath.');
      } catch (eRas) {
        notes.push('NOTE: detected a scalable Smart Filter on layer "' + nm + '" but could not duplicate/rasterize it (' + eRas.message + '). Its blur will NOT scale with this resize.');
      }
    }
  }
}

// ---- injected per-file parameters ----
var FILE_PATH   = %(FILE_PATH)s;
var DO_RESIZE   = %(DO_RESIZE)s;
var RESIZE_W    = %(RESIZE_W)s;
var RESIZE_DPI  = %(RESIZE_DPI)s;
var BOX_LIST    = %(BOX_LIST)s;

// ---- run: resize FIRST (if requested), THEN place text -- a single
// open/close per file. This is safe because BOX_LIST's ratios were built
// entirely from each file's ORIGINAL dimensions (scale-invariant), not
// from a predicted post-resize size -- so it doesn't matter whether the
// resize below produces exactly what was predicted or not; placeOneBox()
// always re-reads the doc's actual current size to compute pixel targets.
var result = { ok: 0, total: BOX_LIST.length, errors: [] };
var doc = null;
try {
  var targetFsName = new File(FILE_PATH).fsName;

  try {
    doc = app.open(new File(FILE_PATH));
  } catch (openErr) {
    // Photoshop throws "...linked smart object file is missing" partway
    // through opening a file with a broken linked asset -- but the
    // document is very often already open by the time it throws. Look
    // for it in app.documents and recover it instead of aborting the
    // whole file.
    var recovered = null;
    try {
      for (var di = 0; di < app.documents.length; di++) {
        var candidate = app.documents[di];
        try {
          if (candidate.fullName && candidate.fullName.fsName === targetFsName) {
            recovered = candidate;
            break;
          }
        } catch (eMatch) {}
      }
      if (!recovered && app.documents.length > 0) {
        // fullName match failed (unusual) but a doc did open -- best guess
        // is the active one, since this file was just what we asked for.
        try { recovered = app.activeDocument; } catch (eActive) {}
      }
    } catch (eScan) {}

    if (recovered) {
      doc = recovered;
      result.errors.push('NOTE: Photoshop reported a missing linked asset while opening -- recovered the document and will auto-rasterize the affected layer(s).');
    } else {
      // Genuinely didn't open at all -- one retry, since this can be transient.
      doc = app.open(new File(FILE_PATH));
    }
  }

  app.activeDocument = doc;

  rasterizeMissingLinkedSmartObjects(doc, result.errors);

  var prevRuler = app.preferences.rulerUnits;
  app.preferences.rulerUnits = Units.PIXELS;

  // Ratio between the post-resize and pre-resize width -- reused below to
  // keep the text stroke (added after resize, so scaleStyles can't touch
  // it) visually consistent at any resize target. Stays 1 when no resize
  // happens, which preserves the original flat-40px stroke behavior.
  var STYLE_SCALE_RATIO = 1;

  if (DO_RESIZE) {
    // Bake any Motion Blur/blur Smart Filters into flat pixels BEFORE
    // resizing -- see rasterizeScalableFilterLayers() above for why this
    // (rather than trying to rewrite the filter's live parameters) is
    // what actually makes the blur end up proportionally correct.
    try {
      rasterizeScalableFilterLayers(doc, result.errors);
    } catch (sfErr) {
      result.errors.push('NOTE: could not check/rasterize Smart Filter layers ("' + sfErr.message + '").');
    }

    var origW = doc.width.as('px');
    var origH = doc.height.as('px');
    var newW  = RESIZE_W;
    var newH  = Math.round(origH * (newW / origW));
    STYLE_SCALE_RATIO = newW / origW;
    try {
      resizeImageWithScaledStyles(doc, newW, newH, RESIZE_DPI, true);
    } catch (resizeErr) {
      try {
        // Attempt 2: same idea, but without forcing a specific resample
        // method -- in case THAT part of the descriptor is what this
        // Photoshop version doesn't like.
        result.errors.push('NOTE: resize-with-scaled-styles failed ("' + resizeErr.message + '") -- retrying without forcing a resample method.');
        resizeImageWithScaledStyles(doc, newW, newH, RESIZE_DPI, false);
      } catch (resizeErr2) {
        try {
          // Attempt 3: rasterize Smart Objects (covers the "missing linked
          // asset" case this retry loop originally existed for) and try
          // the full scaled-styles call once more.
          result.errors.push('NOTE: still failing ("' + resizeErr2.message + '") -- rasterizing all Smart Objects and retrying once more.');
          rasterizeAllSmartObjects(doc, result.errors);
          resizeImageWithScaledStyles(doc, newW, newH, RESIZE_DPI, true);
        } catch (resizeErr3) {
          // Last resort: fall back to the plain resize that's always
          // worked (just without layer effects being scaled), so the
          // file still gets resized and text still gets placed instead
          // of the whole run aborting.
          result.errors.push('NOTE: scaled-styles resize could not be made to work on this Photoshop version ("' + resizeErr3.message + '") -- falling back to a plain resize. Layer effect sizes (stroke/glow/shadow) will NOT be scaled for this file.');
          doc.resizeImage(new UnitValue(newW, 'px'), new UnitValue(newH, 'px'),
                           RESIZE_DPI, ResampleMethod.BILINEAR);
        }
      }
    }
  }

  var group = null;
  try { group = doc.layerSets.add(); group.name = 'TR'; applyStrokeStyle(group, STYLE_SCALE_RATIO); }
  catch (ge) { group = null; }

  for (var i = 0; i < BOX_LIST.length; i++) {
    var entry = BOX_LIST[i];
    try {
      placeOneBox(doc, entry.tb, entry.xRatioLocal, entry.yRatioLocal,
                  entry.fontRatioLocal, entry.pageLabel, group);
      result.ok++;
    } catch (err) {
      result.errors.push(entry.pageLabel + ' Box#' + entry.tb.id + ': ' + err.message);
    }
  }

  app.preferences.rulerUnits = prevRuler;

  // Try a normal save first. A .psd file has a hard 2GB data limit; once a
  // (resized-up, or just large/many-layered) file crosses that, Photoshop
  // throws here instead of saving. In that case, fall back to saving as
  // Large Document Format (.psb), which has no practical size limit, and
  // remove the old .psd so it doesn't linger alongside the replacement.
  try {
    doc.close(SaveOptions.SAVECHANGES);
    doc = null;
  } catch (saveErr) {
    var savedAsPsb = false;
    try {
      var psbPath = FILE_PATH.replace(/\.psd$/i, '.psb');
      if (psbPath === FILE_PATH) { psbPath = FILE_PATH + '.psb'; }

      var psbOptions = new LargeDocumentFormatSaveOptions();
      psbOptions.layers = true;
      psbOptions.alphaChannels = true;
      psbOptions.spotColors = true;
      psbOptions.embedColorProfile = true;

      doc.saveAs(new File(psbPath), psbOptions, false);
      doc.close(SaveOptions.DONOTSAVECHANGES);
      doc = null;

      var oldPsdFile = new File(FILE_PATH);
      if (psbPath !== FILE_PATH && oldPsdFile.exists) {
        oldPsdFile.remove();
      }

      savedAsPsb = true;
      result.errors.push('NOTE: file exceeded the 2GB .psd limit -- saved as ' + psbPath + ' instead, and removed the old .psd.');
    } catch (psbErr) {
      result.errors.push('FATAL: could not save as .psd (' + saveErr.message + ') or as .psb (' + psbErr.message + ')');
      try { if (doc) doc.close(SaveOptions.DONOTSAVECHANGES); } catch (e2) {}
      doc = null;
    }
  }

} catch (fatal) {
  result.errors.push('FATAL: ' + fatal.message);
  try { if (doc) doc.close(SaveOptions.DONOTSAVECHANGES); } catch (e2) {}
}

// Photoshop's ExtendScript engine has no built-in JSON object, so we can't
// use JSON.stringify() here. Build a simple delimited string instead,
// using control characters that will never appear in normal text.
function encodeResult(r) {
  var FS = String.fromCharCode(1); // separates ok / total / errors-blob
  var RS = String.fromCharCode(2); // separates individual error messages
  var errStr = '';
  for (var i = 0; i < r.errors.length; i++) {
    if (i > 0) errStr += RS;
    errStr += String(r.errors[i]).replace(/\r|\n/g, ' ');
  }
  return r.ok + FS + r.total + FS + errStr;
}

encodeResult(result);
"""


def js_string(s):
    return json.dumps(s)


def parse_result(raw):
    """Parses the FS/RS-delimited string produced by encodeResult() in the JSX payload."""
    try:
        FS = chr(1)
        RS = chr(2)
        parts = str(raw).split(FS)
        ok = int(parts[0])
        total = int(parts[1])
        err_blob = parts[2] if len(parts) > 2 else ""
        errors = err_blob.split(RS) if err_blob else []
        return {"ok": ok, "total": total, "errors": errors}
    except Exception:
        return {"ok": 0, "total": 0, "errors": ["Could not parse PS result: " + str(raw)]}


# ══════════════════════════════════════════════════════════════════
#  MAC PHOTOSHOP BRIDGE
#  win32com's DoJavaScript() just hands a chunk of ExtendScript text to
#  Photoshop and gets a string back -- Photoshop understands that exact
#  same JS on macOS too, it's just reached a different way: via macOS's
#  JXA (JavaScript for Automation), driven through the built-in
#  `osascript` command line tool instead of COM. This class exposes the
#  same one method win32com's COM object is actually used for below
#  (.DoJavaScript), so run_one_file() and the rest of the pipeline don't
#  need to know which platform they're on.
# ══════════════════════════════════════════════════════════════════

        # NOTE: JXA's Application().doJavaScript() is unreliable against
        # some Photoshop versions and fails with a generic AppleEvent
        # error (-1708 "Message not understood") even though the exact
        # same script runs fine through plain AppleScript's "do
        # javascript" command against the same Photoshop process. So the
        # bridge below uses plain AppleScript instead of JXA: it writes
        # the JS to a temp .js file (avoiding any quoting/escaping
        # problems with large scripts) and drives Photoshop with
        # `do javascript (POSIX file ...)`, run through a small
        # AppleScript wrapper file passed to osascript.
def _mac_applescript_runner(proc_name, js_source):
    # We switched from passing the JS as a *file* (`POSIX file jsPath as
    # alias`) to embedding it directly as an inline string. The file-based
    # form compiles and sends fine, but was confirmed (via manual Script
    # Editor testing) to fail at runtime on this Photoshop build with a
    # generic, contentless "(8800) General Photoshop error occurred" --
    # while the plain inline-string form of `do javascript` works.
    # AppleScript double-quoted strings can safely contain literal
    # newlines (the compiler just looks for the closing quote), so
    # multi-line JS embeds fine; we only need to escape backslashes and
    # double quotes.
    escaped_proc = proc_name.replace("\\", "\\\\").replace('"', '\\"')
    escaped_js = js_source.replace("\\", "\\\\").replace('"', '\\"')
    return (
        'tell application "%s"\n'
        '    do javascript "%s"\n'
        'end tell\n'
    ) % (escaped_proc, escaped_js)


def _ensure_js_warnings_disabled(proc_name):
    """
    Makes sure Photoshop's PSUserConfig.txt has WarnRunningScripts 0 set,
    which suppresses the "a script wants to run" warning dialog Photoshop
    shows by default for scripts triggered from outside its own File >
    Scripts menu (exactly our case, since we trigger via AppleScript). With
    nothing available to click "OK" on that dialog, do-javascript calls can
    silently fail with a generic, contentless "(8800) General Photoshop
    error occurred".

    Returns True if it just created/modified the file (meaning Photoshop
    needs to be restarted for the change to take effect), False if the
    setting was already correctly present.
    """
    prefs_dir = os.path.expanduser("~/Library/Preferences")
    settings_dir = os.path.join(prefs_dir, f"{proc_name} Settings")

    if not os.path.isdir(settings_dir):
        candidates = glob.glob(os.path.join(prefs_dir, "Adobe Photoshop*Settings"))
        if candidates:
            settings_dir = candidates[0]
        else:
            os.makedirs(settings_dir, exist_ok=True)

    config_path = os.path.join(settings_dir, "PSUserConfig.txt")

    existing = ""
    if os.path.isfile(config_path):
        with open(config_path, "r", encoding="utf-8", errors="ignore") as f:
            existing = f.read()

    if "WarnRunningScripts" in existing and "WarnRunningScripts 0" in existing.replace("\t", " "):
        return False

    if "WarnRunningScripts" in existing:
        lines = [ln for ln in existing.splitlines() if "WarnRunningScripts" not in ln]
        new_content = "\n".join(lines).rstrip() + "\n"
    else:
        new_content = existing.rstrip()
        if new_content:
            new_content += "\n"
    new_content += "# Disable Javascript Warnings\nWarnRunningScripts 0\n"

    with open(config_path, "w", encoding="utf-8") as f:
        f.write(new_content)
    return True


class MacPhotoshopApp:
    def __init__(self, process_name):
        self.process_name = process_name

    def DoJavaScript(self, script):
        if _ensure_js_warnings_disabled(self.process_name):
            raise RuntimeError(
                "Adjusted a Photoshop setting (disabled the script warning "
                "dialog) that requires a restart to take effect. Please "
                "fully quit Adobe Photoshop and reopen it, then try again."
            )

        # No separate js_path file / argv passing anymore -- the JS is
        # embedded directly in the AppleScript as an inline string (see
        # note in _mac_applescript_runner). Still use /tmp explicitly for
        # the .applescript file itself: tempfile.mkstemp()'s default
        # $TMPDIR is a private per-process folder under /var/folders/...
        # that osascript (a separate process) may not reliably read.
        runner_fd, runner_path = tempfile.mkstemp(suffix=".applescript", dir="/tmp")
        try:
            with os.fdopen(runner_fd, "w", encoding="utf-8") as f:
                f.write(_mac_applescript_runner(self.process_name, script))
            proc = subprocess.run(
                ["osascript", runner_path],
                capture_output=True, text=True, timeout=600,
            )
        finally:
            try:
                os.remove(runner_path)
            except OSError:
                pass
        if proc.returncode != 0:
            raise RuntimeError(
                "Photoshop JavaScript execution failed: %s"
                % (proc.stderr.strip() or proc.stdout.strip() or "unknown osascript error")
            )
        return proc.stdout.rstrip("\n")


def connect_mac_photoshop():
    """
    Finds an already-running Photoshop process via System Events -- this
    mirrors win32com's GetActiveObject() behavior on Windows (attach to
    whatever copy is already open, don't launch a new one). Tab 1 already
    requires Photoshop to be open before clicking Run on Windows, so the
    same requirement on Mac is not a new restriction.
    """
    try:
        proc = subprocess.run(
            ["osascript", "-e",
             'tell application "System Events" to get name of every process '
             'whose name contains "Photoshop"'],
            capture_output=True, text=True, timeout=15,
        )
    except Exception as e:
        raise RuntimeError("Could not query running applications: %s" % e)

    if proc.returncode != 0:
        raise RuntimeError(
            "Could not query running applications via System Events: %s"
            % (proc.stderr.strip() or "unknown error")
        )

    names = [n.strip() for n in proc.stdout.split(",") if n.strip()]
    if not names:
        raise RuntimeError(
            "Photoshop does not appear to be running. Make sure Photoshop is "
            "already open, then click Run again."
        )
    return MacPhotoshopApp(names[0])


# ══════════════════════════════════════════════════════════════════
#  "CANNOT LOCATE LINKED ASSETS" WATCHDOG
#  When a PSD has a broken linked Smart Object (e.g. a Creative Cloud
#  Library graphic that's since been moved/deleted), Photoshop pops a
#  modal "Cannot locate linked assets" dialog the moment the file is
#  opened -- and this dialog is NOT suppressed by app.displayDialogs,
#  since it's a document-integrity warning rather than an ordinary
#  scripting alert. Because it's modal, it blocks Photoshop's UI thread,
#  which blocks the DoJavaScript() call this script is waiting on, which
#  is what causes a run to stall/skip that file.
#
#  Fix, in two parts:
#   1) A small background thread runs alongside every DoJavaScript()
#      call, watching for that exact dialog and dismissing it (pressing
#      Enter, which lands on the default "OK" button) the instant it
#      appears -- so the open completes instead of hanging.
#   2) Once the file is open, run_one_file's JSX walks every layer and
#      rasterizes any Smart Object Photoshop is flagging with a missing
#      linked file, so it turns into a normal pixel layer using its last
#      known appearance and the run continues instead of erroring out
#      on that layer.
# ══════════════════════════════════════════════════════════════════

def _dismiss_ps_dialogs_windows(stop_event, poll_interval=0.3):
    """Watches for a Photoshop dialog window (title is just "Adobe
    Photoshop", distinct from the main app window) and presses Enter to
    accept its default button. Runs in a background thread for the
    duration of one DoJavaScript() call."""
    if win32gui is None:
        return
    while not stop_event.is_set():
        try:
            def _cb(hwnd, _):
                if not win32gui.IsWindowVisible(hwnd):
                    return
                title = (win32gui.GetWindowText(hwnd) or "").strip()
                if title.lower() == "adobe photoshop":
                    try:
                        win32gui.SetForegroundWindow(hwnd)
                    except Exception:
                        pass
                    try:
                        win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
                        win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
                    except Exception:
                        pass
            win32gui.EnumWindows(_cb, None)
        except Exception:
            pass
        stop_event.wait(poll_interval)


_MAC_DISMISS_DIALOG_SCRIPT = r"""
on run argv
    set procName to item 1 of argv
    tell application "System Events"
        if exists process procName then
            tell process procName
                repeat with w in windows
                    try
                        if exists (button "OK" of w) then
                            click button "OK" of w
                        end if
                    end try
                end repeat
            end tell
        end if
    end tell
end run
"""


def _dismiss_ps_dialogs_mac(process_name, stop_event, poll_interval=0.3):
    """Mac equivalent: uses System Events UI scripting to click "OK" on
    any Photoshop dialog. Requires the app running this script (Terminal,
    or the built .app) to have Accessibility permission granted in
    System Settings -> Privacy & Security -> Accessibility -- macOS will
    prompt for this automatically the first time it's needed."""
    while not stop_event.is_set():
        try:
            subprocess.run(
                ["osascript", "-e", _MAC_DISMISS_DIALOG_SCRIPT, process_name],
                capture_output=True, text=True, timeout=10,
            )
        except Exception:
            pass
        stop_event.wait(poll_interval)


def _run_with_dialog_watchdog(ps_app, fn):
    """Runs fn() (a DoJavaScript call) while a background thread dismisses
    any blocking Photoshop dialog that pops up during it."""
    stop_event = threading.Event()
    watcher = None
    if sys.platform.startswith("win") and win32gui is not None:
        watcher = threading.Thread(target=_dismiss_ps_dialogs_windows, args=(stop_event,), daemon=True)
    elif sys.platform == "darwin":
        process_name = getattr(ps_app, "process_name", None)
        if process_name:
            watcher = threading.Thread(
                target=_dismiss_ps_dialogs_mac, args=(process_name, stop_event), daemon=True
            )
    if watcher:
        watcher.start()
    try:
        return fn()
    finally:
        stop_event.set()
        if watcher:
            watcher.join(timeout=2)


def run_one_file(ps_app, file_path, box_list, do_resize, resize_w, resize_dpi):
    """Resize (if requested) and place text -- a single Photoshop open/close for this file."""
    script = JSX_TEMPLATE % {
        "FILE_PATH": js_string(file_path.replace("\\", "/")),
        "DO_RESIZE": "true" if do_resize else "false",
        "RESIZE_W": resize_w if resize_w else 0,
        "RESIZE_DPI": resize_dpi if resize_dpi else 72,
        "BOX_LIST": json.dumps(box_list),
    }
    raw = _run_with_dialog_watchdog(ps_app, lambda: ps_app.DoJavaScript(script))
    return parse_result(raw)


# ______________________________________________________________________
#  WEB EXTRACTION -- runs your existing browser console scripts
#  automatically via Playwright (real Chrome), instead of you pasting
#  them into DevTools by hand. Captures the Blob download they already
#  produce and saves it as textboxes.json. The scripts themselves are
#  embedded verbatim/unmodified below.
# ______________________________________________________________________

SCRIPT_MOI_JS = r"""/**
 * STEP 1 — Paste in browser DevTools Console (F12 → Console tab).
 * Works on: tooning.studio.ailosy.com translate viewer
 *
 * v8-ailosy — ROOT CAUSE FIX:
 *
 *  THE BUG (v6/v7): renderScale was computed from the .losy-editor-viewer-layer-image
 *  CONTAINER width, which is wider than the actual <img> elements. This made
 *  nativeY coordinates too small, so boxes near the top of page N got wrongly
 *  assigned to page N-1, and yRatio came out > 1.0, placing text off the bottom
 *  of the Photoshop document.
 *
 *  THE FIX (v8): renderScale is now derived from EACH PAGE'S OWN <img> element
 *  using getBoundingClientRect().height / nativeHeight.  This is correct
 *  regardless of viewport size, zoom level, or container padding.
 *
 *  Additionally: the cumulative CSS Y of each page is computed from the img
 *  element's actual top offset (getBoundingClientRect().top - firstImgTop)
 *  rather than from accumulated calculated heights.  This removes all
 *  rounding drift and matches exactly where the browser places the images.
 *
 *  Output JSON carries BOTH:
 *   • xRatio / yRatio  — fraction of THIS PAGE's native size  (AUTO / SINGLE)
 *   • xRatioGlobal / yRatioGlobal — fraction of total stacked height (COMBINED)
 */

(function extractTextBoxes() {

  // ── 1. Find the VISIBLE box-layer wrapper ────────────────────────────────
  const allWrappers = Array.from(document.querySelectorAll('.box-layer__box-wrapper'));
  let container = allWrappers.find(el => {
    const s = el.getAttribute('style') || '';
    return !s.includes('display: none') && !s.includes('display:none');
  }) || allWrappers.find(el => el.querySelector('.editor-box p.paragraph'));

  if (!container) {
    console.error('❌ Could not find the visible text layer.');
    return;
  }
  console.log('✅ Visible layer found.');

  // ── 2. Build page table from actual <img> elements ───────────────────────
  //
  //  We iterate .losy-editor-viewer-layer-image__panel in DOM order.
  //  For each panel:
  //    • nativeW/nativeH  → from aspect-ratio style OR title attribute
  //    • renderScale      → img.getBoundingClientRect().height / nativeH
  //    • cssY0            → img.getBoundingClientRect().top − firstImgTop
  //      (scroll-independent: we subtract the first image's top so Y0 of
  //       page 1 = 0 regardless of scroll position)
  //
  //  This means the cumulative Y values are MEASURED from the DOM, not
  //  calculated, so they match perfectly where the browser placed the images.

  const rawPanels = Array.from(
    document.querySelectorAll('.losy-editor-viewer-layer-image__panel')
  );

  if (rawPanels.length === 0) {
    console.error('❌ No .losy-editor-viewer-layer-image__panel elements found.');
    return;
  }

  // ── 2a. De-duplicate panels ───────────────────────────────────────────
  //
  //  Some viewers keep previously-rendered panels mounted in the DOM
  //  (virtualized scroll, reflow on resize/zoom, etc). If the same image
  //  panel appears more than once, querySelectorAll returns it twice and
  //  every downstream calculation (total height, page count, box
  //  assignment) silently doubles up. We key on filename (from the title
  //  attribute) since that's the one thing that's identical for true
  //  duplicates and different for genuinely different pages.
  //
  //  Important: this can't just dedupe by DOM node reference, because a
  //  genuine re-render can create a *new* node for the *same* image —
  //  filename is the actual identity we care about.

  const seenFilenames = new Map(); // filename -> first panel kept
  const panels = [];
  let duplicatesRemoved = 0;

  rawPanels.forEach((panel) => {
    const titleEl = panel.hasAttribute('title') ? panel : panel.closest('[title]');
    const titleStr = titleEl ? titleEl.getAttribute('title') : '';
    const fname = titleStr.split(/\r?\n/)[0] || panel.getAttribute('data-image-id') || null;

    // Only dedupe when we have a real identifying filename. Panels with no
    // filename/id at all are kept as-is (can't safely judge them duplicates).
    if (fname) {
      if (seenFilenames.has(fname)) {
        duplicatesRemoved++;
        return; // skip — already have this exact page
      }
      seenFilenames.set(fname, panel);
    }
    panels.push(panel);
  });

  if (duplicatesRemoved > 0) {
    console.warn(
      `⚠️  Removed ${duplicatesRemoved} duplicate panel(s) with repeated filenames ` +
      `(found ${rawPanels.length} panels in DOM, kept ${panels.length} unique pages). ` +
      `This usually means the page re-rendered or kept old content mounted — ` +
      `if this number looks unexpectedly large, reload the page and re-run the script.`
    );
  }

  const pages = [];
  let firstImgTop = null;
  let lastGoodScale = null;

  panels.forEach((panel, idx) => {
    const img = panel.querySelector('img');

    // ── Native dimensions ──────────────────────────────────────────────
    let nativeW = 690, nativeH = 1000;

    // Method A: aspect-ratio inline style (most reliable)
    if (img) {
      const arStyle = img.style.aspectRatio || '';
      const arMatch = arStyle.match(/(\d+)\s*\/\s*(\d+)/);
      if (arMatch) {
        nativeW = parseInt(arMatch[1], 10);
        nativeH = parseInt(arMatch[2], 10);
      }
    }

    // Method B: title attribute fallback  "IMAGE_001.jpg\n690x3000\n96dpi"
    if (nativeW === 690 && nativeH === 1000) {
      const titleEl = panel.hasAttribute('title') ? panel
        : panel.closest('[title]');
      if (titleEl) {
        const m = titleEl.getAttribute('title').match(/(\d+)x(\d+)/);
        if (m) { nativeW = parseInt(m[1], 10); nativeH = parseInt(m[2], 10); }
      }
    }

    // ── Rendered dimensions from DOM ───────────────────────────────────
    // Use the img element's bounding rect to get the EXACT rendered height.
    // This is the only value that correctly accounts for viewport width,
    // device pixel ratio, and any CSS applied to the image.
    //
    // IMPORTANT (broken/unrendered page fallback): if the image's CSS box
    // never got a real size — most commonly because this <img> never
    // mounted with layout at all — its bounding rect collapses to 0, which
    // would make cssY0..cssY1 collapse to a single point. Every box that
    // should belong to that page then fails the `cssY0 <= cssY < cssY1`
    // test and gets pulled onto a *different* page entirely -- exactly the
    // "boxes land outside the image" symptom. We detect that and estimate
    // this page's height from the last successfully-measured page's scale
    // instead, chaining its top from the previous page's bottom so the
    // stacking order/height stays intact even though this page's own
    // measurement is a guess.
    //
    // NOTE: this does NOT require the image bytes to have actually
    // finished downloading (img.complete/naturalWidth). Sites that set an
    // explicit `aspect-ratio` (as this one does) or fixed width/height in
    // CSS get correctly laid out as soon as CSS applies, regardless of
    // whether the pixel data has loaded yet -- requiring naturalWidth here
    // would wrongly mark perfectly-positioned-but-still-downloading pages
    // as broken (and needlessly wait on network fetches that don't matter
    // for position extraction).
    let renderScale, cssY0, cssY1, estimated = false;

    const rect = img ? img.getBoundingClientRect() : null;
    const loadedOk = !!(rect && rect.height > 0);

    if (loadedOk) {
      const renderedH = rect.height;
      renderScale = renderedH / nativeH;
      lastGoodScale = renderScale;

      // Absolute top of this image relative to the first image
      if (firstImgTop === null) firstImgTop = rect.top;
      cssY0 = rect.top - firstImgTop;
      cssY1 = cssY0 + renderedH;
    } else {
      estimated = true;
      renderScale = lastGoodScale !== null ? lastGoodScale : (320 / nativeW);
      cssY0 = pages.length > 0
        ? pages[pages.length - 1].cssY1
        : 0;
      if (firstImgTop === null) firstImgTop = 0;
      cssY1 = cssY0 + nativeH * renderScale;
    }

    // ── Filename ───────────────────────────────────────────────────────
    const titleEl = panel.hasAttribute('title') ? panel : panel.closest('[title]');
    const titleStr = titleEl ? titleEl.getAttribute('title') : '';
    // title may use \r\n (Windows) or \n; strip \r
    const filename = titleStr.split(/\r?\n/)[0]
      || panel.getAttribute('data-image-id')
      || `page_${idx + 1}`;

    if (estimated) {
      console.warn(`  ⚠️  Page ${idx + 1} (${filename}): image did not load/render — ` +
        `using an ESTIMATED position. Boxes on this page may be placed slightly off. ` +
        `This usually means the page wasn't fully scrolled through before it was saved, ` +
        `or the image resource didn't survive the save/unpack.`);
    }

    pages.push({
      idx, filename, nativeW, nativeH, renderScale,
      cssY0, cssY1, cssH: cssY1 - cssY0,
      nativeY0: 0, // filled in below
      estimated,
    });
  });

  // ── Assign cumulative nativeY0 for global coords ─────────────────────────
  let cumNativeY = 0;
  pages.forEach(p => { p.nativeY0 = cumNativeY; cumNativeY += p.nativeH; });
  const totalNativeH = cumNativeY;
  const totalNativeW = pages[0].nativeW;

  console.log(`\n🖼  ${pages.length} pages (total native: ${totalNativeW}×${totalNativeH}):`);
  pages.forEach(p =>
    console.log(
      `  [${p.idx + 1}] ${p.filename} | ${p.nativeW}×${p.nativeH}` +
      ` | renderScale=${p.renderScale.toFixed(4)}` +
      ` | cssY ${p.cssY0.toFixed(1)}–${p.cssY1.toFixed(1)}` +
      (p.estimated ? '  ⚠️ ESTIMATED (image did not load)' : '')
    )
  );
  {
    const brokenPages = pages.filter(p => p.estimated);
    if (brokenPages.length > 0) {
      console.error(
        `⚠️⚠️⚠️  ${brokenPages.length} of ${pages.length} page image(s) never loaded/rendered ` +
        `and used ESTIMATED positions: ${brokenPages.map(p => p.filename).join(', ')}. ` +
        `Boxes on these pages may be placed slightly off (or worse if several broken pages ` +
        `are adjacent). This almost always means the file was saved before the whole chapter ` +
        `was scrolled through (lazy-loaded images never got fetched), or an image resource ` +
        `didn't survive the save/unpack. Re-save after scrolling to the very bottom first, ` +
        `or extract straight from the live site instead.`
      );
    }
  }

  // ── 3. Extract boxes ──────────────────────────────────────────────────────
  //
  //  IMPORTANT: boxes marked "Proof" in the Translation Box menu are rendered
  //  with a DIFFERENT class — `.editor-box-review` — instead of `.editor-box`.
  //  These boxes typically have background-color: rgb(244, 78, 59) (#f44e3b),
  //  while normal/approved boxes use `.editor-box` (e.g. yellow rgb(252,220,0)
  //  or green rgb(164,221,0)). Selecting only `.editor-box` silently skips
  //  every "Proof" box. We select both classes here so no box — regardless
  //  of its color or review status — gets left out.
  const boxEls = Array.from(
    container.querySelectorAll('.editor-box, .editor-box-review')
  );
  console.log(`\n🔍 ${boxEls.length} boxes in visible layer (including "Proof" boxes).`);

  const pageResults = pages.map(p => ({
    pageIndex:   p.idx,
    filename:    p.filename,
    nativeWidth: p.nativeW,
    nativeHeight: p.nativeH,
    nativeY0:    p.nativeY0,
    textBoxes:   [],
  }));

  let assigned = 0, skipped = 0;

  boxEls.forEach((el, i) => {
    // ── Text: each <p> → one line ──────────────────────────────────────
    // NOTE: only the FIRST <p> in a box carries class="paragraph" -- every
    // subsequent <p> (whether it's a genuinely separate authored line or
    // not) is unclassed. So the class is NOT a reliable "new line" marker;
    // it's just a styling hook on the first line. Selecting only
    // p.paragraph silently dropped every line after the first. Select ALL
    // <p> children in order and keep each as its own line.
    const lines = Array.from(el.querySelectorAll('p'))
      .map(p => (p.innerText || '').trim())
      .filter(t => t.length > 0);
    const text = lines.join('\n');
    if (!text) { skipped++; return; }

    // ── Mark (e.g. "#B93") ──────────────────────────────────────────────
    // Lives in a hidden label div that is a descendant of this same
    // .editor-box / .editor-box-review element. IMPORTANT: normal boxes
    // use class "editor-box__label", but boxes that were fixed/edited and
    // turned into a "Proof" review box use a DIFFERENT class,
    // "editor-box-review__label" — and it's THIS one that carries the
    // box's current/new number (e.g. "#B130"), not the old one (e.g. "#B18").
    // Querying only .editor-box__label silently misses every Proof box's
    // mark. We check both, in DOM order, and take whichever exists.
    // Raw text is like " #B93" — strip whitespace and the leading "#".
    const labelEl  = el.querySelector('.editor-box__label, .editor-box-review__label');
    const markRaw  = labelEl ? labelEl.textContent.trim() : '';
    const mark     = markRaw.replace(/^#/, '');           // "B93" or ""
    if (!mark) {
      console.warn(`  ⚠️  [${i}] No mark (#B...) found for box — layer will be unlabeled.`);
    }
    // Ready-to-use Photoshop layer name: "B93_first 40 chars of text"
    const layerNameSafe = text.replace(/\n/g, ' ').slice(0, 40).trim();
    const layerName = mark ? `${mark}_${layerNameSafe}` : layerNameSafe;

    // ── Position from inline style transform ──────────────────────────
    const styleStr = el.getAttribute('style') || '';
    const tMatch = styleStr.match(/translate\(([0-9.]+)px,\s*([0-9.]+)px\)/);
    if (!tMatch) { skipped++; return; }

    const cssX = parseFloat(tMatch[1]);
    const cssY = parseFloat(tMatch[2]);  // Y in total stacked CSS canvas
    const wM   = styleStr.match(/width:\s*([0-9.]+)px/);
    const hM   = styleStr.match(/height:\s*([0-9.]+)px/);
    const cssW = wM ? parseFloat(wM[1]) : 0;
    const cssH = hM ? parseFloat(hM[1]) : 0;

    // ── Page assignment by cssY ───────────────────────────────────────
    // Compare box Y to each page's cssY0..cssY1 range (measured from DOM).
    let pi = pages.findIndex(p => p.cssY0 <= cssY && cssY < p.cssY1);
    if (pi === -1) {
      // Fallback: nearest page center
      pi = pages.reduce((best, p, j) =>
        Math.abs(cssY - (p.cssY0 + p.cssH / 2)) <
        Math.abs(cssY - (pages[best].cssY0 + pages[best].cssH / 2)) ? j : best, 0);
    }

    const page = pages[pi];
    const rs   = page.renderScale;

    // ── Convert to native coordinates ────────────────────────────────
    // Per-page native coords (Y relative to this page's top)
    const xN  = cssX              / rs;
    const yN  = (cssY - page.cssY0) / rs;  // subtract this page's CSS Y offset
    const wN  = cssW              / rs;
    const hN  = cssH              / rs;

    // Global native coords (Y relative to total canvas top)
    const yNg = page.nativeY0 + yN;

    // ── Font size ─────────────────────────────────────────────────────
    const wrapper  = el.querySelector('[data-content-wrapper]');
    const fsCss    = wrapper ? parseFloat(getComputedStyle(wrapper).fontSize) || 10 : 10;
    const fsNative = fsCss / rs;

    // ── Color and alignment ───────────────────────────────────────────
    const tiptap    = el.querySelector('.losy-tiptap-editor__content');
    const tStyle    = tiptap ? getComputedStyle(tiptap) : null;
    const color     = tStyle ? tStyle.color : 'rgb(0,0,0)';
    const textAlign = wrapper ? (getComputedStyle(wrapper).textAlign || 'center') : 'center';

    // ── Sanity check: warn if yRatio is out of expected [0,1] range ───
    const yRatio = yN / page.nativeH;
    if (yRatio < -0.05 || yRatio > 1.05) {
      console.warn(
        `  ⚠️  [${i}] yRatio=${yRatio.toFixed(3)} out of range!` +
        ` cssY=${cssY.toFixed(1)} page.cssY0=${page.cssY0.toFixed(1)}` +
        ` page.cssY1=${page.cssY1.toFixed(1)} → check renderScale`
      );
    }

    console.log(
      `  ✅ [${i}] ${mark ? '#' + mark : '(no mark)'} → Page ${pi + 1}` +
      ` | page-ratio (${(xN/page.nativeW).toFixed(3)},${yRatio.toFixed(3)})` +
      ` | global-Y ${yNg.toFixed(0)}/${totalNativeH}` +
      ` | "${text.slice(0, 30).replace(/\n/g, '↵')}"`
    );

    // ── Background color & box type (for diagnostics) ─────────────────
    const bgMatch  = styleStr.match(/background-color:\s*(rgb\([^)]+\))/);
    const bgColor  = bgMatch ? bgMatch[1] : 'unknown';
    const isReviewBox = el.classList.contains('editor-box-review');

    pageResults[pi].textBoxes.push({
      id: i, text,
      mark,               // e.g. "B93" (empty string if not found)
      layerName,          // e.g. "B93_craaac" — ready to use as PS layer name

      // Per-page ratios (for AUTO / SINGLE Photoshop mode)
      xRatio: +(xN  / page.nativeW).toFixed(6),
      yRatio: +(yN  / page.nativeH).toFixed(6),
      wRatio: +(wN  / page.nativeW).toFixed(6),
      hRatio: +(hN  / page.nativeH).toFixed(6),

      // Global ratios (for COMBINED Photoshop mode — one tall merged doc)
      xRatioGlobal: +(xN   / totalNativeW).toFixed(6),
      yRatioGlobal: +(yNg  / totalNativeH).toFixed(6),
      wRatioGlobal: +(wN   / totalNativeW).toFixed(6),
      hRatioGlobal: +(hN   / totalNativeH).toFixed(6),

      fontSizeNativePx:    +fsNative.toFixed(2),
      fontSizeRatio:       +(fsNative / page.nativeH).toFixed(6),   // per-page
      fontSizeRatioGlobal: +(fsNative / totalNativeH).toFixed(6),   // global

      color, textAlign,
      opacity: parseFloat(getComputedStyle(el).opacity) || 1,
      bgColor, isReviewBox,
    });
    assigned++;
  });

  console.log(`\n📊 ${assigned} assigned, ${skipped} skipped.`);
  pageResults.forEach(p =>
    console.log(`  Page ${p.pageIndex + 1} (${p.filename}): ${p.textBoxes.length} boxes`));

  const unmarkedCount = pageResults.reduce(
    (n, p) => n + p.textBoxes.filter(b => !b.mark).length, 0);
  if (unmarkedCount > 0) {
    console.warn(`\n⚠️  ${unmarkedCount} box(es) have no #B mark — check .editor-box__label exists for them.`);
  }

  // ── Breakdown by background color — confirms "Proof" boxes weren't missed ─
  const colorCounts = {};
  pageResults.forEach(p => p.textBoxes.forEach(b => {
    colorCounts[b.bgColor] = (colorCounts[b.bgColor] || 0) + 1;
  }));
  console.log('\n🎨 Boxes by background color:');
  Object.entries(colorCounts).forEach(([c, n]) =>
    console.log(`   ${c}: ${n} box(es)`));

  if (assigned === 0) { console.error('❌ No boxes extracted.'); return; }

  // ── 3a. Final sanity check — catch repeated-content patterns ─────────────
  //
  //  Even after dedup-by-filename above, it's worth double-checking the
  //  FINAL page list for a pattern that strongly suggests duplicated
  //  content slipped through (e.g. two different DOM nodes that happened
  //  to carry the same filename were treated as distinct because some
  //  other identifying attribute differed). The clearest signal: the
  //  list of filenames repeats itself exactly, split in half.
  const filenameList = pages.map(p => p.filename);
  let repeatWarning = null;
  if (filenameList.length % 2 === 0) {
    const half = filenameList.length / 2;
    const firstHalf  = filenameList.slice(0, half).join('|');
    const secondHalf = filenameList.slice(half).join('|');
    if (firstHalf === secondHalf) {
      repeatWarning =
        `⚠️⚠️⚠️  The ${filenameList.length} pages look like the SAME ${half} filenames, ` +
        `repeated twice in a row. This almost always means the page content was captured ` +
        `twice (e.g. duplicate DOM panels, or the script ran while the viewer was still ` +
        `re-rendering). The downloaded file may have DOUBLE the real canvas height, which ` +
        `will throw off every position calculation in Photoshop.\n` +
        `   → Reload the page, wait for it to fully finish loading, then re-run this script.`;
    }
  }

  if (repeatWarning) {
    console.error(repeatWarning);
    if (!confirm(
      'WARNING: this export looks like it contains duplicate pages ' +
      '(same filenames repeated twice). This will likely cause incorrect ' +
      'text placement in Photoshop.\n\nDownload anyway?'
    )) {
      console.log('Cancelled — reload the page and re-run the script.');
      return;
    }
  }

  // ── 4. Download — ASCII-safe JSON ─────────────────────────────────────────
  function toAsciiSafeJSON(obj) {
    return JSON.stringify(obj, null, 2).replace(/[\u0080-\uFFFF]/g, c =>
      '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0')
    );
  }

  const output = {
    extractedAt:      new Date().toISOString(),
    pageURL:          window.location.href,
    scriptVersion:    'v13-ailosy-broken-img-fallback',
    coordSystem:      'ratios — per-page (xRatio/yRatio) AND global (xRatioGlobal/yRatioGlobal)',
    totalPages:       pages.length,
    totalTextBoxes:   assigned,
    totalNativeWidth: totalNativeW,
    totalNativeHeight: totalNativeH,
    duplicatePanelsRemoved: duplicatesRemoved,
    pages: pageResults,
  };

  const blob = new Blob([toAsciiSafeJSON(output)], { type: 'application/json' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url; a.download = 'textboxes.json';
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);

  console.log(`\n🎉 Downloaded textboxes.json — ${pages.length} page(s), ${assigned} box(es).`);
  if (duplicatesRemoved > 0) {
    console.log(`   (${duplicatesRemoved} duplicate panel(s) were auto-removed before export.)`);
  }
  console.log('   If any ⚠️ warnings appeared above, check that all page images are');
  console.log('   fully loaded and visible before running this script.');

})();"""  # tooning.studio.ailosy.com (per-page box-layer layout)


def get_extraction_profile_dir():
    """
    Directory for a dedicated, persistent Chrome profile used only for
    extraction. Kept separate from your everyday Chrome profile so this
    tool never touches your regular browsing session/tabs -- but it DOES
    persist across runs, so once you log into the target site here once,
    that login is remembered for next time (until the site's session
    naturally expires).
    """
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA") or str(Path.home())
    profile_dir = os.path.join(base, "PSD-Batch-Placer", "chrome-profile")
    os.makedirs(profile_dir, exist_ok=True)
    return profile_dir


def unpack_mhtml(mhtml_path, log_cb=None):
    """
    Unpacks a saved .mhtml/.mht file (a MIME multipart/related message, the
    format behind Chrome's "Webpage, Single File" save) into a plain,
    self-contained .html file plus its resource files (images, CSS, fonts,
    etc.), written to a fresh temp folder -- then returns the path to that
    extracted .html file.

    Why this is needed: Chrome opens .mhtml/.mht files directly through a
    sandboxed viewer that blocks ALL script execution, which breaks the
    getBoundingClientRect()/getComputedStyle()-based extraction script. A
    plain, re-assembled .html file with everything saved alongside it and
    re-linked opens normally with scripts fully enabled and renders the
    same page.

    How it works: an .mhtml file is a MIME message -- one part is the page's
    HTML, and the rest are its resources, each tagged with a Content-Location
    (the resource's original URL) and/or a Content-ID. This walks every part,
    saves each resource to disk under a safe local filename, and rewrites
    every reference to a resource's Content-Location URL or "cid:<id>" inside
    the HTML text to point at that local file instead.

    Returns the absolute path to the extracted index.html file.
    Raises RuntimeError with a descriptive message if the file can't be
    parsed as MHTML or has no HTML part.
    """
    def log(msg):
        if log_cb:
            log_cb(msg)

    try:
        with open(mhtml_path, "rb") as f:
            raw = f.read()
        msg = email.message_from_bytes(raw)
    except Exception as e:
        raise RuntimeError("Could not parse '%s' as MHTML: %s" % (os.path.basename(mhtml_path), e))

    out_dir = tempfile.mkdtemp(prefix="mhtml_unpack_")

    if not msg.is_multipart():
        # Rare: a degenerate single-part save that's really just HTML.
        payload = msg.get_payload(decode=True)
        if not payload:
            raise RuntimeError("This .mhtml/.mht file has no readable HTML content.")
        out_path = os.path.join(out_dir, "index.html")
        with open(out_path, "wb") as f:
            f.write(payload)
        log("MHTML had a single part; saved directly as HTML (no resources to unpack).")
        return out_path

    log("Unpacking MHTML to a plain HTML file in: " + out_dir)

    html_part = None
    resources = []

    for part in msg.walk():
        if part.is_multipart():
            continue
        ctype = (part.get_content_type() or "").lower()
        payload = part.get_payload(decode=True)
        if payload is None:
            continue

        location = part.get("Content-Location")
        cid = part.get("Content-ID")
        if cid:
            cid = cid.strip().lstrip("<").rstrip(">")

        if html_part is None and ctype == "text/html":
            html_part = {
                "payload": payload,
                "charset": part.get_content_charset() or "utf-8",
            }
        else:
            resources.append({"location": location, "cid": cid, "ctype": ctype, "payload": payload})

    if html_part is None:
        raise RuntimeError(
            "Could not find an HTML part inside '%s' -- it may be corrupted "
            "or not actually a saved MHTML page." % os.path.basename(mhtml_path)
        )

    # Save every resource part to disk under a safe, unique filename, and
    # remember every string (its Content-Location URL and/or "cid:<id>")
    # that the HTML might reference it by.
    url_to_local = {}
    used_names = {"index.html"}

    for i, res in enumerate(resources):
        ext = mimetypes.guess_extension(res["ctype"].split(";")[0].strip()) or ""

        base_name = None
        if res["location"]:
            tail = res["location"].split("?")[0].split("#")[0].rstrip("/").rsplit("/", 1)[-1]
            tail = "".join(c for c in tail if c not in '<>:"/\\|?*')
            if tail:
                base_name = tail
        if not base_name:
            base_name = "resource_%03d%s" % (i, ext)
        elif not os.path.splitext(base_name)[1] and ext:
            base_name = base_name + ext

        name = base_name
        n = 1
        while name in used_names:
            stem, e = os.path.splitext(base_name)
            name = "%s_%d%s" % (stem, n, e)
            n += 1
        used_names.add(name)

        try:
            with open(os.path.join(out_dir, name), "wb") as f:
                f.write(res["payload"])
        except OSError:
            name = "resource_%03d%s" % (i, ext)
            with open(os.path.join(out_dir, name), "wb") as f:
                f.write(res["payload"])
            used_names.add(name)

        if res["location"]:
            url_to_local[res["location"]] = name
        if res["cid"]:
            url_to_local["cid:" + res["cid"]] = name

    # Rewrite every reference inside the HTML text to point at the locally
    # saved file instead of the original (often now-unreachable) URL. Longer
    # URLs are replaced first so a short URL that happens to be a substring
    # of a longer one can't clobber the longer match first.
    html_text = html_part["payload"].decode(html_part["charset"], errors="replace")
    for url in sorted(url_to_local, key=len, reverse=True):
        html_text = html_text.replace(url, url_to_local[url])

    # Strip ALL <link rel="modulepreload"/"preload"/"prefetch"> hints,
    # unconditionally -- whether their href points at an unresolved live
    # https:// URL, or at a resource that WAS embedded and rewritten to a
    # local file path.
    #
    # These are pure performance hints (never required for rendering -- the
    # browser only actually *uses* the resource when the real <link
    # rel="stylesheet">/<script> tag for it loads). Two different failure
    # modes show up depending on the href:
    #   - unresolved https:// href -> blocked by CORS from a "null" file://
    #     origin (no Access-Control-Allow-Origin header)
    #   - rewritten LOCAL href -> ALSO blocked, because rel="preload"/
    #     "prefetch" always issue a CORS-mode fetch, and Chromium refuses
    #     that under file:// origin even for a file in the very same folder
    #     ("Cross origin requests are only supported for protocol schemes:
    #     ... http, https, ..." -- file:// isn't in that list)
    # Either way the fetch can never succeed locally and the hint does
    # nothing useful once removed -- it just floods the console log with
    # scary-looking but harmless [console] error lines. Since there are no
    # <script> tags left in this saved snapshot anyway (Chrome doesn't
    # serialize those either), nothing actually depends on these hints.
    #
    # ALSO strip any leftover crossorigin="" / crossorigin="anonymous" /
    # crossorigin="use-credentials" attribute from whatever <link> tags
    # remain (the real rel="stylesheet" tags this site actually needs).
    # That attribute is what forces the browser into CORS-mode fetching in
    # the first place -- fine on the live https:// site, but once the href
    # has been rewritten to a plain local filename, CORS mode has nothing
    # to do except fail, because file:// is never allowed as a CORS request
    # origin (same "Cross origin requests are only supported for protocol
    # schemes: ... http, https ..." error, this time on the actual
    # stylesheets rather than the hints). A same-folder local file load
    # never needs crossorigin at all, so dropping the attribute is safe and
    # lets these stylesheets load normally instead of failing outright.
    stripped_count = 0
    crossorigin_stripped = 0

    def _strip_link_tag(match):
        nonlocal stripped_count, crossorigin_stripped
        tag = match.group(0)
        rel_m = re.search(r'\brel=["\']([a-zA-Z-]+)["\']', tag)
        if rel_m and rel_m.group(1).lower() in ("modulepreload", "preload", "prefetch"):
            stripped_count += 1
            return ""
        new_tag, n = re.subn(r'\s+crossorigin(=["\'][^"\']*["\'])?', '', tag, flags=re.I)
        if n:
            crossorigin_stripped += 1
        return new_tag

    html_text = re.sub(r'<link\b[^>]*>', _strip_link_tag, html_text)

    if stripped_count:
        log("Removed %d modulepreload/preload/prefetch hint(s) -- these are "
            "performance-only hints that always fail under file:// (CORS), "
            "whether or not the resource they point to was embedded. Removing "
            "them avoids harmless-but-noisy CORS console errors; the actual "
            "<link rel=\"stylesheet\">/content that matters is untouched." % stripped_count)
    if crossorigin_stripped:
        log("Removed crossorigin=\"\" from %d <link> tag(s) -- otherwise even "
            "the real stylesheets fail to load locally, since crossorigin "
            "forces a CORS-mode fetch that file:// pages can never satisfy, "
            "even for a file in the same folder." % crossorigin_stripped)

    out_path = os.path.join(out_dir, "index.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html_text)

    log("Unpacked %d resource file(s)." % len(resources))
    return out_path


def extract_textboxes(source, is_url, output_json_path, log_cb=None, headless=False):
    """
    Runs the site's extraction console script against either:
      - a live website (is_url=True, `source` is the URL), or
      - a locally saved HTML/MHTML file (is_url=False, `source` is a file path)

    Either way this opens the page in a dedicated, persistent Chrome/Edge
    profile via Playwright (tries Chrome, then Edge -- driven directly, no
    extra browser download needed). "Persistent" means: the FIRST time you
    use a live URL, you'll need to log in manually in the window that opens
    (you get up to 2 minutes to do so) -- after that, the session is saved
    in this dedicated profile and reused automatically on later runs, no
    re-login needed until the site's session actually expires. This is a
    separate profile from your everyday Chrome, so it never touches your
    regular tabs/history/logins.

    Auto-detects which of the two site layouts it is, runs the matching
    (unmodified) console script, captures the textboxes.json download it
    triggers, and saves it to output_json_path.

    For the local-file case: no login/network is strictly needed if the
    page was already fully rendered when saved -- this just re-runs the
    extraction script against that already-rendered DOM. A real browser is
    still required here (not just an HTML parser) because the extraction
    script reads live layout via getBoundingClientRect() / getComputedStyle(),
    which only a rendering engine can compute -- static parsing of the
    file's markup can't reproduce those numbers.

    Deliberately does NOT fall back to Playwright's own bundled Chromium --
    that would require the extra `playwright install chromium` download.
    If neither Chrome nor Edge can be found, this raises a clear error.

    Raises RuntimeError with a descriptive message on any failure.
    """
    if sync_playwright is None:
        raise RuntimeError(
            "Web extraction needs the 'playwright' package.\n"
            "Install it with:\n    pip install playwright\n"
            "(no browser download needed -- this uses your existing installed Chrome/Edge)"
        )

    def log(msg):
        if log_cb:
            log_cb(msg)

    def _scroll_through(page):
        """Scrolls all the way down and back up once, pausing briefly at
        each step, so anything that mounts/loads lazily as it enters the
        viewport gets a chance to do so before extraction reads the DOM."""
        page.evaluate("""
            async () => {
                function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
                const step = Math.max(500, Math.floor(window.innerHeight * 0.85));
                for (let i = 0; i < 5000; i++) {
                    const before = window.scrollY;
                    window.scrollBy(0, step);
                    await sleep(180);
                    const atBottom = (window.scrollY + window.innerHeight) >=
                                     (document.documentElement.scrollHeight - 2);
                    if (atBottom || window.scrollY === before) break;
                }
                await sleep(400);
                window.scrollTo(0, 0);
                await sleep(200);
            }
        """, timeout=180000)

    def _wait_for_images(page, max_wait_ms):
        """
        Actively polls (instead of a single fixed-length sleep) until every
        currently-known page image finishes loading, or max_wait_ms elapses.
        A single check right after scrolling is exactly what causes
        intermittent misalignment on a live site: an image can genuinely
        still be mid-download at that instant, and checking only once
        catches it "not loaded" even though it finishes a moment later --
        which is why re-running the same job a second time often "fixes"
        it (everything's cached by then). Polling instead of sleeping once
        means we only wait as long as actually needed, and no longer than
        max_wait_ms.
        Returns the list of image labels still not loaded when it gives up.
        """
        return page.evaluate("""
            async () => {
                function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
                const imgs = Array.from(document.querySelectorAll(
                    '.losy-editor-viewer-layer-image__panel img, img[title]'
                ));
                const isPending = img => !(img.complete && img.naturalWidth > 0);
                const deadline = Date.now() + %d;
                let pending = imgs.filter(isPending);
                while (pending.length && Date.now() < deadline) {
                    await sleep(300);
                    pending = imgs.filter(isPending);
                }
                return pending.map(img =>
                    (img.getAttribute('title') || img.src || '(no title)').split(/\\r?\\n/)[0]
                );
            }
        """ % max_wait_ms, timeout=max_wait_ms + 10000)

    if is_url:
        nav_target = source
    else:
        local_path = source
        if os.path.splitext(source)[1].lower() in (".mhtml", ".mht"):
            log("Detected .mhtml/.mht file -- unpacking to plain HTML first "
                "(avoids Chrome's mhtml script-sandbox) ...")
            local_path = unpack_mhtml(source, log_cb=log)
        nav_target = Path(local_path).resolve().as_uri()

    profile_dir = get_extraction_profile_dir()

    with sync_playwright() as p:
        context = None
        launch_errors = []
        for channel in ("chrome", "msedge"):
            try:
                context = p.chromium.launch_persistent_context(
                    profile_dir, channel=channel, headless=headless
                )
                log("Using browser channel: %s (persistent profile at %s)" % (channel, profile_dir))
                break
            except Exception as e:
                launch_errors.append("%s: %s" % (channel, e))

        if context is None:
            raise RuntimeError(
                "Could not launch Chrome or Edge via Playwright.\n\n"
                "Make sure Google Chrome or Microsoft Edge is installed on "
                "this machine.\n\nDetails:\n" + "\n".join(launch_errors)
            )

        try:
            page = context.pages[0] if context.pages else context.new_page()

            def handle_dialog(dialog):
                log("  [dialog] (%s) %s" % (dialog.type, dialog.message))
                # Safer default: dismiss rather than accept, so a detected
                # "duplicate pages" warning does NOT silently download a
                # bad file. Alerts (info-only) are just closed either way.
                try:
                    dialog.dismiss()
                except Exception:
                    pass

            page.on("dialog", handle_dialog)

            interesting_markers = ("assigned", "boxes", "Downloaded", "Could not", "duplicate", "WARNING")
            def handle_console(msg):
                text = msg.text
                if msg.type in ("warning", "error") or any(m in text for m in interesting_markers):
                    log("  [console] " + text)

            page.on("console", handle_console)

            log("Opening " + (source if is_url else os.path.basename(source)) + " ...")
            # For a local saved file, waiting for the full "load" event means
            # waiting for EVERY subresource to settle -- including any
            # cross-origin script/CSS references that were never actually
            # captured in the save (common on sites whose JS bundle isn't
            # saved by "Save Page As" at all) and can only fail after a real
            # network round-trip. With a few hundred of those, "load" can
            # take a very long time or appear to hang, even though nothing
            # we need is waiting on it -- the DOM/CSS needed for position
            # extraction is already complete once the document itself has
            # parsed. So for local files we only wait for DOMContentLoaded;
            # the explicit wait_for_selector() call right after still makes
            # sure the actual content we need is present before continuing.
            goto_wait = "load" if is_url else "domcontentloaded"
            page.goto(nav_target, wait_until=goto_wait, timeout=120000)

            if is_url:
                log("Waiting for content (log in now in the opened window if a login "
                    "screen appears -- you have up to 2 minutes)...")
            try:
                # state="attached" (NOT the default "visible") is deliberate here.
                #
                # For a saved local file, there is NO JavaScript running at all
                # (Chrome's MHTML/HTML save bakes in the already-rendered DOM but
                # strips <script> tags entirely) -- and on some sites the box-layer
                # overlay container's own HEIGHT is normally set by client-side JS
                # reacting to the page images' total stacked height. Without that
                # JS ever running, the container can end up with a real, valid
                # computed style (no display:none, opacity:1) but a collapsed 0x0
                # bounding box -- which Playwright's default "visible" wait state
                # treats as not-yet-ready, and then waits the full timeout for a
                # visibility that will never arrive, hanging with no further
                # progress (and no error) until it finally times out.
                #
                # This doesn't affect correctness: the actual extraction script
                # never depends on this container's own size. It locates the
                # container purely by checking its inline style string for the
                # literal text "display: none", and reads each text box's
                # position straight out of that box's own inline
                # style="transform: translate(...)" attribute -- never via
                # getBoundingClientRect() on the container or the boxes. So DOM
                # presence (attached) is exactly what's needed; requiring visual
                # visibility on top of that is stricter than the actual script
                # needs and can hang indefinitely on this exact scenario.
                page.wait_for_selector(
                    ".box-layer__box-wrapper, article.box, .box-layer",
                    state="attached",
                    timeout=120000
                )
            except PlaywrightTimeoutError:
                raise RuntimeError(
                    "Timed out after 2 minutes waiting for expected page content. "
                    + ("If a login screen appeared, log in within the opened browser "
                       "window, then click Run again -- your session will be remembered "
                       "next time." if is_url else
                       "Make sure the page fully loaded/rendered before it was saved.")
                )

            if page.query_selector(".box-layer__box-wrapper"):
                script_text = SCRIPT_MOI_JS
                site_label = "tooning.studio.ailosy.com"
            else:
                raise RuntimeError(
                    "Could not detect the site layout from this file (no known "
                    "selectors found). This tool only supports "
                    "tooning.studio.ailosy.com. Make sure the page fully "
                    "loaded/rendered before it was saved."
                )

            log("Detected site: " + site_label)

            # ── tooning.studio.ailosy.com specific safeguard ──────────────
            #
            # This viewer VIRTUALIZES its page list: only a handful of page
            # panels are ever mounted in the DOM at once, and the rest get
            # unmounted as you scroll past them. When extracting from a
            # LOCAL saved .html/.mhtml file, whatever panels happened to be
            # mounted at the moment the page was saved is *permanently* all
            # this file will ever contain -- there is no live JavaScript in
            # a saved snapshot to mount the missing ones as you scroll
            # (Chrome's "Save Page As" does not bundle the site's JS bundle
            # at all for this site). So if this file only captured a
            # handful of pages, no amount of re-scrolling here will recover
            # the rest -- only extracting straight from the live URL can.
            if not is_url:
                try:
                    panel_count = page.evaluate(
                        "document.querySelectorAll('.losy-editor-viewer-layer-image__panel').length"
                    )
                except Exception:
                    panel_count = None
                if panel_count is not None and panel_count > 0 and panel_count <= 20:
                    log(
                        "  WARNING: this saved file only has %d page-image panel(s) mounted "
                        "in the DOM. This site only keeps a limited window of pages mounted "
                        "at a time (the rest are removed from the DOM once scrolled past), "
                        "and a saved .html/.mhtml file has no working JavaScript left to "
                        "mount the missing ones -- scrolling after reopening this file can't "
                        "recover them. If your chapter has more than %d page(s), this file is "
                        "an incomplete capture. For this site, extracting directly from the "
                        "live URL (not a saved file) is the only way to reliably get every "
                        "page." % (panel_count, panel_count)
                    )

            # This page can be extremely tall (many stacked panels), and
            # sites like this commonly only render/mount a panel's text
            # overlay once it's scrolled into view (virtualization/lazy
            # loading) -- a page that was just loaded (or a saved snapshot
            # that was never scrolled through before saving) can be missing
            # most of that content even though wait_for_selector above found
            # SOME box elements near the top. Scroll all the way through
            # once, pausing briefly at each step, so anything that renders
            # lazily from data already on the page gets a chance to mount
            # before the extraction script runs and reads the DOM.
            log("Scrolling through the full page once to load any lazy-rendered content ...")
            try:
                _scroll_through(page)
                log("  Done scrolling.")
            except Exception as scroll_err:
                log("  (scroll-through step hit an issue, continuing anyway: %s)" % scroll_err)

            # ── Actively wait for any not-yet-loaded page images, then
            #    report which (if any) are still broken. This matters most
            #    for the "saved .html/.mhtml file" case: if an image never
            #    made it into the saved page (e.g. it was lazy-loaded and
            #    the chapter wasn't scrolled through before saving, or the
            #    resource didn't survive the save/unpack), it will NEVER
            #    load here either -- there's no network fallback for a
            #    local file, so there's nothing to gain from waiting long.
            #
            #    For a LIVE site, though, a single fixed sleep right after
            #    scrolling is exactly what caused the intermittent
            #    misalignment this is guarding against: an image can
            #    genuinely still be mid-download at that instant. So for a
            #    live URL, this polls (rather than sleeping once) for up
            #    to 15s, and if anything's still pending, scrolls through
            #    again and gives it a second, shorter poll -- instead of
            #    requiring a manual re-run to "get lucky" on timing.
            first_wait_ms = 15000 if is_url else 3000
            try:
                broken = _wait_for_images(page, first_wait_ms)
            except Exception as img_check_err:
                log("  (image-load check skipped: %s)" % img_check_err)
                broken = []

            if broken and is_url:
                log("  %d image(s) still not loaded after %.0fs -- scrolling through "
                    "again and giving them more time ..." % (len(broken), first_wait_ms / 1000))
                try:
                    _scroll_through(page)
                except Exception:
                    pass
                try:
                    broken = _wait_for_images(page, 10000)
                except Exception as img_check_err:
                    log("  (second image-load check skipped: %s)" % img_check_err)

            if broken:
                log("  WARNING: %d page image(s) failed to load: %s" % (
                    len(broken), ", ".join(broken[:10]) + (" ..." if len(broken) > 10 else "")
                ))
                if not is_url:
                    log("  This usually means the saved file was created before the whole "
                        "chapter was scrolled through (so these images were never fetched "
                        "into the save), or the image resource didn't survive the "
                        "save/unpack. The extraction script will estimate positions for "
                        "these pages, but for fully correct placement, re-save the page "
                        "after scrolling all the way to the bottom first, or extract "
                        "directly from the live site/URL instead.")
                else:
                    log("  These images never fully loaded even after two scroll/wait "
                        "passes -- this run's positions for those pages may still be "
                        "estimated. This is usually a slow/unstable connection; if it "
                        "keeps happening, try running again once, or check your network.")
            else:
                log("  All page images loaded OK.")

            log("Running extraction script ...")

            try:
                with page.expect_download(timeout=30000) as download_info:
                    page.evaluate(script_text)
                download = download_info.value
            except PlaywrightTimeoutError:
                raise RuntimeError(
                    "No textboxes.json download happened within 30s. This usually means "
                    "either: the extraction script hit an error (check the [console] lines "
                    "above), or a warning dialog (e.g. duplicate pages) was auto-dismissed. "
                    "Scroll up in the log for details."
                )

            download.save_as(output_json_path)
            log("Saved: " + output_json_path)

        finally:
            context.close()



def find_web_file(folder):
    """
    Looks for a saved HTML/MHTML file already sitting in the folder.
    Prefers plain .html/.htm over .mhtml/.mht when both exist -- both work
    (an .mhtml/.mht file is automatically unpacked to plain HTML before
    extraction), but .html needs no unpacking step so it's picked first
    when there's a choice. Returns (path_or_None, is_mhtml, multiple_found).
    """
    names = [f for f in os.listdir(folder) if f.lower().endswith(WEB_EXTENSIONS)]
    if not names:
        return None, False, False

    def sort_key(name):
        ext = os.path.splitext(name)[1].lower()
        # .html/.htm sort before .mhtml/.mht
        return (0 if ext in (".html", ".htm") else 1, name.lower())

    names.sort(key=sort_key)
    chosen = names[0]
    is_mhtml = os.path.splitext(chosen)[1].lower() in (".mhtml", ".mht")
    return os.path.join(folder, chosen), is_mhtml, (len(names) > 1)


def scan_folder(folder):
    """
    Returns (json_path_or_None, files_info, multiple_json,
             html_path_or_None, is_mhtml, multiple_html).

    Only PSD/PSB/TIFF files are required. An existing textboxes.json in the
    folder is optional -- if a saved HTML/MHTML file or a website URL is
    given in the GUI, a fresh textboxes.json is extracted and overwrites
    whatever is here.
    """
    json_names = [f for f in os.listdir(folder) if f.lower().endswith(".json")]
    json_path = os.path.join(folder, json_names[0]) if json_names else None

    html_path, is_mhtml, multiple_html = find_web_file(folder)

    image_names = sorted(
        [f for f in os.listdir(folder) if f.lower().endswith(IMAGE_EXTENSIONS)],
        key=lambda s: s.lower()
    )
    if not image_names:
        raise ValueError("No .psd/.psb/.tif files found in that folder.")

    files_info = []
    for name in image_names:
        path = os.path.join(folder, name)
        dims = read_psd_dimensions(path)
        files_info.append({"path": path, "name": name, "width": dims["width"], "height": dims["height"]})

    return json_path, files_info, (len(json_names) > 1), html_path, is_mhtml, multiple_html


# ══════════════════════════════════════════════════════════════════
#  GUI
# ══════════════════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════════════════
#  TAB 1 GUI: BATCH TEXT PLACER + RESIZER
# ══════════════════════════════════════════════════════════════════

class PlacerTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, style="Card.TFrame")

        self.folder = tk.StringVar(value="No folder selected")
        self.do_resize = tk.BooleanVar(value=False)
        self.resize_width = tk.StringVar(value="")
        self.resize_dpi = tk.StringVar(value="")
        self.html_file_path = tk.StringVar(value="")
        self.status_text = tk.StringVar(value="")
        self.show_details = tk.BooleanVar(value=False)

        self._json_path = None
        self._files_info = None
        self._data = None
        self._stop_event = threading.Event()

        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self, style="Card.TFrame", padding=(14, 14, 14, 6))
        top.pack(fill='x')
        ttk.Button(top, text="Choose Folder...", style="Primary.TButton",
                   command=self.choose_folder).pack(side='left')
        ttk.Label(top, textvariable=self.folder, style="CardMuted.TLabel").pack(side='left', padx=10)

        opts = ttk.LabelFrame(self, text="Extraction source", style="Card.TLabelframe", padding=12)
        opts.pack(fill='x', padx=14, pady=(6, 6))

        file_row = ttk.Frame(opts, style="Card.TFrame")
        file_row.pack(fill='x')
        ttk.Label(file_row, text="Saved HTML/MHTML file:", style="Card.TLabel").pack(side='left')
        self.file_path_label = ttk.Label(file_row, textvariable=self.html_file_path, style="CardMuted.TLabel")
        self.file_path_label.pack(side='left', padx=(6, 6))
        self.browse_file_btn = ttk.Button(file_row, text="Browse...", command=self.choose_html_file)
        self.browse_file_btn.pack(side='left')

        ttk.Label(
            opts, style="CardMuted.TLabel",
            text="Leave unset to reuse an existing textboxes.json already in the chosen folder."
        ).pack(fill='x', pady=(4, 0))

        opts2 = ttk.Frame(self, style="Card.TFrame", padding=(14, 0))
        opts2.pack(fill='x')
        ttk.Checkbutton(opts2, text="Resize before placing text (bilinear resample)", variable=self.do_resize,
                         style="Card.TCheckbutton", command=self._toggle_resize_fields).pack(side='left')
        ttk.Label(opts2, text="Width (px):", style="Card.TLabel").pack(side='left', padx=(15, 3))
        self.width_entry = ttk.Entry(opts2, textvariable=self.resize_width, width=8, state='disabled')
        self.width_entry.pack(side='left')
        ttk.Label(opts2, text="DPI:", style="Card.TLabel").pack(side='left', padx=(15, 3))
        self.dpi_entry = ttk.Entry(opts2, textvariable=self.resize_dpi, width=8, state='disabled')
        self.dpi_entry.pack(side='left')

        btns = ttk.Frame(self, style="Card.TFrame", padding=14)
        btns.pack(fill='x')
        self.run_btn = ttk.Button(btns, text="Run", style="Primary.TButton",
                                   command=self.run_batch, state='disabled')
        self.run_btn.pack(side='left')
        self.stop_btn = ttk.Button(btns, text="Stop", command=self.stop_batch, state='disabled')
        self.stop_btn.pack(side='left', padx=(8, 0))
        self.details_btn = ttk.Checkbutton(
            btns, text="Show details", variable=self.show_details,
            style="Card.TCheckbutton", command=self._toggle_details
        )
        self.details_btn.pack(side='left', padx=(15, 0))

        self.summary = ttk.Label(self, text="", style="CardMuted.TLabel", padding=(14, 0))
        self.summary.pack(fill='x')

        self.status_label = ttk.Label(self, textvariable=self.status_text, style="Status.TLabel",
                                       background=COLORS["surface"], padding=(14, 4))
        self.status_label.pack(fill='x')

        self.log_frame = ttk.Frame(self, style="Card.TFrame", padding=14)
        self.log = scrolledtext.ScrolledText(
            self.log_frame, wrap='word', height=20, state='disabled',
            bg=COLORS["bg"], fg=COLORS["text"], relief='flat', borderwidth=0,
            insertbackground=COLORS["text"],
        )
        self.log.pack(fill='both', expand=True)
        # log_frame is intentionally NOT packed here -- it stays hidden until
        # "Show details" is checked, so day-to-day use just shows the one-line
        # status above instead of a big scroll of progress text.

    def _toggle_details(self):
        if self.show_details.get():
            self.log_frame.pack(fill='both', expand=True)
            self.winfo_toplevel().geometry("780x600")
        else:
            self.log_frame.pack_forget()
            self.winfo_toplevel().geometry("780x300")

    def _status(self, text):
        """One-line, always-visible status update (e.g. 'Processing 3/12...')."""
        self.status_text.set(text)

    def _toggle_resize_fields(self):
        state = 'normal' if self.do_resize.get() else 'disabled'
        self.width_entry.config(state=state)
        self.dpi_entry.config(state=state)

    def choose_html_file(self):
        path = filedialog.askopenfilename(
            title="Select saved HTML or MHTML file",
            filetypes=[
                ("HTML files", "*.html *.htm"),
                ("MHTML files (auto-unpacked to HTML)", "*.mhtml *.mht"),
                ("All files", "*.*"),
            ]
        )
        if not path:
            return
        self.html_file_path.set(path)

    def _log(self, text):
        self.log.config(state='normal')
        self.log.insert('end', text + "\n")
        self.log.see('end')
        self.log.config(state='disabled')

    def choose_folder(self):
        folder = filedialog.askdirectory(title="Select folder with PSD files")
        if not folder:
            return
        self.folder.set(folder)

        try:
            json_path, files_info, multiple_json, html_path, is_mhtml, multiple_html = scan_folder(folder)
        except ValueError as e:
            messagebox.showwarning("Nothing to process", str(e))
            self.run_btn.config(state='disabled')
            return

        self._json_path = json_path
        self._files_info = files_info
        self._data = None

        self.log.config(state='normal')
        self.log.delete('1.0', 'end')
        self.log.config(state='disabled')
        for f in files_info:
            self._log("  %-40s %dx%d" % (f["name"], f["width"], f["height"]))

        if html_path:
            self.html_file_path.set(html_path)
            self._log("Found saved web file: %s%s%s" % (
                os.path.basename(html_path),
                "  (multiple found, using this one -- .html preferred over .mhtml)" if multiple_html else "",
                "  (.mhtml -- will be auto-unpacked to plain HTML before extraction)" if is_mhtml else ""
            ))

        if json_path:
            # A JSON already sits in the folder. If a saved file/URL is set,
            # it gets overwritten by a fresh extraction when Run is clicked.
            try:
                with open(json_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if data.get("pages"):
                    self._data = data
                    self._log("Existing JSON found: %s (%d page(s))" % (
                        os.path.basename(json_path), len(data["pages"])
                    ))
                else:
                    self._log("Existing JSON found but has no pages: %s" % os.path.basename(json_path))
            except (json.JSONDecodeError, OSError) as e:
                self._log("Existing JSON found but could not be read: %s" % e)
        elif not html_path:
            self._log("No existing textboxes.json -- pick a saved HTML/MHTML file above before running.")

        self.summary.config(
            text="%d file(s) found.%s" % (
                len(files_info),
                "  JSON: %s" % os.path.basename(json_path) if json_path else "  No JSON yet"
            )
        )

        self.run_btn.config(state='normal')

    def run_batch(self):
        if sys.platform != "darwin" and win32com is None:
            messagebox.showerror("Missing dependency", "This needs pywin32. Install with:\n    pip install pywin32")
            return

        resize_width = None
        resize_dpi = None
        if self.do_resize.get():
            try:
                resize_width = int(self.resize_width.get())
                resize_dpi = int(self.resize_dpi.get())
                if resize_width <= 0 or resize_dpi <= 0:
                    raise ValueError()
            except ValueError:
                messagebox.showerror("Invalid input", "Width and DPI must be positive whole numbers.")
                return

        source = self.html_file_path.get().strip()
        is_url = False
        do_extract = bool(source)

        if not do_extract and not self._data:
            messagebox.showerror(
                "Nothing to run",
                "No JSON is loaded and no extraction source was given. Either "
                "browse to a saved HTML/MHTML file, or make sure a valid "
                "textboxes.json is already in the folder."
            )
            return

        if sync_playwright is None and do_extract:
            messagebox.showerror(
                "Missing dependency",
                "Extraction needs the 'playwright' package. Install with:\n"
                "    pip install playwright\n"
                "(no browser download needed -- this uses your existing installed Chrome/Edge)"
            )
            return

        if do_extract and not os.path.isfile(source):
            messagebox.showerror("File not found", "The selected HTML/MHTML file no longer exists:\n" + source)
            return

        confirm_lines = []
        if do_extract:
            confirm_lines.append("1) Open %s and extract textboxes.json (opens a Chrome/Edge window)." % (
                os.path.basename(source)
            ))
        confirm_lines.append("%d) Make sure Photoshop is already OPEN, then process %d file(s)." % (
            2 if do_extract else 1, len(self._files_info)
        ))
        if not messagebox.askyesno("Confirm", "\n".join(confirm_lines)):
            return

        self._stop_event.clear()
        self.run_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        threading.Thread(
            target=self._run_pipeline_thread,
            args=(do_extract, source, is_url, resize_width, resize_dpi),
            daemon=True
        ).start()

    def stop_batch(self):
        self._stop_event.set()
        self.stop_btn.config(state='disabled')
        self._status("Stopping after the current file finishes...")
        self._log("\n>>> Stop requested -- finishing the file in progress, then stopping.")

    def _run_pipeline_thread(self, do_extract, source, is_url, resize_width, resize_dpi):
        if pythoncom is not None:
            pythoncom.CoInitialize()
        try:
            self.after(0, lambda: self._log("Pipeline build: %s" % PIPELINE_BUILD))
            # ── Step 1: extract textboxes.json from the chosen source, if requested ──
            if do_extract:
                output_json_path = os.path.join(self.folder.get(), "textboxes.json")
                label = source if is_url else os.path.basename(source)
                self.after(0, lambda: self._status("Extracting text data from %s ..." % label))
                self.after(0, lambda: self._log("\n=== Extracting textboxes.json from %s ===" % label))
                try:
                    extract_textboxes(
                        source, is_url, output_json_path,
                        log_cb=lambda m: self.after(0, lambda m=m: self._log(m))
                    )
                except Exception as e:
                    self.after(0, lambda e=e: messagebox.showerror("Extraction failed", str(e)))
                    return

                with open(output_json_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if not data.get("pages"):
                    self.after(0, lambda: messagebox.showerror("Extraction failed", "Downloaded JSON has no pages."))
                    return
                self._json_path = output_json_path
                self._data = data
                self.after(0, lambda: self._log(
                    "Extraction done: %d page(s).\n" % len(data["pages"])
                ))

            if self._stop_event.is_set():
                self.after(0, lambda: self._log("\nStopped before processing started."))
                return

            # ── Step 2: build the placement plan. This needs no resize size
            #     at all (predicted or real) -- it works entirely in each
            #     file's original dimensions, which are scale-invariant, so
            #     it's exactly as correct whether or not resizing happens
            #     later. ─────────────────────────────────────────────────────
            self.after(0, lambda: self._status("Building placement plan..."))
            dbg = debug_scale_info(self._files_info, self._data)
            if dbg is not None:
                self.after(0, lambda dbg=dbg: self._log(
                    "Scale check: web_total_h=%.2f  orig_total_h(sum of PSD headers)=%.2f  design_scale=%.6f" %
                    (dbg["web_total_h"], dbg["orig_total_h"], dbg["design_scale"])
                ))
            try:
                per_file_boxes = build_plan(self._files_info, self._data)
            except ValueError as e:
                self.after(0, lambda e=e: messagebox.showerror("Cannot build plan", str(e)))
                return

            do_resize = resize_width is not None
            resolved_dims = compute_resolved_dims(self._files_info, resize_width)  # for the log line only
            total_boxes = sum(len(b) for b in per_file_boxes)
            self.after(0, lambda: self._log("=== Plan ==="))
            for i, f in enumerate(self._files_info):
                dim = resolved_dims[i]
                note = " (resized to %dx%d)" % (dim["w"], dim["h"]) if do_resize else ""
                n = len(per_file_boxes[i])
                self.after(0, lambda f=f, n=n, note=note: self._log("  %-40s -> %d text box(es)%s" % (f["name"], n, note)))
            self.after(0, lambda: self._log("Total: %d text box(es) across %d file(s).\n" % (total_boxes, len(self._files_info))))

            # ── Step 3: connect to Photoshop ──────────────────────────────────
            self.after(0, lambda: self._status("Connecting to Photoshop..."))
            self.after(0, lambda: self._log("Connecting to Photoshop..."))
            ps_app = None
            connect_errors = []

            if sys.platform == "darwin":
                try:
                    ps_app = connect_mac_photoshop()
                except Exception as e:
                    connect_errors.append(str(e))
            else:
                try:
                    ps_app = win32com.client.GetActiveObject("Photoshop.Application")
                except Exception as e:
                    connect_errors.append("GetActiveObject: %s" % e)

                if ps_app is None:
                    try:
                        ps_app = win32com.client.Dispatch("Photoshop.Application")
                    except Exception as e:
                        connect_errors.append("Dispatch: %s" % e)

            if ps_app is None:
                msg = (
                    "Could not connect to Photoshop.\n\n"
                    "Make sure Photoshop is already open (not just installed), "
                    "and that it is NOT set to \"Run as administrator\" "
                    "(that must match how this script is run).\n\nDetails:\n" +
                    "\n".join(connect_errors)
                )
                self.after(0, lambda m=msg: messagebox.showerror("Connection failed", m))
                return

            ps_app.DoJavaScript("app.displayDialogs = DialogModes.NO;")
            self.after(0, lambda: self._log("Connected."))

            # ── Step 4: resize + place text for each file, one open/close ────
            total_ok = 0
            stopped_early = False
            for i, f in enumerate(self._files_info):
                if self._stop_event.is_set():
                    stopped_early = True
                    self.after(0, lambda: self._log(
                        "\nStopped: %d/%d file(s) were not processed." % (len(self._files_info) - i, len(self._files_info))
                    ))
                    break

                box_list = per_file_boxes[i]
                self.after(0, lambda i=i, f=f: self._status(
                    "Processing %d/%d: %s ..." % (i + 1, len(self._files_info), f["name"])
                ))
                try:
                    result = run_one_file(ps_app, f["path"], box_list, do_resize, resize_width, resize_dpi)
                except Exception as e:
                    result = {"ok": 0, "total": len(box_list), "errors": [str(e)]}

                total_ok += result["ok"]
                msg = "[%d/%d] %s: %d/%d placed" % (
                    i + 1, len(self._files_info), f["name"], result["ok"], result["total"]
                )
                self.after(0, lambda m=msg: self._log(m))
                for err in result.get("errors", []):
                    prefix = "    - " if err.startswith("NOTE:") else "    ! "
                    self.after(0, lambda e=err, p=prefix: self._log(p + e))

            if stopped_early:
                final_msg = "\nStopped by user. %d/%d text boxes placed before stopping." % (total_ok, total_boxes)
            else:
                final_msg = "\nDone. %d/%d text boxes placed across %d file(s)." % (
                    total_ok, total_boxes, len(self._files_info)
                )
            self.after(0, lambda: self._log(final_msg))
            self.after(0, lambda: self.summary.config(text=final_msg.strip()))
            self.after(0, lambda: self._status(""))

        except Exception as e:
            traceback.print_exc()
            self.after(0, lambda e=e: messagebox.showerror("Error", str(e)))
        finally:
            if pythoncom is not None:
                pythoncom.CoUninitialize()
            self.after(0, lambda: self.run_btn.config(state='normal'))
            self.after(0, lambda: self.stop_btn.config(state='disabled'))



# ══════════════════════════════════════════════════════════════════
#  TAB 2: PSD <-> WEBSITE TEXT CHECKER
#  (merged from check_type.py -- compares "#Bx"/"#x" translation boxes
#  on a saved website export against "Bx_..." Photoshop text layers)
# ══════════════════════════════════════════════════════════════════

CODE_PATTERNS = [
    re.compile(r'^B(\d+)_', re.IGNORECASE),        # e.g. "B1_arien !"
    re.compile(r'^DKI:\s*(\d+)_', re.IGNORECASE),  # e.g. "DKI: 112_..."
]


def _match_layer_code(layer_name):
    """Tries each known layer-naming convention in turn. Returns the
    normalized code (e.g. "B112") or None if nothing matches."""
    name = layer_name.strip()
    for pattern in CODE_PATTERNS:
        m = pattern.match(name)
        if m:
            return f"B{m.group(1)}".upper()
    return None

CHECKER_HTML_EXTS = {'.mhtml', '.html', '.htm'}
CHECKER_PSD_EXTS = {'.psd', '.psb'}


# ---- Website (.mhtml / .html) parsing --------------------------------

def load_html_text(path):
    path = Path(path)
    if path.suffix.lower() == '.mhtml':
        with open(path, 'rb') as f:
            msg = email.message_from_binary_file(f, policy=email_policy.default)
        for part in msg.walk():
            if part.get_content_type() == 'text/html':
                raw = part.get_payload(decode=True)
                charset = part.get_content_charset() or 'utf-8'
                try:
                    return raw.decode(charset)
                except (LookupError, UnicodeDecodeError):
                    return raw.decode('utf-8', errors='replace')
        raise ValueError(f"No text/html part found in {path}")
    else:
        return path.read_text(encoding='utf-8', errors='replace')


def _parse_style_attr(style_str):
    """Returns (bold, italic, superscript) implied by an inline
    `style="..."` attribute."""
    style_str = (style_str or '').lower()
    bold = False
    italic = False
    superscript = False
    m = re.search(r'font-weight\s*:\s*([a-z0-9]+)', style_str)
    if m:
        val = m.group(1)
        if val == 'bold' or (val.isdigit() and int(val) >= 600):
            bold = True
    m = re.search(r'font-style\s*:\s*([a-z]+)', style_str)
    if m and m.group(1) == 'italic':
        italic = True
    m = re.search(r'vertical-align\s*:\s*([a-z-]+)', style_str)
    if m and m.group(1) in ('super', 'text-top', 'top'):
        superscript = True
    return bold, italic, superscript


def _walk_inline(node, bold, italic, superscript, out_chars, out_styles):
    """Recursively walks an HTML node, tracking bold/italic/superscript
    state, and appends (char, (bold, italic, superscript)) pairs for
    every character of text."""
    if isinstance(node, NavigableString):
        for ch in str(node):
            out_chars.append(ch)
            out_styles.append((bold, italic, superscript))
        return
    if not isinstance(node, Tag):
        return

    tag = node.name.lower()
    new_bold, new_italic, new_super = bold, italic, superscript
    if tag in ('strong', 'b'):
        new_bold = True
    if tag in ('em', 'i'):
        new_italic = True
    if tag in ('sup',):
        new_super = True
    if tag in ('sub',):
        # Subscript isn't a format we track separately -- treat it as
        # "not superscript" (distinct from normal text isn't requested).
        pass
    style_attr = node.get('style')
    if style_attr:
        sb, si, ss = _parse_style_attr(style_attr)
        new_bold = new_bold or sb
        new_italic = new_italic or si
        new_super = new_super or ss

    if tag == 'br':
        out_chars.append('\n')
        out_styles.append((new_bold, new_italic, new_super))
        return

    for child in node.children:
        _walk_inline(child, new_bold, new_italic, new_super, out_chars, out_styles)


def _node_to_text_and_styles(container):
    """
    Given a BeautifulSoup element (or fragment) that contains one or more
    <p> paragraphs, returns (text, styles) where styles is a list of
    (bold, italic, superscript) tuples, one per character of text
    (paragraph breaks become '\\n' with style (False, False, False)).
    """
    chars, styles = [], []
    if container is None:
        return '', []

    paragraphs = container.find_all('p')
    nodes = paragraphs if paragraphs else [container]

    for i, p in enumerate(nodes):
        if i > 0:
            chars.append('\n')
            styles.append((False, False, False))
        for child in p.children:
            _walk_inline(child, False, False, False, chars, styles)

    return ''.join(chars), styles


def extract_website_boxes(path):
    """
    Auto-detects which of the supported website formats this export
    uses, and returns {code: (text, styles)}.

    FORMAT A ("panel" review sidebar):
        <article class="panel__body__list__item__inner">
          <span class="panel__body__list__item__inner__header__id">#B1</span>
          ... text inside <p class="paragraph"> ...
        </article>
        Codes look like "#B1", "#B2", ...

    FORMAT B ("box overlay" with a JSON blob per box):
        <article data-item='{"id":1,"content":"<p>rustle</p>",...}'>
        Codes look like "#1", "#2", ... (no "B" -- box id is a plain number)

    FORMAT C ("translation management" button list):
        <span class="text-on-surface-variant-1 text-xs!">#1</span>
        ... followed by a sibling <div> holding the text (with possible
        <strong>/<em>/<b>/<i> formatting) ...
        Codes look like "#1", "#2", ... (same as Format B)

    Either way, the returned dict keys are normalized to "B<number>"
    (e.g. "B1") so they line up with the PSD layer codes.
    """
    html = load_html_text(path)
    soup = BeautifulSoup(html, 'html.parser')

    boxes = _extract_format_a(soup)
    if boxes:
        return boxes

    boxes = _extract_format_b(soup)
    if boxes:
        return boxes

    boxes = _extract_format_c(soup)
    if boxes:
        return boxes

    raise ValueError(
        f"Could not find any recognizable translation boxes in {path}.\n"
        "None of the supported website formats were detected."
    )


def _extract_format_a(soup):
    boxes = {}
    for item in soup.select('article.panel__body__list__item__inner'):
        header = item.select_one('.panel__body__list__item__inner__header__id')
        if not header:
            continue
        code = header.get_text(strip=True).lstrip('#').strip()
        if not code:
            continue

        text_container = item.select_one('.panel__body__list__item__inner__text')
        text, styles = _node_to_text_and_styles(text_container)
        if not text.strip():
            # Blank textbox on the website (e.g. a placeholder/unused box
            # that was never filled in) -- skip it so it doesn't show up
            # as a spurious "missing in PSD" row, and (for position-based
            # matching) doesn't throw off the top-to-bottom pairing either.
            continue

        norm_code = code if code.upper().startswith('B') else f"B{code}"
        boxes[norm_code.upper()] = (text, styles)
    return boxes


def _extract_format_b(soup):
    boxes = {}
    for article in soup.select('article[data-item]'):
        raw = article.get('data-item')
        if not raw:
            continue
        try:
            data = json.loads(raw)
        except (json.JSONDecodeError, TypeError):
            continue

        box_id = data.get('id')
        if box_id is None:
            continue

        content_html = data.get('content') or ''
        content_soup = BeautifulSoup(content_html, 'html.parser')
        text, styles = _node_to_text_and_styles(content_soup)
        if not text.strip():
            # Blank textbox -- skip (see note in _extract_format_a).
            continue

        boxes[f"B{box_id}".upper()] = (text, styles)
    return boxes


def _extract_format_c(soup):
    """
    Handles a "translation management" style list, e.g.:
        <span class="text-on-surface-variant-1 text-xs!">#1</span>
        ...
        <div class="text-primary ...">si j'étais ton frère...</div>
    where the text div is the next sibling of the code's containing
    header row.
    """
    boxes = {}
    for span in soup.select('span.text-on-surface-variant-1'):
        raw_code = span.get_text(strip=True)
        if not raw_code.startswith('#'):
            continue
        code_num = raw_code.lstrip('#').strip()
        if not code_num.isdigit():
            continue

        header_row = span.find_parent('div', class_='flex')
        if header_row is None:
            continue
        text_div = header_row.find_next_sibling('div')
        if text_div is None:
            continue

        text, styles = _node_to_text_and_styles(text_div)
        if not text.strip():
            # Blank textbox -- skip (see note in _extract_format_a).
            continue
        boxes[f"B{code_num}".upper()] = (text, styles)

    return boxes


# ---- Photoshop (.psd / .psb) parsing ----------------------------------

def _font_implies_bold(font_name):
    name = str(font_name or '').lower()
    return any(tag in name for tag in ('bold', 'black', 'heavy', 'extrabold'))


def _font_implies_italic(font_name):
    name = str(font_name or '').lower()
    return any(tag in name for tag in ('italic', 'oblique'))


def _extract_layer_text_and_styles(layer):
    """
    Returns (text, styles) for a single Photoshop text layer, where
    styles is a list of (bold, italic, superscript) tuples, one per
    character.

    Bold/italic is True if EITHER:
      - the run's Faux Bold / Faux Italic toggle is on, OR
      - the actual font used for that run has "Bold"/"Italic"/etc.
        in its name (covers real bold/italic font variants).

    Superscript is True if the run's baseline style (Photoshop's
    "Superscript" character option) is set to superscript.
    """
    text = (layer.text or '').replace('\r', '\n')

    try:
        engine = layer.engine_dict or {}
        style_run = engine.get('StyleRun', {})
        lengths = style_run.get('RunLengthArray', [])
        runs = style_run.get('RunArray', [])
        fontset = (layer.document_resources or {}).get('FontSet', [])
    except Exception:
        lengths, runs, fontset = [], [], []

    styles = []
    for length, run in zip(lengths, runs):
        sheet = run.get('StyleSheet', {}).get('StyleSheetData', {})
        faux_bold = bool(sheet.get('FauxBold', False))
        faux_italic = bool(sheet.get('FauxItalic', False))
        # FontBaseline: 0 = normal, 1 = superscript, 2 = subscript.
        font_baseline = sheet.get('FontBaseline', 0)
        try:
            superscript = int(font_baseline) == 1
        except (TypeError, ValueError):
            superscript = False
        font_idx = sheet.get('Font', 0)
        font_name = None
        try:
            font_idx = int(font_idx)
            if 0 <= font_idx < len(fontset):
                font_name = fontset[font_idx].get('Name')
        except (TypeError, ValueError):
            font_name = None

        bold = faux_bold or _font_implies_bold(font_name)
        italic = faux_italic or _font_implies_italic(font_name)
        styles.extend([(bold, italic, superscript)] * length)

    # Safety net: if run lengths don't add up to the text length
    # (can happen with unusual documents), pad/trim rather than crash.
    if len(styles) < len(text):
        styles.extend([(False, False, False)] * (len(text) - len(styles)))
    elif len(styles) > len(text):
        styles = styles[:len(text)]

    return text, styles


def _extract_smart_object_text(layer):
    """
    Best-effort attempt to read text out of a text layer that's been
    converted into a Smart Object. Photoshop still embeds the full
    PSD/PSB content inside a "Convert to Smart Object" layer, so if
    that's what happened here, we can open it and pull the text back
    out of whatever type layers are inside it.

    Returns (text, styles) if any text was found inside, or None if
    the smart object isn't PSD/PSB content (e.g. it's a flattened
    image, or an externally-linked file) so its text can't be read
    this way -- callers should fall back to marking it as "converted"
    rather than "missing" in that case.
    """
    try:
        so = layer.smart_object
        filetype = (so.detected_filetype or '').lower()
        if filetype not in ('psd', 'psb'):
            return None
        data = so.data
    except Exception:
        return None

    try:
        inner = PSDImage.open(io.BytesIO(data))
    except Exception:
        return None

    chars, styles = [], []

    def walk_inner(layer_list):
        for sub in layer_list:
            try:
                if sub.kind == 'type':
                    t, s = _extract_layer_text_and_styles(sub)
                    if t:
                        if chars:
                            chars.append('\n')
                            styles.append((False, False, False))
                        chars.extend(t)
                        styles.extend(s)
                if sub.is_group():
                    walk_inner(sub)
            except Exception:
                continue

    try:
        walk_inner(inner)
    except Exception:
        return None

    if not chars:
        return None
    return ''.join(chars), styles


def extract_psd_layers(path):
    """
    Returns (layers, smart_object_codes, layer_tops):
      - layers: {code: (text, styles)} for every readable text layer,
        including text recovered from inside PSD/PSB smart objects.
      - smart_object_codes: codes that were found as a smart-object
        layer whose text could NOT be automatically read (e.g. it was
        flattened to an image, or links to an external file). These
        should be treated as "converted, please check manually"
        rather than "missing".
      - layer_tops: {code: top_y_pixel} -- each text layer's vertical
        position on the canvas. This lets callers sort layers by their
        actual top-to-bottom position on the page, which is more
        reliable than sorting by the code/mark number when the mark
        numbers themselves have gotten out of sync with the website
        (e.g. a layer named "B184" that's actually positioned between
        "B72" and "B74").
    """
    psd = PSDImage.open(path)
    layers = {}
    smart_object_codes = set()
    layer_tops = {}

    def walk(layer_list):
        for layer in layer_list:
            if layer.kind == 'type':
                code = _match_layer_code(layer.name)
                if code:
                    if code not in layers:
                        layers[code] = _extract_layer_text_and_styles(layer)
                        layer_tops[code] = layer.top
            elif layer.kind == 'smartobject':
                code = _match_layer_code(layer.name)
                if code and code not in layers:
                    extracted = _extract_smart_object_text(layer)
                    if extracted is not None:
                        layers[code] = extracted
                        layer_tops[code] = layer.top
                    else:
                        smart_object_codes.add(code)
                        layer_tops.setdefault(code, layer.top)
            if layer.is_group():
                walk(layer)

    walk(psd)
    return layers, smart_object_codes, layer_tops


# ---- Comparison --------------------------------------------------------

def _collapse_whitespace(text, styles=None):
    """
    Collapses every run of whitespace (including newlines) into a single
    space, and strips leading/trailing whitespace. Used both to ignore
    line-break differences in the text itself, and to line up character
    positions for the formatting check regardless of how each side wraps
    its lines.
    """
    out_chars = []
    out_styles = [] if styles is not None else None
    prev_was_space = False

    for i, ch in enumerate(text):
        if ch.isspace():
            if out_chars and not prev_was_space:
                out_chars.append(' ')
                if out_styles is not None:
                    out_styles.append(styles[i])
            prev_was_space = True
        else:
            out_chars.append(ch)
            if out_styles is not None:
                out_styles.append(styles[i])
            prev_was_space = False

    while out_chars and out_chars[-1] == ' ':
        out_chars.pop()
        if out_styles is not None:
            out_styles.pop()

    return ''.join(out_chars), out_styles


def normalize_text(text, ignore_case=False, ignore_whitespace=False, ignore_linebreaks=False):
    if ignore_linebreaks:
        text, _ = _collapse_whitespace(text)
    elif ignore_whitespace:
        text = '\n'.join(line.strip() for line in text.splitlines()).strip()
    if ignore_case:
        text = text.lower()
    return text


def _style_label(bold, italic, superscript=False, check_bold_italic=True, check_superscript=True):
    parts = []
    if check_bold_italic and bold:
        parts.append('bold')
    if check_bold_italic and italic:
        parts.append('italic')
    if check_superscript and superscript:
        parts.append('superscript')
    return '+'.join(parts) if parts else 'normal'


def _format_diff_label(mismatches):
    """
    Turns a list of per-character mismatches (each carrying diff_bold /
    diff_italic / diff_super flags) into a short label describing what
    kind of formatting differs, e.g. 'Bold', 'Italic', 'bold-ita', 'SS',
    or a combo like 'Bold+SS'.
    """
    any_bold = any(m.get('diff_bold') for m in mismatches)
    any_italic = any(m.get('diff_italic') for m in mismatches)
    any_super = any(m.get('diff_super') for m in mismatches)

    parts = []
    if any_bold and any_italic:
        parts.append('bold-ita')
    elif any_bold:
        parts.append('Bold')
    elif any_italic:
        parts.append('Italic')
    if any_super:
        parts.append('SS')

    return '+'.join(parts) if parts else 'MISMATCH'


def _extra_psd_label(extra_in_psd):
    """
    Describes bold/italic formatting that's present in the PSD but not on
    the website. This is informational only -- it's never reported as a
    mismatch, just noted on an otherwise-matching row.
    """
    any_bold = any(m.get('extra_bold') for m in extra_in_psd)
    any_italic = any(m.get('extra_italic') for m in extra_in_psd)

    parts = []
    if any_bold and any_italic:
        parts.append('bold-ita')
    elif any_bold:
        parts.append('Bold')
    elif any_italic:
        parts.append('Italic')

    return '+'.join(parts)


def compare_formatting(web_text, web_styles, psd_text, psd_styles, ignore_case=False,
                        check_bold_italic=True, check_superscript=True):
    """
    Compares formatting character-by-character. Text is always aligned
    by collapsing whitespace/line-breaks first (formatting can't rely on
    exact line-break positions matching between a website export and a
    hand-wrapped PSD layer).

    check_bold_italic / check_superscript control which attribute(s) are
    actually compared -- a difference in an attribute that isn't being
    checked is not reported as a mismatch.

    Bold/italic is asymmetric on purpose:
      - Website has bold/italic that the PSD doesn't -> reported as a
        real MISMATCH (the site is showing formatting that isn't there).
      - PSD has bold/italic that the website doesn't -> NOT a mismatch,
        just noted informationally (returned separately as
        `extra_in_psd`) -- e.g. the PSD used faux-bold for emphasis that
        never made it onto the website, which is fine.
    Superscript is unaffected -- it's still a plain two-way comparison.

    Returns (status, mismatches, extra_in_psd) where status is 'MATCH',
    'MISMATCH', or 'N/A' (text doesn't line up character-for-character,
    so formatting can't be reliably compared).
    """
    w_text, w_styles = _collapse_whitespace(web_text, web_styles)
    p_text, p_styles = _collapse_whitespace(psd_text, psd_styles)

    wt_cmp = w_text.lower() if ignore_case else w_text
    pt_cmp = p_text.lower() if ignore_case else p_text

    if wt_cmp != pt_cmp or len(w_styles) != len(p_styles):
        return 'N/A', [], []

    mismatches = []
    extra_in_psd = []
    for i, ch in enumerate(w_text):
        w_bold, w_italic, w_super = w_styles[i]
        p_bold, p_italic, p_super = p_styles[i]

        # Website has it, PSD doesn't -> real mismatch.
        diff_bold = check_bold_italic and (w_bold and not p_bold)
        diff_italic = check_bold_italic and (w_italic and not p_italic)
        diff_super = check_superscript and (w_super != p_super)

        # PSD has it, website doesn't -> informational only.
        extra_bold = check_bold_italic and (p_bold and not w_bold)
        extra_italic = check_bold_italic and (p_italic and not w_italic)

        if diff_bold or diff_italic or diff_super:
            mismatches.append({
                'index': i,
                'char': ch,
                'website': w_styles[i],
                'psd': p_styles[i],
                'diff_bold': diff_bold,
                'diff_italic': diff_italic,
                'diff_super': diff_super,
            })
        elif extra_bold or extra_italic:
            extra_in_psd.append({
                'index': i,
                'char': ch,
                'website': w_styles[i],
                'psd': p_styles[i],
                'extra_bold': extra_bold,
                'extra_italic': extra_italic,
            })

    status = 'MISMATCH' if mismatches else 'MATCH'
    return status, mismatches, extra_in_psd


def compare(website_boxes, psd_layers, website_sources=None, psd_sources=None,
            smart_object_codes=None, ignore_case=False, ignore_whitespace=False,
            ignore_linebreaks=False, check_bold_italic=False, check_superscript=False):
    """
    website_boxes / psd_layers: {code: (text, styles)}
    website_sources / psd_sources: {code: source filename}, optional
    smart_object_codes: codes that are a smart-object layer in the PSD
        whose text couldn't be read automatically -- reported as
        'CONVERTED' instead of 'MISSING_IN_PSD'.
    check_bold_italic / check_superscript: which formatting attribute(s)
        to compare; if both are False formatting isn't checked at all.
    Returns a list of row dicts.
    """
    website_sources = website_sources or {}
    psd_sources = psd_sources or {}
    smart_object_codes = smart_object_codes or set()
    check_formatting = check_bold_italic or check_superscript

    rows = []
    all_codes = sorted(
        set(website_boxes) | set(psd_layers) | set(smart_object_codes),
        key=lambda c: int(re.sub(r'\D', '', c) or 0)
    )

    for code in all_codes:
        web = website_boxes.get(code)
        psd = psd_layers.get(code)
        note = ''

        if web is None:
            status = 'MISSING_ON_WEBSITE'
        elif psd is None:
            if code in smart_object_codes:
                status = 'CONVERTED'
                note = ("Converted to a Smart Object in Photoshop -- its text "
                         "couldn't be read automatically. Please check it manually.")
            else:
                status = 'MISSING_IN_PSD'
        else:
            web_text, _ = web
            psd_text, _ = psd
            a = normalize_text(web_text, ignore_case, ignore_whitespace, ignore_linebreaks)
            b = normalize_text(psd_text, ignore_case, ignore_whitespace, ignore_linebreaks)
            status = 'MATCH' if a == b else 'MISMATCH'

        format_status = 'N/A'
        format_status_display = 'N/A'
        format_details = ''
        if check_formatting and web is not None and psd is not None:
            format_status, mismatches, extra_in_psd = compare_formatting(
                web[0], web[1], psd[0], psd[1], ignore_case,
                check_bold_italic=check_bold_italic,
                check_superscript=check_superscript,
            )
            if format_status == 'MISMATCH':
                format_status_display = _format_diff_label(mismatches)
                parts = []
                for m in mismatches[:6]:
                    parts.append(
                        f"{m['char']!r} web={_style_label(*m['website'], check_bold_italic, check_superscript)} "
                        f"psd={_style_label(*m['psd'], check_bold_italic, check_superscript)}"
                    )
                format_details = '; '.join(parts)
                if len(mismatches) > 6:
                    format_details += f' ... (+{len(mismatches) - 6} more)'
            elif extra_in_psd:
                # PSD has bold/italic the website doesn't -- not a
                # mismatch, just a note on an otherwise-matching row.
                format_status_display = f"MATCH (PSD only: {_extra_psd_label(extra_in_psd)})"
                format_details = f"PSD has {_extra_psd_label(extra_in_psd)} not present on the website (not an error)."
            else:
                format_status_display = format_status

        rows.append({
            'code': code,
            'status': status,
            'format_status': format_status,
            'format_status_display': format_status_display,
            'format_details': format_details,
            'details': note or format_details,
            'website_text': web[0] if web else '',
            'psd_text': psd[0] if psd else '',
            'website_file': website_sources.get(code, ''),
            'psd_file': psd_sources.get(code, ''),
        })

    return rows


def _align_sequences(website_items, psd_items, ignore_case=False,
                      ignore_whitespace=False, ignore_linebreaks=False):
    """
    Lines up two ordered lists of (code, value) items by their actual TEXT
    CONTENT, using the same longest-common-subsequence technique text-diff
    tools use (Python's difflib) -- instead of naively pairing item #1
    with item #1, item #2 with item #2, and so on by raw index.

    Why this matters: with plain index-pairing, a single box that's
    inserted, deleted, or reordered on just ONE side permanently shifts
    every later index out of sync, so every following row falsely reports
    as a mismatch even though the actual text lines up fine. Diffing by
    content instead finds the long runs that really do match first (their
    text is identical), and only the genuinely different stretch in
    between is treated as suspect -- normal matching resumes immediately
    afterward instead of cascading through the rest of the list.

    Returns a list of (website_item_or_None, psd_item_or_None) pairs, in
    reading order, where each item is the original (code, value) tuple
    (value is None for a PSD smart-object placeholder with no readable
    text).
    """
    def _key(value, idx):
        if value is None:
            # Smart-object text can't be read -- give it a per-item
            # unique key so it never spuriously "equals" another
            # smart-object placeholder or an empty string elsewhere.
            return f"\x00__NO_TEXT__{idx}\x00"
        text, _ = value
        return normalize_text(text, ignore_case, ignore_whitespace, ignore_linebreaks)

    web_keys = [_key(val, i) for i, (_, val) in enumerate(website_items)]
    psd_keys = [_key(val, i) for i, (_, val) in enumerate(psd_items)]

    matcher = difflib.SequenceMatcher(None, web_keys, psd_keys, autojunk=False)

    pairs = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            for k in range(i2 - i1):
                pairs.append((website_items[i1 + k], psd_items[j1 + k]))
        elif tag == 'replace':
            # Genuinely different stretch on both sides -- pair up
            # whatever overlaps 1:1 (this is the normal "same slot, real
            # mismatch" case), and treat any leftover on the longer side
            # as missing on the other side, rather than dragging the
            # whole rest of the list out of alignment.
            w_len, p_len = i2 - i1, j2 - j1
            common = min(w_len, p_len)
            for k in range(common):
                pairs.append((website_items[i1 + k], psd_items[j1 + k]))
            for k in range(common, w_len):
                pairs.append((website_items[i1 + k], None))
            for k in range(common, p_len):
                pairs.append((None, psd_items[j1 + k]))
        elif tag == 'delete':
            # Only in the website list -- no PSD counterpart here.
            for k in range(i1, i2):
                pairs.append((website_items[k], None))
        elif tag == 'insert':
            # Only in the PSD list -- no website counterpart here.
            for k in range(j1, j2):
                pairs.append((None, psd_items[k]))

    return pairs


def _repair_orphan_pairs(pairs, ignore_case=False, ignore_whitespace=False,
                          ignore_linebreaks=False, window=6):
    """
    Safety net that runs after _align_sequences().

    Rarely, two boxes that carry IDENTICAL text end up as two separate
    unmatched rows -- one "missing on website", one "missing in PSD" --
    instead of being paired with each other. This happens when a
    duplicate/near-duplicate line elsewhere in the same script pulls the
    LCS alignment's attention onto a different, more distant equal run,
    leaving the real match sitting right next to it unpaired.

    This pass looks in a small window around every unmatched
    website-only row for an unmatched PSD-only row with the exact same
    normalized text, and pairs them together -- fixing exactly that
    "same text, but counted as two separate mismatched lines" case
    without touching anything that's already matched or genuinely
    missing.
    """
    def _key(item):
        text, _ = item[1]
        return normalize_text(text, ignore_case, ignore_whitespace, ignore_linebreaks)

    def _has_text(entry):
        return entry is not None and entry[1] is not None

    result = list(pairs)
    n = len(result)
    for i in range(n):
        web_entry, psd_entry = result[i]
        if not _has_text(web_entry) or psd_entry is not None:
            continue  # only look from "website-only orphan" rows (with real text)
        target_key = _key(web_entry)
        if not target_key.strip():
            continue
        lo, hi = max(0, i - window), min(n, i + window + 1)
        for j in range(lo, hi):
            if j == i:
                continue
            w2, p2 = result[j]
            if w2 is not None or not _has_text(p2):
                continue  # only look at "PSD-only orphan" rows (with real text)
            if _key(p2) == target_key:
                result[i] = (web_entry, p2)
                result[j] = (None, None)  # consumed -- dropped below
                break

    return [(w, p) for (w, p) in result if not (w is None and p is None)]


def compare_by_position(website_boxes, psd_layers, website_sources=None, psd_sources=None,
                         smart_object_codes=None, psd_positions=None, ignore_case=False,
                         ignore_whitespace=False, ignore_linebreaks=False,
                         check_bold_italic=False, check_superscript=False):
    """
    Same idea as compare(), but instead of requiring the website box's code
    to be IDENTICAL to the PSD layer's code, it pairs them up by POSITION:

      - website boxes are compared in the order extract_website_boxes()
        found them in the page (top-to-bottom, since that's the order
        BeautifulSoup encounters them in the document and dicts preserve
        insertion order)
      - PSD layers are sorted by their real top-to-bottom position on the
        canvas (psd_positions), when available -- because the mark number
        baked into a layer's name (e.g. "B184") is exactly the kind of
        thing that can drift out of sync: a layer can be named "B184" but
        actually sit between "B72" and "B74" on the page (e.g. because it
        was added later and tacked onto the end of the numbering). Sorting
        by the *actual pixel position* rather than the number avoids that
        trap. If no position is known for a code, it falls back to sorting
        by the numeric code instead.

    Then, rather than naively pairing the 1st website box with the 1st
    PSD layer, the 2nd with the 2nd, and so on (which would cause a
    single inserted/missing box to shift every following pair out of
    sync), the two sequences are aligned by their actual TEXT CONTENT
    using an LCS-based diff (see _align_sequences()) -- so matching
    automatically resyncs right after any genuinely inserted, missing,
    or reordered box instead of cascading false mismatches through the
    rest of the list.

    Use this mode when the website's mark numbers have drifted out of sync
    with the PSD (e.g. a box was inserted later and got tacked onto the end
    of the numbering, or the site renumbers boxes 1..N by current on-screen
    order) but the underlying top-to-bottom reading order is still the same
    on both sides.
    """
    website_sources = website_sources or {}
    psd_sources = psd_sources or {}
    smart_object_codes = smart_object_codes or set()
    psd_positions = psd_positions or {}
    check_formatting = check_bold_italic or check_superscript

    # Website boxes: keep them in the order they were found (top-to-bottom).
    website_items = list(website_boxes.items())

    # PSD layers: sort by real canvas position when known, otherwise fall
    # back to the numeric code. Smart-object-only codes (no readable text)
    # still need a slot so the sequence lines up.
    psd_items = dict(psd_layers.items())
    for code in smart_object_codes:
        psd_items.setdefault(code, None)

    def _sort_key(kv):
        code = kv[0]
        if code in psd_positions:
            return (0, psd_positions[code])
        return (1, int(re.sub(r'\D', '', code) or 0))

    psd_items = sorted(psd_items.items(), key=_sort_key)

    rows = []
    aligned_pairs = _align_sequences(
        website_items, psd_items,
        ignore_case=ignore_case,
        ignore_whitespace=ignore_whitespace,
        ignore_linebreaks=ignore_linebreaks,
    )
    aligned_pairs = _repair_orphan_pairs(
        aligned_pairs,
        ignore_case=ignore_case,
        ignore_whitespace=ignore_whitespace,
        ignore_linebreaks=ignore_linebreaks,
    )
    for web_entry, psd_entry in aligned_pairs:
        web_code, web = web_entry if web_entry is not None else (None, None)
        psd_code, psd = psd_entry if psd_entry is not None else (None, None)
        pair_label = f"{web_code or '—'} / {psd_code or '—'}"
        note = ''

        if web is None:
            status = 'MISSING_ON_WEBSITE'
        elif psd is None:
            if psd_code in smart_object_codes:
                status = 'CONVERTED'
                note = ("Converted to a Smart Object in Photoshop -- its text "
                        "couldn't be read automatically. Please check it manually.")
            else:
                status = 'MISSING_IN_PSD'
        else:
            web_text, _ = web
            psd_text, _ = psd
            a = normalize_text(web_text, ignore_case, ignore_whitespace, ignore_linebreaks)
            b = normalize_text(psd_text, ignore_case, ignore_whitespace, ignore_linebreaks)
            status = 'MATCH' if a == b else 'MISMATCH'

        format_status = 'N/A'
        format_status_display = 'N/A'
        format_details = ''
        if check_formatting and web is not None and psd is not None:
            format_status, mismatches, extra_in_psd = compare_formatting(
                web[0], web[1], psd[0], psd[1], ignore_case,
                check_bold_italic=check_bold_italic,
                check_superscript=check_superscript,
            )
            if format_status == 'MISMATCH':
                format_status_display = _format_diff_label(mismatches)
                parts = []
                for m in mismatches[:6]:
                    parts.append(
                        f"{m['char']!r} web={_style_label(*m['website'], check_bold_italic, check_superscript)} "
                        f"psd={_style_label(*m['psd'], check_bold_italic, check_superscript)}"
                    )
                format_details = '; '.join(parts)
                if len(mismatches) > 6:
                    format_details += f' ... (+{len(mismatches) - 6} more)'
            elif extra_in_psd:
                # PSD has bold/italic the website doesn't -- not a
                # mismatch, just a note on an otherwise-matching row.
                format_status_display = f"MATCH (PSD only: {_extra_psd_label(extra_in_psd)})"
                format_details = f"PSD has {_extra_psd_label(extra_in_psd)} not present on the website (not an error)."
            else:
                format_status_display = format_status

        rows.append({
            'code': pair_label,
            'status': status,
            'format_status': format_status,
            'format_status_display': format_status_display,
            'format_details': format_details,
            'details': note or format_details,
            'website_text': web[0] if web else '',
            'psd_text': psd[0] if psd else '',
            'website_file': website_sources.get(web_code, ''),
            'psd_file': psd_sources.get(psd_code, ''),
        })

    return rows


def checker_scan_folder(folder):
    folder = Path(folder)
    html_files = [p for p in folder.rglob('*') if p.suffix.lower() in CHECKER_HTML_EXTS]
    psd_files = [p for p in folder.rglob('*') if p.suffix.lower() in CHECKER_PSD_EXTS]
    return sorted(html_files), sorted(psd_files)


CHECKER_ROW_COLORS = {
    'MATCH': '#d4edda',
    'MISMATCH': '#f8d7da',
    'MISSING_IN_PSD': '#fff3cd',
    'MISSING_ON_WEBSITE': '#fff3cd',
    'FORMAT_MISMATCH': '#ffe0b3',
    'CONVERTED': '#d1ecf1',
}

STATUS_LABELS = {
    'MATCH': 'Match',
    'MISMATCH': 'Text mismatch',
    'MISSING_IN_PSD': 'Missing in PSD',
    'MISSING_ON_WEBSITE': 'Missing on website',
    'FORMAT_MISMATCH': 'Formatting mismatch',
    'CONVERTED': 'Converted to smart object',
}


def _checker_row_tag(row):
    """Content problems take priority; a format-only mismatch gets its
    own distinct color so it's not confused with a text mismatch."""
    if row['status'] != 'MATCH':
        return row['status']
    if row['format_status'] == 'MISMATCH':
        return 'FORMAT_MISMATCH'
    return 'MATCH'


class CheckerTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, style="Card.TFrame")

        self.folder = tk.StringVar(value="No folder selected")
        self.ignore_case = tk.BooleanVar(value=False)
        self.ignore_whitespace = tk.BooleanVar(value=False)
        self.ignore_linebreaks = tk.BooleanVar(value=True)
        self.check_bold_italic = tk.BooleanVar(value=False)
        self.check_superscript = tk.BooleanVar(value=False)
        self.match_by_position = tk.BooleanVar(value=False)
        self.rows = []
        self._html_files = []
        self._psd_files = []
        self._status_items = {}   # status -> [treeview item ids], for jump-to-next
        self._jump_index = {}     # status -> last-jumped-to index in that list
        self.chip_labels = {}     # status -> chip Label widget

        self._build_ui()

    def _build_ui(self):
        missing = []
        if BeautifulSoup is None:
            missing.append('beautifulsoup4')
        if PSDImage is None:
            missing.append('psd-tools')
        if missing:
            ttk.Label(
                self, style="Card.TLabel", foreground='#B3261E', padding=10,
                text="Missing required package(s): %s\n"
                     "(this only happens running the .py directly without them installed -- "
                     "the built .exe already has these bundled in)" % ", ".join(missing)
            ).pack(fill='x')

        top = ttk.Frame(self, style="Card.TFrame", padding=(14, 14, 14, 6))
        top.pack(fill='x')
        ttk.Button(top, text="Choose Folder...", style="Primary.TButton",
                   command=self.choose_folder).pack(side='left')
        ttk.Label(top, textvariable=self.folder, style="CardMuted.TLabel").pack(side='left', padx=10)

        opts = ttk.Frame(self, style="Card.TFrame", padding=(14, 4))
        opts.pack(fill='x')
        ttk.Checkbutton(opts, text="Ignore case", variable=self.ignore_case,
                         style="Card.TCheckbutton").pack(side='left')
        ttk.Checkbutton(opts, text="Ignore line whitespace", style="Card.TCheckbutton",
                         variable=self.ignore_whitespace).pack(side='left', padx=10)
        ttk.Checkbutton(opts, text="Ignore line breaks (wrapped differently is OK)", style="Card.TCheckbutton",
                         variable=self.ignore_linebreaks).pack(side='left', padx=10)
        ttk.Checkbutton(opts, text="Check bold/italic", style="Card.TCheckbutton",
                         variable=self.check_bold_italic).pack(side='left', padx=10)
        ttk.Checkbutton(opts, text="Check superscript", style="Card.TCheckbutton",
                         variable=self.check_superscript).pack(side='left', padx=10)

        opts2 = ttk.Frame(self, style="Card.TFrame", padding=(14, 0, 14, 4))
        opts2.pack(fill='x')
        ttk.Checkbutton(
            opts2,
            text="Match by posion - Safari only",
            style="Card.TCheckbutton",
            variable=self.match_by_position,
        ).pack(side='left')
        ttk.Label(
            opts2,
            text="  Vẫn còn nhiều lỗi chưa nghiên cứu xong",
            style="CardMuted.TLabel",
        ).pack(side='left')

        btns = ttk.Frame(self, style="Card.TFrame", padding=14)
        btns.pack(fill='x')
        self.run_btn = ttk.Button(btns, text="Run Check", style="Primary.TButton",
                                   command=self.run_check, state='disabled')
        self.run_btn.pack(side='left')
        self.save_btn = ttk.Button(btns, text="Save CSV Report...", command=self.save_csv, state='disabled')
        self.save_btn.pack(side='left', padx=10)

        self.summary = ttk.Label(
            self, text="", style="CardMuted.TLabel", padding=(14, 0)
        )
        self.summary.pack(fill='x')

        legend = ttk.Frame(self, style="Card.TFrame", padding=(14, 6))
        legend.pack(fill='x')
        ttk.Label(legend, text="Click a count to jump to it:", style="CardMuted.TLabel").pack(side='left', padx=(0, 6))
        for key in ('MISMATCH', 'MISSING_IN_PSD', 'MISSING_ON_WEBSITE', 'CONVERTED', 'FORMAT_MISMATCH', 'MATCH'):
            chip = tk.Label(
                legend, text=f"  {STATUS_LABELS[key]}: 0  ",
                background=CHECKER_ROW_COLORS[key],
                foreground=COLORS["text"], font=("Segoe UI", 9),
                relief='flat', padx=4, pady=2, cursor='hand2',
            )
            chip.pack(side='left', padx=3)
            chip.bind('<Button-1>', lambda e, k=key: self._jump_to_next(k))
            self.chip_labels[key] = chip

        columns = ('code', 'status', 'format_status', 'psd_file',
                   'website_text', 'psd_text')
        self.tree = ttk.Treeview(self, columns=columns, show='headings')
        headers = {
            'code': 'Code',
            'status': 'Text Status',
            'format_status': 'Format Status',
            'psd_file': 'PSD File',
            'website_text': 'Website Text',
            'psd_text': 'PSD Text',
        }
        widths = {
            'code': 55,
            'status': 110,
            'format_status': 100,
            'psd_file': 130,
            'website_text': 220,
            'psd_text': 220,
        }
        for col in columns:
            self.tree.heading(col, text=headers[col])
            self.tree.column(col, width=widths[col],
                              anchor='center' if col in ('code', 'status', 'format_status') else 'w')

        vsb = ttk.Scrollbar(self, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side='left', fill='both', expand=True, padx=(10, 0), pady=10)
        vsb.pack(side='right', fill='y', pady=10)

        for status, color in CHECKER_ROW_COLORS.items():
            self.tree.tag_configure(status, background=color)

    def _jump_to_next(self, status):
        """Selects and scrolls to the next row tagged with `status`,
        cycling back to the first one once the end is reached."""
        items = self._status_items.get(status) or []
        if not items:
            return
        idx = self._jump_index.get(status, -1) + 1
        if idx >= len(items):
            idx = 0
        self._jump_index[status] = idx
        item = items[idx]
        self.tree.selection_set(item)
        self.tree.focus(item)
        self.tree.see(item)

    def choose_folder(self):
        folder = filedialog.askdirectory(title="Select folder with your .mhtml and .psd/.psb files")
        if not folder:
            return
        self.folder.set(folder)
        html_files, psd_files = checker_scan_folder(folder)
        if not html_files:
            messagebox.showwarning("No website file found",
                                    "No .mhtml/.html file was found in that folder.")
        if not psd_files:
            messagebox.showwarning("No PSD file found",
                                    "No .psd/.psb file was found in that folder.")
        self._html_files = html_files
        self._psd_files = psd_files
        self.run_btn.config(state='normal' if (html_files and psd_files) else 'disabled')
        self.summary.config(
            text=f"Found {len(html_files)} website file(s), {len(psd_files)} PSD file(s)."
        )

    def run_check(self):
        missing = []
        if BeautifulSoup is None:
            missing.append('beautifulsoup4')
        if PSDImage is None:
            missing.append('psd-tools')
        if missing:
            messagebox.showerror(
                "Missing dependency",
                "This needs: %s\nInstall with:\n    pip install %s"
                % (", ".join(missing), " ".join(missing))
            )
            return

        self.run_btn.config(state='disabled')
        self.summary.config(text="Checking...")
        threading.Thread(target=self._run_check_thread, daemon=True).start()

    def _run_check_thread(self):
        try:
            website_boxes = {}
            website_sources = {}
            for f in self._html_files:
                for code, val in extract_website_boxes(f).items():
                    if code not in website_boxes:
                        website_boxes[code] = val
                        website_sources[code] = f.name

            psd_layers = {}
            psd_sources = {}
            smart_object_codes = set()
            psd_positions = {}
            # Files are handled in sorted-filename order (see
            # checker_scan_folder), which for a multi-page chapter usually
            # matches page order -- so we offset each file's own top-pixel
            # positions by a running total to keep later pages sorting
            # after earlier ones even though each PSD's own coordinates
            # restart at 0.
            running_offset = 0
            for f in self._psd_files:
                layers, so_codes, tops = extract_psd_layers(f)
                file_max_bottom = 0
                for code, val in layers.items():
                    if code not in psd_layers:
                        psd_layers[code] = val
                        psd_sources[code] = f.name
                    if code in tops:
                        psd_positions.setdefault(code, running_offset + tops[code])
                        file_max_bottom = max(file_max_bottom, tops[code])
                for code in so_codes:
                    if code not in psd_layers and code not in smart_object_codes:
                        smart_object_codes.add(code)
                        psd_sources.setdefault(code, f.name)
                        if code in tops:
                            psd_positions.setdefault(code, running_offset + tops[code])
                            file_max_bottom = max(file_max_bottom, tops[code])
                running_offset += file_max_bottom + 1

            compare_fn = compare_by_position if self.match_by_position.get() else compare
            extra_kwargs = {'psd_positions': psd_positions} if self.match_by_position.get() else {}
            rows = compare_fn(
                website_boxes, psd_layers,
                website_sources=website_sources,
                psd_sources=psd_sources,
                smart_object_codes=smart_object_codes,
                ignore_case=self.ignore_case.get(),
                ignore_whitespace=self.ignore_whitespace.get(),
                ignore_linebreaks=self.ignore_linebreaks.get(),
                check_bold_italic=self.check_bold_italic.get(),
                check_superscript=self.check_superscript.get(),
                **extra_kwargs
            )
        except Exception as e:
            traceback.print_exc()
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
            self.after(0, lambda: self.run_btn.config(state='normal'))
            return

        self.after(0, lambda: self._show_results(rows))

    def _show_results(self, rows):
        self.rows = rows
        for item in self.tree.get_children():
            self.tree.delete(item)

        self._status_items = {}
        self._jump_index = {}

        counts = {'MATCH': 0, 'MISMATCH': 0, 'MISSING_IN_PSD': 0,
                  'MISSING_ON_WEBSITE': 0, 'FORMAT_MISMATCH': 0, 'CONVERTED': 0}
        for r in rows:
            tag = _checker_row_tag(r)
            counts[tag] = counts.get(tag, 0) + 1
            web_preview = r['website_text'].replace('\n', ' / ')
            psd_preview = r['psd_text'].replace('\n', ' / ')
            item_id = self.tree.insert(
                '', 'end',
                values=(r['code'], r['status'], r['format_status_display'],
                        r['psd_file'],
                        web_preview, psd_preview),
                tags=(tag,)
            )
            self._status_items.setdefault(tag, []).append(item_id)

        for key, chip in self.chip_labels.items():
            chip.config(text=f"  {STATUS_LABELS[key]}: {counts.get(key, 0)}  ")

        summary_text = "  |  ".join(f"{STATUS_LABELS[k]}: {v}" for k, v in counts.items())
        self.summary.config(text=summary_text)
        self.run_btn.config(state='normal')
        self.save_btn.config(state='normal' if rows else 'disabled')

    def save_csv(self):
        if not self.rows:
            return
        path = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[('CSV files', '*.csv')],
            initialfile='report.csv',
        )
        if not path:
            return
        fieldnames = ['code', 'status', 'format_status', 'format_status_display', 'website_file', 'psd_file',
                      'website_text', 'psd_text', 'details']
        with open(path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
            writer.writeheader()
            writer.writerows(self.rows)
        messagebox.showinfo("Saved", f"Report saved to:\n{path}")


# ══════════════════════════════════════════════════════════════════
#  TOP-LEVEL WINDOW: combines both tools as tabs in one window/exe
# ══════════════════════════════════════════════════════════════════

class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Add & check type")
        self.geometry("900x650")
        self.minsize(760, 480)

        # On Windows, unless the process tells the shell it's its own
        # distinct app, the TASKBAR icon specifically can keep showing
        # Python's own icon (grouped under python.exe/pythonw.exe) even
        # after the in-window titlebar icon below is set correctly. This
        # has no effect when running as a frozen PyInstaller .exe (which
        # is already its own process), only when running the .py
        # directly with Python installed.
        if sys.platform == "win32":
            try:
                import ctypes
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                    "PSDTools.BatchPlacerAndChecker"
                )
            except Exception as e:
                print(f"[icon] couldn't set AppUserModelID (taskbar grouping): {e}")

        # Use the same icon as the built app's file thumbnail, instead of
        # Tk's default feather/quill icon. Windows/Linux Tk can load a
        # .ico via iconbitmap(); macOS Tk cannot load .ico at all (it's
        # not a supported format there), so a .icns is used instead if
        # one is bundled. Any failure here is printed (not silently
        # swallowed) so it's obvious *why* the default icon is still
        # showing if this doesn't work.
        icon_filename = "icon.icns" if sys.platform == "darwin" else "icon.ico"
        icon_path = resource_path(icon_filename)
        if not os.path.isfile(icon_path):
            print(f"[icon] {icon_filename} not found at: {icon_path} "
                  f"-- falling back to the default Tk icon.")
        else:
            try:
                if sys.platform == "darwin":
                    # Tk on macOS has no iconbitmap() equivalent for a
                    # running app's Dock icon (that comes from the .app
                    # bundle's Info.plist / --icon at build time
                    # instead) -- there's simply nothing more to do here
                    # once the file exists, so this branch is a no-op by
                    # design, not a missing feature.
                    pass
                else:
                    self.iconbitmap(default=icon_path)
            except tk.TclError as e:
                print(f"[icon] found {icon_filename} at {icon_path} but Tk "
                      f"couldn't load it ({e}) -- make sure it's a valid "
                      f"multi-size .ico file, not a renamed .png.")

        setup_style(self)

        # Outer light-gray canvas with a little breathing room around the
        # white "card" that holds the actual tabs/content.
        outer = ttk.Frame(self, style="App.TFrame", padding=16)
        outer.pack(fill='both', expand=True)

        # ── Update status bar: small, tucked in the bottom-right corner ──
        # Packed BEFORE the notebook (with side='bottom') so it reserves
        # its own thin strip at the very bottom; the notebook below then
        # fills whatever space remains.
        update_bar = ttk.Frame(outer, style="App.TFrame")
        update_bar.pack(side='bottom', fill='x', pady=(8, 0))

        self.update_status_label = ttk.Label(
            update_bar, text=f"v{APP_VERSION}",
            foreground=COLORS["text_muted"], background=COLORS["bg"],
            font=("Segoe UI", 8),
        )
        self.update_status_label.pack(side='right', padx=(0, 4))

        self.update_btn = tk.Button(
            update_bar, text="Update", font=("Segoe UI", 8),
            bg=COLORS["primary"], fg=COLORS["text_on_blue"],
            activebackground=COLORS["primary_dk"], activeforeground=COLORS["text_on_blue"],
            relief='flat', bd=0, padx=8, pady=1, cursor="hand2",
            command=self._on_update_button_click,
        )
        # Hidden until an update is actually found (see _on_update_found).
        self._update_info = None

        notebook = ttk.Notebook(outer)
        notebook.pack(fill='both', expand=True)

        self.placer_tab = PlacerTab(notebook)
        self.checker_tab = CheckerTab(notebook)

        notebook.add(self.placer_tab, text="Batch Text Placer")
        notebook.add(self.checker_tab, text="PSD ↔ Website Checker")

        # Kiem tra update ngam ngay khi vua mo app, khong lam dung UI.
        self._check_update_on_startup()

    # ── Auto-update: check on startup, small corner UI, update on click ──

    def _check_update_on_startup(self):
        if check_for_update is None:
            # updater.py chua duoc bundle (vd chay .py truc tiep ma
            # thieu file) -- bo qua nhe nhang, khong lam phien nguoi dung.
            return

        def worker():
            try:
                info = check_for_update(current_version=APP_VERSION)
            except Exception as e:
                print(f"[update] check that bai: {e}")
                return
            if info:
                self.after(0, lambda: self._on_update_found(info))

        threading.Thread(target=worker, daemon=True).start()

    def _on_update_found(self, info):
        self._update_info = info
        self.update_status_label.config(text=f"v{APP_VERSION} → v{info['version']} có sẵn")
        # Chi hien nut luc nay co ban moi thuc su, tranh lam roi UI luc binh thuong.
        self.update_btn.pack(side='right', padx=(0, 6))

    def _on_update_button_click(self):
        if not self._update_info or download_update is None or apply_update is None:
            return

        self.update_btn.config(state='disabled', text="Đang tải...")
        self.update_status_label.config(text="Đang tải bản cập nhật...")

        def worker():
            try:
                new_file = download_update(self._update_info["download_url"])
            except Exception as e:
                print(f"[update] tai that bai: {e}")
                self.after(0, lambda: self._on_update_failed())
                return
            self.after(0, lambda: self._on_update_downloaded(new_file))

        threading.Thread(target=worker, daemon=True).start()

    def _on_update_downloaded(self, new_file):
        self.update_status_label.config(text="Đang cài đặt, ứng dụng sẽ khởi động lại...")
        # apply_update() thoat app o cuoi ham -- script phu ben ngoai se
        # thay the file va mo lai app ban moi.
        apply_update(new_file)

    def _on_update_failed(self):
        self.update_btn.config(state='normal', text="Update")
        self.update_status_label.config(text="Cập nhật thất bại, thử lại sau")


def main():
    if sys.platform != "darwin" and win32com is None:
        print("Batch Text Placer needs pywin32. Install it with:\n    pip install pywin32")
    if BeautifulSoup is None or PSDImage is None:
        print("PSD <-> Website Checker needs beautifulsoup4 and psd-tools. "
              "Install with:\n    pip install beautifulsoup4 psd-tools")
    # Still show the window either way -- each tab shows its own clear error
    # only when you actually try to use a feature that's missing a dependency,
    # rather than refusing to open the whole app.
    app = MainApp()
    app.mainloop()


if __name__ == "__main__":
    main()
