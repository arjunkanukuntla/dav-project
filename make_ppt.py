# -*- coding: utf-8 -*-
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

"""
TG-SPDCL Electricity Consumption & Monitoring
Professional Presentation Generator — Redesigned
Authors: D.Chander, K.Arjun, P.Sachin
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from lxml import etree
from pptx.enum.shapes import MSO_CONNECTOR_TYPE
import os, copy

# ─── DESIGN TOKENS ─────────────────────────────────────────────────────────────
# Professional "energy consulting" palette — clean, not AI-dark-mode
NAVY       = RGBColor(0x0d, 0x1b, 0x3e)   # deep navy (headers, accents)
NAVY_MED   = RGBColor(0x1a, 0x31, 0x6b)   # medium navy
TEAL       = RGBColor(0x00, 0xb4, 0xd8)   # electric teal
TEAL_DARK  = RGBColor(0x02, 0x7a, 0x9b)   # teal dark
GOLD       = RGBColor(0xf5, 0xa6, 0x23)   # warm gold accent
GOLD_DARK  = RGBColor(0xc7, 0x83, 0x00)   # darker gold
GREEN      = RGBColor(0x2d, 0xb8, 0x7d)   # success green
RED        = RGBColor(0xe0, 0x3e, 0x3e)   # alert red
SLATE      = RGBColor(0x64, 0x74, 0x8b)   # slate text
LIGHT_BG   = RGBColor(0xf4, 0xf6, 0xf9)   # near white background
WHITE      = RGBColor(0xff, 0xff, 0xff)
DARK_TEXT  = RGBColor(0x1a, 0x20, 0x2e)   # almost black text
MID_GRAY   = RGBColor(0x8a, 0x97, 0xa8)
CARD_BG    = RGBColor(0xf0, 0xf4, 0xfa)   # very light blue-white

SW = Inches(13.33)
SH = Inches(7.5)
MSO_CONNECTOR_STRAIGHT = MSO_CONNECTOR_TYPE.STRAIGHT


# ─── UTILITY FUNCTIONS ─────────────────────────────────────────────────────────

def set_bg(slide, color=LIGHT_BG):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def rect(slide, l, t, w, h, fill, line=None, lw=Pt(0), radius=None):
    """Add a rectangle, optionally with rounded corners."""
    if radius:
        shp = slide.shapes.add_shape(5, l, t, w, h)  # 5 = rounded rectangle
        shp.adjustments[0] = radius
    else:
        shp = slide.shapes.add_shape(1, l, t, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line:
        shp.line.color.rgb = line
        shp.line.width = lw
    else:
        shp.line.fill.background()
    return shp

def txt(slide, text, l, t, w, h, size=Pt(14), bold=False, color=DARK_TEXT,
        align=PP_ALIGN.LEFT, italic=False, name="Calibri"):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    para = tf.paragraphs[0]
    para.alignment = align
    run = para.add_run()
    run.text = text
    run.font.size = size
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = name
    return box

def multiline_tb(slide, lines, l, t, w, h, size=Pt(13), color=DARK_TEXT,
                 bold=False, name="Calibri", spacing=Pt(6)):
    """Textbox with multiple bullet lines."""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        if i == 0:
            para = tf.paragraphs[0]
        else:
            para = tf.add_paragraph()
        para.space_before = spacing
        run = para.add_run()
        run.text = line
        run.font.size = size
        run.font.bold = bold
        run.font.color.rgb = color
        run.font.name = name
    return box

def header_bar(slide, title, subtitle=None, bar_color=NAVY):
    """Standard page header — navy bar + title."""
    # Top bar
    rect(slide, 0, 0, SW, Inches(1.35), bar_color)
    # Accent stripe
    rect(slide, 0, Inches(1.35), SW, Inches(0.045), GOLD)
    # Title
    txt(slide, title, Inches(0.5), Inches(0.2), Inches(11.0), Inches(0.75),
        size=Pt(30), bold=True, color=WHITE, name="Calibri Light")
    if subtitle:
        txt(slide, subtitle, Inches(0.5), Inches(0.9), Inches(11.0), Inches(0.42),
            size=Pt(14), color=RGBColor(0xb0, 0xc8, 0xe8), italic=True, name="Calibri")
    # Slide watermark right
    txt(slide, "TG-SPDCL  |  Data Analytics", Inches(8.5), Inches(0.92),
        Inches(4.6), Inches(0.38), size=Pt(11), color=RGBColor(0x7a, 0xa0, 0xcc),
        align=PP_ALIGN.RIGHT, name="Calibri")

def footer(slide, num, total=17):
    rect(slide, 0, SH - Inches(0.38), SW, Inches(0.38), NAVY)
    txt(slide, "TG-SPDCL Domestic Electricity Consumption Analysis  |  July 2026",
        Inches(0.4), SH - Inches(0.34), Inches(10.0), Inches(0.32),
        size=Pt(10), color=RGBColor(0x90, 0xb4, 0xd8), name="Calibri")
    txt(slide, f"{num} / {total}", Inches(12.5), SH - Inches(0.34), Inches(0.7), Inches(0.32),
        size=Pt(10), color=WHITE, align=PP_ALIGN.RIGHT, name="Calibri")

def divider(slide, y=Inches(1.38), color=GOLD, h=Inches(0.045)):
    rect(slide, 0, y, SW, h, color)

def stat_card(slide, l, t, w, h, label, value, sub, accent=TEAL, bg=WHITE):
    """Clean KPI card with left accent stripe."""
    rect(slide, l, t, w, h, bg, TEAL, Pt(0.5))
    rect(slide, l, t, Inches(0.08), h, accent)   # left accent stripe
    txt(slide, label.upper(), l + Inches(0.18), t + Inches(0.12), w - Inches(0.25), Inches(0.3),
        size=Pt(9.5), bold=False, color=SLATE, name="Calibri")
    txt(slide, value, l + Inches(0.18), t + Inches(0.38), w - Inches(0.25), Inches(0.65),
        size=Pt(24), bold=True, color=NAVY, name="Calibri Light")
    txt(slide, sub, l + Inches(0.18), t + h - Inches(0.42), w - Inches(0.25), Inches(0.35),
        size=Pt(10), color=MID_GRAY, name="Calibri")

def flow_box(slide, l, t, w, h, label, sub="", fill=NAVY, text_color=WHITE, small=False):
    """Box for flowcharts."""
    shp = rect(slide, l, t, w, h, fill, line=fill, lw=Pt(0), radius=0.08)
    font_size = Pt(11) if small else Pt(13)
    txt(slide, label, l + Inches(0.1), t + Inches(0.1) if not sub else t + Inches(0.07),
        w - Inches(0.2), h - Inches(0.25), size=font_size, bold=True,
        color=text_color, align=PP_ALIGN.CENTER, name="Calibri")
    if sub:
        txt(slide, sub, l + Inches(0.08), t + h - Inches(0.33),
            w - Inches(0.16), Inches(0.3), size=Pt(9), color=RGBColor(0xcc, 0xdd, 0xee),
            align=PP_ALIGN.CENTER, name="Calibri")
    return shp

def arrow_h(slide, l, t, length, color=SLATE):
    """Horizontal arrow."""
    cxn = slide.shapes.add_connector(MSO_CONNECTOR_STRAIGHT, l, t, l + length, t)
    cxn.line.color.rgb = color
    cxn.line.width = Pt(1.8)

def arrow_v(slide, l, t, length, color=SLATE):
    """Vertical arrow."""
    cxn = slide.shapes.add_connector(MSO_CONNECTOR_STRAIGHT, l, t, l, t + length)
    cxn.line.color.rgb = color
    cxn.line.width = Pt(1.8)

def diamond(slide, cx, cy, half_w, half_h, fill, label, text_color=WHITE):
    """Diamond shape for decision nodes."""
    shp = slide.shapes.add_shape(4, cx - half_w, cy - half_h, half_w*2, half_h*2)  # 4 = diamond
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    tf = shp.text_frame
    tf.text = label
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    run = tf.paragraphs[0].runs[0]
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = text_color
    run.font.name = "Calibri"
    return shp

def bullet_list(slide, items, l, t, w, size=Pt(13.5), color=DARK_TEXT,
                marker="›", spacing=Pt(4)):
    box = slide.shapes.add_textbox(l, t, w, Inches(6))
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        if i == 0:
            para = tf.paragraphs[0]
        else:
            para = tf.add_paragraph()
        para.space_before = spacing
        para.space_after = spacing
        run = para.add_run()
        if isinstance(item, tuple):
            run.text = f"{marker}  {item[0]}"
            run.font.bold = True
            run.font.size = size
            run.font.color.rgb = NAVY
            run.font.name = "Calibri"
            # subtext
            sub_para = tf.add_paragraph()
            sub_para.space_before = Pt(1)
            sub_para.space_after = Pt(6)
            sub_run = sub_para.add_run()
            sub_run.text = f"      {item[1]}"
            sub_run.font.size = Pt(11.5)
            sub_run.font.color.rgb = SLATE
            sub_run.font.name = "Calibri"
        else:
            run.text = f"{marker}  {item}"
            run.font.size = size
            run.font.color.rgb = color
            run.font.name = "Calibri"
    return box


# ─── BUILD PRESENTATION ────────────────────────────────────────────────────────

def build():
    prs = Presentation()
    prs.slide_width  = SW
    prs.slide_height = SH
    BLANK = prs.slide_layouts[6]

    # ══════════════════════════════════════════════════════════════════════════════
    # SLIDE 1 — TITLE
    # ══════════════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, NAVY)

    # Large diagonal accent shape (visual interest — bottom right)
    r = rect(sl, Inches(8.8), Inches(3.5), Inches(5.2), Inches(4.5), NAVY_MED)
    r2 = rect(sl, Inches(9.8), Inches(4.5), Inches(4.0), Inches(3.5), RGBColor(0x0a, 0x14, 0x2e))

    # Gold horizontal divider accent
    rect(sl, Inches(0.55), Inches(3.5), Inches(5.5), Inches(0.055), GOLD)

    # Category tag
    tag = rect(sl, Inches(0.55), Inches(1.1), Inches(2.8), Inches(0.42), TEAL, radius=0.5)
    txt(sl, "DATA ANALYTICS PROJECT", Inches(0.65), Inches(1.12), Inches(2.65), Inches(0.38),
        size=Pt(10), bold=True, color=WHITE, name="Calibri")

    # Main title lines
    txt(sl, "Electricity Consumption",
        Inches(0.55), Inches(1.65), Inches(8.0), Inches(1.1),
        size=Pt(46), bold=True, color=WHITE, name="Calibri Light")
    txt(sl, "& Monitoring Analysis",
        Inches(0.55), Inches(2.6), Inches(8.0), Inches(0.9),
        size=Pt(46), bold=True, color=TEAL, name="Calibri Light")

    # Subtitle
    txt(sl, "TG-SPDCL  |  Domestic Consumer Dataset  |  July 2026",
        Inches(0.55), Inches(3.65), Inches(8.0), Inches(0.48),
        size=Pt(15), color=RGBColor(0x90, 0xb4, 0xd8), name="Calibri")

    # Team card
    rect(sl, Inches(0.55), Inches(4.4), Inches(5.8), Inches(2.65), RGBColor(0x12, 0x25, 0x52),
         TEAL, Pt(1))
    rect(sl, Inches(0.55), Inches(4.4), Inches(5.8), Inches(0.06), TEAL)
    txt(sl, "Team Members",
        Inches(0.75), Inches(4.5), Inches(5.4), Inches(0.38),
        size=Pt(11), bold=True, color=TEAL, name="Calibri")
    for i, (name_, roll) in enumerate([
        ("D. Chander", "Roll: 2503A52214"),
        ("K. Arjun",   "Roll: 2503A52235"),
        ("P. Sachin",  "Roll: 2503A52240"),
    ]):
        y_off = Inches(4.95) + i * Inches(0.58)
        txt(sl, name_, Inches(0.75), y_off, Inches(2.8), Inches(0.45),
            size=Pt(15), bold=True, color=WHITE, name="Calibri")
        txt(sl, roll, Inches(3.7), y_off + Inches(0.05), Inches(2.4), Inches(0.38),
            size=Pt(11), color=GOLD, name="Calibri")

    # Stats strip (right side)
    stats = [
        ("15,947", "Sector Records", TEAL),
        ("20",     "Circles",        GOLD),
        ("1,245",  "GWh Consumed",  GREEN),
        ("92.68%", "Billing Rate",   RGBColor(0xe0,0x8c,0x3e)),
    ]
    for i, (val, lbl, ac) in enumerate(stats):
        col = i % 2; row = i // 2
        cx = Inches(9.5) + col * Inches(1.95)
        cy = Inches(2.3) + row * Inches(2.0)
        rect(sl, cx, cy, Inches(1.75), Inches(1.65), RGBColor(0x12,0x25,0x52), ac, Pt(1.5))
        rect(sl, cx, cy, Inches(1.75), Inches(0.055), ac)
        txt(sl, val,  cx+Inches(0.1), cy+Inches(0.15), Inches(1.6), Inches(0.85),
            size=Pt(30), bold=True, color=ac, align=PP_ALIGN.CENTER, name="Calibri Light")
        txt(sl, lbl,  cx+Inches(0.1), cy+Inches(1.22), Inches(1.6), Inches(0.35),
            size=Pt(10.5), color=MID_GRAY, align=PP_ALIGN.CENTER, name="Calibri")

    footer(sl, 1)


    # ══════════════════════════════════════════════════════════════════════════════
    # SLIDE 2 — TABLE OF CONTENTS (visual numbered grid)
    # ══════════════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, LIGHT_BG)
    header_bar(sl, "Table of Contents", "What we cover in this presentation")
    footer(sl, 2)

    items = [
        ("01", "Dataset Overview",         "Source, scope & structure",            TEAL),
        ("02", "Column Guide",             "Data types & variable definitions",     NAVY_MED),
        ("03", "Key Metrics (KPIs)",        "GWh, billing rate, load",              GOLD_DARK),
        ("04", "Circle Rankings",           "Top & bottom performing circles",      GREEN),
        ("05", "Analysis Pipeline",         "End-to-end data workflow",             TEAL_DARK),
        ("06", "Top 20 Areas",             "Highest demand neighborhoods",          NAVY),
        ("07", "Billing Efficiency",        "vs 95% benchmark target",             RGBColor(0xc0,0x50,0x20)),
        ("08", "Load vs Consumption",       "Correlation & regression",             RGBColor(0x4a,0x0a,0x7c)),
        ("09", "Statistical Distributions", "KDE histograms & spread",             TEAL),
        ("10", "Data Quality & Cleanup",    "Issues, outliers & fixes",            RED),
        ("11", "Key Insights",              "Patterns & analytical takeaways",     GOLD_DARK),
        ("12", "Limitations & Conclusion",  "Scope boundaries & summary",          SLATE),
    ]

    cols = 3; col_w = Inches(4.1); row_h = Inches(0.75)
    ox = Inches(0.3); oy = Inches(1.55)
    for i, (num, title_, sub_, color_) in enumerate(items):
        col = i % cols; row = i // cols
        x = ox + col * (col_w + Inches(0.22))
        y = oy + row * (row_h + Inches(0.12))
        rect(sl, x, y, col_w, row_h, WHITE, color_, Pt(0.5), radius=0.08)
        rect(sl, x, y, Inches(0.55), row_h, color_, radius=0.08)
        txt(sl, num, x + Inches(0.06), y + Inches(0.18), Inches(0.42), Inches(0.42),
            size=Pt(18), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri Light")
        txt(sl, title_, x + Inches(0.65), y + Inches(0.08), col_w - Inches(0.75), Inches(0.38),
            size=Pt(13), bold=True, color=NAVY, name="Calibri")
        txt(sl, sub_,  x + Inches(0.65), y + Inches(0.44), col_w - Inches(0.75), Inches(0.28),
            size=Pt(10.5), color=SLATE, name="Calibri")


    # ══════════════════════════════════════════════════════════════════════════════
    # SLIDE 3 — DATASET OVERVIEW (with visual dataset icon)
    # ══════════════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, LIGHT_BG)
    header_bar(sl, "Dataset Overview", "TG-SPDCL Domestic Electricity Consumption — July 2026")
    footer(sl, 3)

    # Left info panel
    rect(sl, Inches(0.3), Inches(1.5), Inches(5.5), Inches(5.6), WHITE, TEAL, Pt(0.5), radius=0.06)
    rect(sl, Inches(0.3), Inches(1.5), Inches(5.5), Inches(0.5), TEAL, radius=0.06)
    txt(sl, "About the Dataset", Inches(0.5), Inches(1.55), Inches(5.1), Inches(0.42),
        size=Pt(14), bold=True, color=WHITE, name="Calibri")

    info_rows = [
        ("Source",    "TG-SPDCL — Telangana Southern Power Distribution Co."),
        ("Period",    "July 2026 (single-month snapshot)"),
        ("Category",  "DOMESTIC consumers only"),
        ("Total Rows","15,947 sector-level records"),
        ("Columns",   "11 columns (6 categorical + 5 numerical)"),
        ("Circles",   "20 circles across Telangana"),
        ("Divisions", "49 divisions | 131 sub-divisions | 476 sections"),
        ("File Size",  "~1.4 MB CSV  |  Dataset link: TG-SPDCL Open Data"),
    ]
    for j, (lbl, val) in enumerate(info_rows):
        y = Inches(2.18) + j * Inches(0.58)
        bg = CARD_BG if j % 2 == 0 else WHITE
        rect(sl, Inches(0.35), y - Inches(0.04), Inches(5.4), Inches(0.52), bg)
        txt(sl, lbl,  Inches(0.5),  y, Inches(1.5),  Inches(0.42),
            size=Pt(11), bold=True, color=TEAL_DARK, name="Calibri")
        txt(sl, val,  Inches(1.95), y, Inches(3.75), Inches(0.42),
            size=Pt(11), color=DARK_TEXT, name="Calibri")

    # Right: visual grid showing dataset scope
    rect(sl, Inches(6.1), Inches(1.5), Inches(6.9), Inches(5.6), WHITE, NAVY, Pt(0.5), radius=0.06)
    rect(sl, Inches(6.1), Inches(1.5), Inches(6.9), Inches(0.5), NAVY, radius=0.06)
    txt(sl, "Geographic & Data Scope", Inches(6.3), Inches(1.55), Inches(6.5), Inches(0.42),
        size=Pt(14), bold=True, color=WHITE, name="Calibri")

    # Scope visual: nested region diagram as rectangles
    scope = [
        (NAVY_MED, "20 CIRCLES",            Inches(6.25), Inches(2.15), Inches(6.6), Inches(4.75)),
        (TEAL_DARK,"49 DIVISIONS",           Inches(6.55), Inches(2.55), Inches(5.9), Inches(3.95)),
        (RGBColor(0x03,0x8c,0x6c), "131 SUBDIVISIONS", Inches(6.9),  Inches(3.0),  Inches(5.2), Inches(3.0)),
        (GREEN,    "476 SECTIONS",           Inches(7.3),  Inches(3.5),  Inches(4.4), Inches(2.05)),
        (GOLD,     "15,947 RECORDS",         Inches(7.7),  Inches(4.0),  Inches(3.6), Inches(1.1)),
    ]
    for (c, lbl, x, y, w, h) in scope:
        rect(sl, x, y, w, h, c, radius=0.06)
        txt(sl, lbl, x + Inches(0.1), y + Inches(0.08), w - Inches(0.2), Inches(0.4),
            size=Pt(10.5), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")


    # ══════════════════════════════════════════════════════════════════════════════
    # SLIDE 4 — COLUMN DEFINITIONS (table style)
    # ══════════════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, LIGHT_BG)
    header_bar(sl, "Dataset Column Guide", "Variable descriptions, data types & roles")
    footer(sl, 4)

    columns_info = [
        ("CIRCLE",        "Object",   "Categorical", "Top-level geographic region (20 unique circles)"),
        ("DIVISION",      "Object",   "Categorical", "Sub-region within a circle (49 divisions)"),
        ("SUBDIVISION",   "Object",   "Categorical", "Subdivision within a division (131 total)"),
        ("SECTION",       "Object",   "Categorical", "Service section area (476 sections)"),
        ("AREA",          "Object",   "Categorical", "Specific neighborhood / locality (47 NULLs)"),
        ("CATDESC/CODE",  "Object/Int","Categorical", "Consumer category — DOMESTIC only (code=1)"),
        ("TOTSERVICES",   "Int64",    "Numerical",   "Total registered electricity connections"),
        ("BILLDSERVICES", "Float64",  "Numerical",   "Billed connections (247 NULLs = zero-unit areas)"),
        ("UNITS",         "Float64",  "Numerical",   "Monthly electricity consumed in kWh (15 negatives)"),
        ("LOAD",          "Float64",  "Numerical",   "Sanctioned connected load in kW (decimal values)"),
    ]

    # Table header
    hdr_y = Inches(1.55)
    for cx, cw, label, col_c in [
        (Inches(0.3),  Inches(1.8),  "COLUMN",     NAVY),
        (Inches(2.15), Inches(1.15), "DTYPE",      NAVY_MED),
        (Inches(3.35), Inches(1.4),  "TYPE",       TEAL_DARK),
        (Inches(4.8),  Inches(8.2),  "DESCRIPTION",NAVY),
    ]:
        rect(sl, cx, hdr_y, cw, Inches(0.42), col_c)
        txt(sl, label, cx + Inches(0.1), hdr_y + Inches(0.06), cw - Inches(0.15), Inches(0.35),
            size=Pt(11), bold=True, color=WHITE, name="Calibri")

    for i, (col, dtype, kind, desc) in enumerate(columns_info):
        y = Inches(1.97) + i * Inches(0.49)
        bg = CARD_BG if i % 2 == 0 else WHITE
        rect(sl, Inches(0.3),  y, Inches(12.7), Inches(0.48), bg)
        # Col name
        txt(sl, col,   Inches(0.4),  y + Inches(0.07), Inches(1.65), Inches(0.38),
            size=Pt(11), bold=True, color=NAVY, name="Calibri")
        # Dtype badge
        dtc = TEAL if dtype == "Object" else NAVY_MED
        rect(sl, Inches(2.18), y + Inches(0.09), Inches(1.1), Inches(0.3), dtc, radius=0.6)
        txt(sl, dtype, Inches(2.22), y + Inches(0.1), Inches(1.05), Inches(0.28),
            size=Pt(9), color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")
        # Kind badge
        kc = GOLD_DARK if kind == "Categorical" else GREEN
        rect(sl, Inches(3.38), y + Inches(0.09), Inches(1.3), Inches(0.3), kc, radius=0.6)
        txt(sl, kind,  Inches(3.42), y + Inches(0.1), Inches(1.25), Inches(0.28),
            size=Pt(9), color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")
        # Desc
        txt(sl, desc,  Inches(4.82), y + Inches(0.07), Inches(8.1), Inches(0.38),
            size=Pt(11), color=DARK_TEXT, name="Calibri")


    # ══════════════════════════════════════════════════════════════════════════════
    # SLIDE 5 — KPI DASHBOARD (8 stat cards)
    # ══════════════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, LIGHT_BG)
    header_bar(sl, "Key Performance Indicators", "Overall summary metrics for July 2026 domestic supply")
    footer(sl, 5)

    kpis = [
        ("Total Records",         "15,947",    "Sector-level area rows",              TEAL),
        ("Circles Covered",       "20",        "20 circles | 49 divisions",            NAVY_MED),
        ("Total Connections",     "8,867,827", "Registered domestic services",         GOLD_DARK),
        ("Billed Connections",    "8,218,867", "Connections that received a bill",     GREEN),
        ("Billing Efficiency",    "92.68%",    "Gap: ~649K unbilled  |  Target: 95%+", RED),
        ("Total Energy",          "1,245 GWh", "1,245,523,023 kWh total monthly",     TEAL_DARK),
        ("Sanctioned Load",       "13,779 MW", "13,778,816 kW connected capacity",    NAVY),
        ("Avg / Household",       "151.5 kWh", "Per billed connection per month",      GOLD_DARK),
    ]

    cw = Inches(3.05); ch = Inches(1.72)
    gx = Inches(0.15); gy = Inches(1.55)
    for i, (lbl, val, sub, ac) in enumerate(kpis):
        col = i % 4; row = i // 4
        x = gx + col * (cw + Inches(0.18))
        y = gy + row * (ch + Inches(0.14))
        stat_card(sl, x, y, cw, ch, lbl, val, sub, accent=ac)

    # Caption note below
    rect(sl, Inches(0.15), Inches(7.02), Inches(12.8), Inches(0.08), GOLD)
    txt(sl, "Note: Billing efficiency gap of 7.32% represents approximately 649,000 connections not billed — a key operational priority for TG-SPDCL.",
        Inches(0.3), Inches(7.1), Inches(12.7), Inches(0.32),
        size=Pt(10.5), color=SLATE, italic=True, name="Calibri")


    # ══════════════════════════════════════════════════════════════════════════════
    # SLIDE 6 — DATA PIPELINE FLOWCHART
    # ══════════════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, LIGHT_BG)
    header_bar(sl, "End-to-End Analysis Pipeline", "From raw CSV to interactive dashboard — data flow")
    footer(sl, 6)

    # Phase label helper
    def phase(x, y, label, color):
        rect(sl, x, y, Inches(1.5), Inches(0.32), color, radius=0.7)
        txt(sl, label, x + Inches(0.06), y + Inches(0.04), Inches(1.4), Inches(0.26),
            size=Pt(9), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")

    # ── Row 1: Ingestion → Cleaning
    box_y = Inches(1.65); bh = Inches(1.1)
    phases_r1 = [
        (Inches(0.3),  NAVY,     "RAW DATA",       "TG-SPDCL CSV\n15,947 rows",           TEAL),
        (Inches(2.9),  TEAL_DARK,"DATA CLEANING",   "Fill NULLs\nClip negatives",          TEAL),
        (Inches(5.5),  NAVY_MED, "FEATURE ENG.",    "Billing rate %\nUnits per service",   GOLD_DARK),
        (Inches(8.1),  GOLD_DARK,"AGGREGATION",     "Group by circle\nDivision pivots",    GOLD_DARK),
        (Inches(10.7), GREEN,    "EXPORT",          "analysis_summary\n.json",             GREEN),
    ]
    for x, ph_col, label, sub, _ in phases_r1:
        flow_box(sl, x, box_y, Inches(2.3), bh, label, sub, fill=NAVY)

    # Phase labels above boxes
    for i, (x, ph_col, *_) in enumerate(phases_r1):
        phase(x + Inches(0.4), box_y - Inches(0.42), ["INGEST","CLEAN","ENGINEER","AGGREGATE","OUTPUT"][i], ph_col)

    # Arrows between boxes
    for i in range(len(phases_r1) - 1):
        ax = phases_r1[i][0] + Inches(2.3)
        arrow_h(sl, ax, box_y + Inches(0.54), Inches(0.6), SLATE)

    # ── Row 2: Analysis outputs
    box2_y = Inches(3.55)
    rect(sl, Inches(0.25), box2_y - Inches(0.15), Inches(12.85), Inches(0.04), TEAL)
    txt(sl, "ANALYSIS OUTPUTS", Inches(0.3), box2_y, Inches(12.5), Inches(0.3),
        size=Pt(10), bold=True, color=TEAL_DARK, name="Calibri")

    outputs = [
        (Inches(0.3),  TEAL,     "circle_consumption.png", "Circle GWh & Load chart"),
        (Inches(2.9),  GOLD_DARK,"top20_areas.png",         "Top demand zones"),
        (Inches(5.5),  NAVY_MED, "billing_efficiency.png",  "Billing rate vs target"),
        (Inches(8.1),  GREEN,    "load_vs_units.png",       "Regression scatter"),
        (Inches(10.7), RGBColor(0x7a,0x1f,0xa0),"distribution_metrics.png","KDE histograms"),
    ]
    for x, c, lbl, sub in outputs:
        flow_box(sl, x, box2_y + Inches(0.42), Inches(2.3), Inches(0.98), lbl, sub, fill=c, small=True)

    # ── Row 3: Web Dashboard
    rect(sl, Inches(0.25), Inches(5.28), Inches(12.85), Inches(0.04), GOLD)
    txt(sl, "WEB DASHBOARD", Inches(0.3), Inches(5.35), Inches(3.0), Inches(0.3),
        size=Pt(10), bold=True, color=GOLD_DARK, name="Calibri")

    dash_items = [
        (Inches(0.3),  NAVY,     "KPI Cards",         "Live metric tiles"),
        (Inches(2.9),  TEAL_DARK,"Chart.js Graphs",   "Interactive visuals"),
        (Inches(5.5),  NAVY_MED, "Filter Slicers",    "Circle / Division"),
        (Inches(8.1),  GOLD_DARK,"Anomaly Alerts",    "<85% billing flags"),
        (Inches(10.7), GREEN,    "CSV Export",        "Paginated data table"),
    ]
    for x, c, lbl, sub in dash_items:
        flow_box(sl, x, Inches(5.75), Inches(2.3), Inches(0.98), lbl, sub, fill=c, small=True)

    # Vertical connectors from PNG outputs → dashboard row
    for x in [Inches(1.45), Inches(4.05), Inches(6.65), Inches(9.25), Inches(11.85)]:
        arrow_v(sl, x, box2_y + Inches(1.42), Inches(0.5), MID_GRAY)

    # Run command strip
    rect(sl, Inches(0.3), SH - Inches(0.85), Inches(12.7), Inches(0.44), DARK_TEXT, radius=0.05)
    txt(sl, "$ python main.py    →    $ python -m http.server 8000    →    http://localhost:8000",
        Inches(0.55), SH - Inches(0.82), Inches(12.2), Inches(0.4),
        size=Pt(11), color=TEAL, name="Courier New")


    # ══════════════════════════════════════════════════════════════════════════════
    # SLIDE 7 — CIRCLE RANKINGS (two-panel)
    # ══════════════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, LIGHT_BG)
    header_bar(sl, "Circle-Wise Consumption Rankings", "Highest and lowest consuming electricity circles")
    footer(sl, 7)

    # Left: Top 5
    rect(sl, Inches(0.3), Inches(1.52), Inches(6.0), Inches(5.6), WHITE, TEAL, Pt(1))
    rect(sl, Inches(0.3), Inches(1.52), Inches(6.0), Inches(0.52), TEAL)
    txt(sl, "TOP 5 CIRCLES — Highest Consumption", Inches(0.45), Inches(1.58),
        Inches(5.8), Inches(0.42), size=Pt(13), bold=True, color=WHITE, name="Calibri")

    top5 = [
        ("Medchal",           "141.7M units",  "11.4% share", 100),
        ("Habsiguda",         "138.1M units",  "11.1% share", 97),
        ("Cybercity",         "125.4M units",  "10.1% share", 88),
        ("Hyderabad Central", "112.5M units",  " 9.0% share", 79),
        ("Saroornagar",       "112.3M units",  " 9.0% share", 79),
    ]
    for i, (name_, units_, share_, pct) in enumerate(top5):
        y = Inches(2.2) + i * Inches(0.94)
        # Rank badge
        rect(sl, Inches(0.45), y, Inches(0.45), Inches(0.45), TEAL_DARK, radius=0.5)
        txt(sl, str(i+1), Inches(0.45), y + Inches(0.04), Inches(0.45), Inches(0.38),
            size=Pt(14), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")
        txt(sl, name_,  Inches(1.05), y,               Inches(3.4), Inches(0.3),
            size=Pt(13), bold=True, color=NAVY, name="Calibri")
        txt(sl, f"{units_}  |  {share_}", Inches(1.05), y + Inches(0.3), Inches(3.4), Inches(0.28),
            size=Pt(10.5), color=SLATE, name="Calibri")
        # Progress bar
        bar_w = Inches(4.4) * (pct / 100)
        rect(sl, Inches(1.05), y + Inches(0.6), Inches(4.4), Inches(0.14), CARD_BG)
        rect(sl, Inches(1.05), y + Inches(0.6), bar_w, Inches(0.14), TEAL)

    # Right: Bottom 5
    rect(sl, Inches(6.6), Inches(1.52), Inches(6.4), Inches(5.6), WHITE, RED, Pt(1))
    rect(sl, Inches(6.6), Inches(1.52), Inches(6.4), Inches(0.52), RED)
    txt(sl, "BOTTOM 5 CIRCLES — Lowest Consumption", Inches(6.75), Inches(1.58),
        Inches(6.2), Inches(0.42), size=Pt(13), bold=True, color=WHITE, name="Calibri")

    bot5 = [
        ("Gadwal",       "10.0M units",  "Smallest consumer"),
        ("Wanaparthy",   "10.1M units",  "Rural district"),
        ("Nagarkurnool", "13.1M units",  "District circle"),
        ("Medak",        "15.9M units",  "District circle"),
        ("Vikarabad",    "17.9M units",  "District circle"),
    ]
    for i, (name_, units_, note_) in enumerate(bot5):
        y = Inches(2.2) + i * Inches(0.94)
        rect(sl, Inches(6.75), y, Inches(0.45), Inches(0.45), RED, radius=0.5)
        txt(sl, str(i+1), Inches(6.75), y + Inches(0.04), Inches(0.45), Inches(0.38),
            size=Pt(14), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")
        txt(sl, name_,  Inches(7.35), y,               Inches(5.2), Inches(0.3),
            size=Pt(13), bold=True, color=NAVY, name="Calibri")
        txt(sl, f"{units_}  |  {note_}", Inches(7.35), y + Inches(0.3), Inches(5.2), Inches(0.28),
            size=Pt(10.5), color=SLATE, name="Calibri")
        pct = [7, 7.1, 9.2, 11.2, 12.6][i]
        rect(sl, Inches(7.35), y + Inches(0.6), Inches(5.2), Inches(0.14), CARD_BG)
        rect(sl, Inches(7.35), y + Inches(0.6), Inches(5.2) * (pct/15), Inches(0.14), RED)

    # Insight box
    rect(sl, Inches(0.3), Inches(7.02), Inches(12.7), Inches(0.06), GOLD)
    txt(sl, "Key Insight: Top 10 circles (all Hyderabad region) account for ~90% of all domestic units. District circles use ~1/3 per service compared to urban circles.",
        Inches(0.3), Inches(7.09), Inches(12.7), Inches(0.3),
        size=Pt(10.5), color=SLATE, italic=True, name="Calibri")


    # ══════════════════════════════════════════════════════════════════════════════
    # SLIDE 8 — CHART: CIRCLE CONSUMPTION
    # ══════════════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, LIGHT_BG)
    header_bar(sl, "Circle Consumption & Sanctioned Load", "Energy consumed (GWh) vs Connected load (MW) — all 20 circles")
    footer(sl, 8)
    if os.path.exists("circle_consumption.png"):
        sl.shapes.add_picture("circle_consumption.png", Inches(0.25), Inches(1.48), Inches(12.8), Inches(5.62))
    txt(sl, "Cyan bars = Energy consumption (GWh)  |  Purple line = Sanctioned load (MW)  |  Sorted: highest to lowest",
        Inches(0.4), Inches(7.1), Inches(12.5), Inches(0.3), size=Pt(10.5),
        color=SLATE, italic=True, align=PP_ALIGN.CENTER, name="Calibri")


    # ══════════════════════════════════════════════════════════════════════════════
    # SLIDE 9 — CHART: TOP 20 AREAS
    # ══════════════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, LIGHT_BG)
    header_bar(sl, "Top 20 Highest Demand Areas", "Largest electricity consuming neighborhoods — Seaborn barplot with circle hue")
    footer(sl, 9)

    # Key callouts beside chart
    if os.path.exists("top20_areas.png"):
        sl.shapes.add_picture("top20_areas.png", Inches(0.2), Inches(1.48), Inches(9.5), Inches(5.7))

    # Callout panel on right
    rect(sl, Inches(9.85), Inches(1.52), Inches(3.3), Inches(5.6), WHITE, GOLD_DARK, Pt(0.8))
    rect(sl, Inches(9.85), Inches(1.52), Inches(3.3), Inches(0.48), GOLD_DARK)
    txt(sl, "TOP 3 HIGHLIGHTS", Inches(10.0), Inches(1.58), Inches(3.0), Inches(0.38),
        size=Pt(12), bold=True, color=WHITE, name="Calibri")
    callouts = [
        ("01", "Hydernagar",   "2.32M kWh", "Largest single sector"),
        ("02", "Moosapet",     "1.81M kWh", "Borabanda division"),
        ("03", "Vinayaknagar", "1.47M kWh", "Medchal circle"),
    ]
    for i, (n, name_, val_, note_) in enumerate(callouts):
        y = Inches(2.15) + i * Inches(1.35)
        rect(sl, Inches(10.0), y, Inches(0.42), Inches(0.42), GOLD, radius=0.5)
        txt(sl, n, Inches(10.0), y + Inches(0.04), Inches(0.42), Inches(0.36),
            size=Pt(15), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")
        txt(sl, name_, Inches(10.5), y, Inches(2.5), Inches(0.35),
            size=Pt(13), bold=True, color=NAVY, name="Calibri")
        txt(sl, val_,  Inches(10.5), y + Inches(0.35), Inches(2.5), Inches(0.32),
            size=Pt(20), bold=True, color=GOLD_DARK, name="Calibri Light")
        txt(sl, note_, Inches(10.5), y + Inches(0.68), Inches(2.5), Inches(0.3),
            size=Pt(10.5), color=SLATE, name="Calibri")


    # ══════════════════════════════════════════════════════════════════════════════
    # SLIDE 10 — CHART: BILLING EFFICIENCY
    # ══════════════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, LIGHT_BG)
    header_bar(sl, "Billing Efficiency by Circle", "% billed connections vs 95% target benchmark")
    footer(sl, 10)
    if os.path.exists("billing_efficiency.png"):
        sl.shapes.add_picture("billing_efficiency.png", Inches(0.2), Inches(1.48), Inches(9.5), Inches(5.7))

    rect(sl, Inches(9.85), Inches(1.52), Inches(3.3), Inches(5.6), WHITE, RED, Pt(0.8))
    rect(sl, Inches(9.85), Inches(1.52), Inches(3.3), Inches(0.48), RED)
    txt(sl, "BILLING LEGEND", Inches(10.0), Inches(1.58), Inches(3.0), Inches(0.38),
        size=Pt(12), bold=True, color=WHITE, name="Calibri")
    legend_items = [
        (GREEN, "> 97%", "High performer"),
        (TEAL,  "93–97%","On target"),
        (RED,   "< 93%", "Needs attention"),
    ]
    for i, (c, rng, note_) in enumerate(legend_items):
        y = Inches(2.18) + i * Inches(0.88)
        rect(sl, Inches(10.0), y, Inches(0.42), Inches(0.42), c, radius=0.5)
        txt(sl, rng,  Inches(10.6), y,               Inches(2.4), Inches(0.32),
            size=Pt(13), bold=True, color=NAVY, name="Calibri")
        txt(sl, note_, Inches(10.6), y + Inches(0.3), Inches(2.4), Inches(0.28),
            size=Pt(10.5), color=SLATE, name="Calibri")

    rect(sl, Inches(10.0), Inches(4.98), Inches(3.0), Inches(1.9), CARD_BG, RED, Pt(0.5))
    txt(sl, "Overall Rate", Inches(10.1), Inches(5.08), Inches(2.8), Inches(0.3),
        size=Pt(10), color=SLATE, name="Calibri")
    txt(sl, "92.68%", Inches(10.1), Inches(5.38), Inches(2.8), Inches(0.65),
        size=Pt(30), bold=True, color=RED, align=PP_ALIGN.CENTER, name="Calibri Light")
    txt(sl, "Gap: ~649K connections", Inches(10.1), Inches(6.0), Inches(2.8), Inches(0.3),
        size=Pt(10), color=SLATE, align=PP_ALIGN.CENTER, name="Calibri")
    txt(sl, "Target: 95%", Inches(10.1), Inches(6.35), Inches(2.8), Inches(0.3),
        size=Pt(10.5), bold=True, color=GOLD_DARK, align=PP_ALIGN.CENTER, name="Calibri")


    # ══════════════════════════════════════════════════════════════════════════════
    # SLIDE 11 — CHART: LOAD vs UNITS (regression)
    # ══════════════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, LIGHT_BG)
    header_bar(sl, "Connected Load vs Units Consumed", "Seaborn regplot — linear correlation analysis")
    footer(sl, 11)
    if os.path.exists("load_vs_units.png"):
        sl.shapes.add_picture("load_vs_units.png", Inches(0.2), Inches(1.48), Inches(9.5), Inches(5.7))

    rect(sl, Inches(9.85), Inches(1.52), Inches(3.3), Inches(5.6), WHITE, GREEN, Pt(0.8))
    rect(sl, Inches(9.85), Inches(1.52), Inches(3.3), Inches(0.48), GREEN)
    txt(sl, "CORRELATION STATS", Inches(10.0), Inches(1.58), Inches(3.0), Inches(0.38),
        size=Pt(12), bold=True, color=WHITE, name="Calibri")

    stat_items = [
        ("Correlation", "r ≈ 0.95", "Strong positive"),
        ("Relationship", "Linear", "Load → Units"),
        ("Hue mapping", "Billing %", "Color = efficiency"),
    ]
    for i, (lbl, val, note_) in enumerate(stat_items):
        y = Inches(2.18) + i * Inches(1.05)
        rect(sl, Inches(10.0), y, Inches(3.0), Inches(0.95), CARD_BG, GREEN, Pt(0.5))
        txt(sl, lbl,  Inches(10.12), y + Inches(0.06), Inches(2.8), Inches(0.28),
            size=Pt(10), color=SLATE, name="Calibri")
        txt(sl, val,  Inches(10.12), y + Inches(0.32), Inches(2.8), Inches(0.4),
            size=Pt(18), bold=True, color=GREEN, name="Calibri Light")
        txt(sl, note_, Inches(10.12), y + Inches(0.7), Inches(2.8), Inches(0.26),
            size=Pt(10), color=SLATE, italic=True, name="Calibri")

    rect(sl, Inches(10.0), Inches(5.42), Inches(3.0), Inches(1.5), DARK_TEXT)
    txt(sl, "Interpretation", Inches(10.1), Inches(5.5), Inches(2.8), Inches(0.3),
        size=Pt(10), color=TEAL, bold=True, name="Calibri")
    txt(sl, "Areas with higher sanctioned load reliably show higher consumption. Load data can serve as a proxy for demand forecasting.",
        Inches(10.1), Inches(5.82), Inches(2.8), Inches(1.0),
        size=Pt(10), color=WHITE, name="Calibri")


    # ══════════════════════════════════════════════════════════════════════════════
    # SLIDE 12 — CHART: DISTRIBUTIONS
    # ══════════════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, LIGHT_BG)
    header_bar(sl, "Statistical Distributions (KDE + Histogram)", "Seaborn histplot with kernel density estimation")
    footer(sl, 12)
    if os.path.exists("distribution_metrics.png"):
        sl.shapes.add_picture("distribution_metrics.png", Inches(0.2), Inches(1.48), Inches(12.9), Inches(5.3))

    # Two annotation boxes below chart
    for x, c, title_, body_ in [
        (Inches(0.3),  TEAL, "Household Consumption",
         "Right-skewed distribution. Median ~90 kWh/month. Some sectors show extreme values (HT/bulk consumers)."),
        (Inches(6.8),  GREEN, "Billing Efficiency Distribution",
         "Strongly left-skewed (peak at 97-100%). Most sectors bill nearly all connections. ~7% have significant gaps."),
    ]:
        rect(sl, x, Inches(6.88), Inches(6.2), Inches(0.52), c)
        txt(sl, title_, x + Inches(0.15), Inches(6.92), Inches(2.8), Inches(0.38),
            size=Pt(12), bold=True, color=WHITE, name="Calibri")
        txt(sl, body_, x + Inches(3.1), Inches(6.92), Inches(3.0), Inches(0.42),
            size=Pt(10), color=WHITE, name="Calibri")


    # ══════════════════════════════════════════════════════════════════════════════
    # SLIDE 13 — DATA QUALITY FLOWCHART
    # ══════════════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, LIGHT_BG)
    header_bar(sl, "Data Quality Issues & Cleaning", "Identified problems and the handling strategy applied")
    footer(sl, 13)

    issues = [
        ("247 NULL\nBILLDSERVICES",  "All 247 rows also have 0 UNITS\n→ Impute from TOTSERVICES",      RED,    GREEN),
        ("47 NULL\nAREA Names",       "Geographic label missing\n→ Fill as 'UNKNOWN AREA'",             GOLD_DARK, GREEN),
        ("15 Negative\nUNITS Values", "Billing corrections / adjustments\n→ Clip to 0 (not errors)",   RED,    GOLD_DARK),
        ("1,999 Repeated\nArea Names","Same name, different sections\n→ Retain — not duplicates",       SLATE,  GREEN),
        ("Extreme Outliers\n(IQR ×1.5)","Max: 2.32M kWh vs median 37K\n→ Flag & retain valid records",RED,    GOLD_DARK),
        ("Zero Consumption\nRows",    "Inactive / seasonal areas\n→ Keep — may be temporarily idle",   NAVY_MED,GREEN),
    ]

    cols3 = 3
    for i, (issue, fix, ic, fc) in enumerate(issues):
        col = i % cols3; row = i // cols3
        bx = Inches(0.3)  + col * Inches(4.35)
        by = Inches(1.58) + row * Inches(2.4)
        bw = Inches(4.1); bh = Inches(2.1)

        # Issue box
        rect(sl, bx, by, bw, Inches(1.0), WHITE, ic, Pt(1.5))
        rect(sl, bx, by, bw, Inches(0.08), ic)
        rect(sl, bx, by, Inches(0.08), Inches(1.0), ic)
        txt(sl, "ISSUE", bx + Inches(0.15), by + Inches(0.08), bw - Inches(0.2), Inches(0.24),
            size=Pt(8.5), bold=True, color=ic, name="Calibri")
        txt(sl, issue, bx + Inches(0.15), by + Inches(0.3), bw - Inches(0.2), Inches(0.65),
            size=Pt(13), bold=True, color=DARK_TEXT, name="Calibri")

        # Arrow down
        arrow_v(sl, bx + bw/2, by + Inches(1.02), Inches(0.32), SLATE)

        # Fix box
        fy = by + Inches(1.36)
        rect(sl, bx, fy, bw, Inches(0.68), RGBColor(0xe8,0xf8,0xf0) if fc == GREEN else RGBColor(0xff,0xf8,0xe8), fc, Pt(1))
        txt(sl, "FIX", bx + Inches(0.12), fy + Inches(0.06), Inches(0.5), Inches(0.22),
            size=Pt(8.5), bold=True, color=fc, name="Calibri")
        txt(sl, fix, bx + Inches(0.12), fy + Inches(0.26), bw - Inches(0.2), Inches(0.42),
            size=Pt(11.5), color=DARK_TEXT, name="Calibri")


    # ══════════════════════════════════════════════════════════════════════════════
    # SLIDE 14 — KEY INSIGHTS (visual callout cards)
    # ══════════════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, LIGHT_BG)
    header_bar(sl, "Key Analytical Insights", "Most important patterns discovered in the data")
    footer(sl, 14)

    insights = [
        (NAVY,     "Urban Concentration",
                   "Top 10 Hyderabad-region circles = ~90% of total domestic units consumed."),
        (TEAL_DARK,"Per-Service Gap",
                   "Banjara Hills: ~240 kWh/service vs district circles: ~75–83 kWh. Urban = 3x higher."),
        (GOLD_DARK,"Billing Revenue Risk",
                   "92.68% billing rate means ~649,000 unbilled connections — a significant revenue leakage."),
        (GREEN,    "Load as Predictor",
                   "Units & load correlation r≈0.95. Sanctioned load is a reliable consumption proxy."),
        (RGBColor(0xb0,0x40,0x00), "Demand Concentration",
                   "Top 15% of sectors by size (2,452 areas) account for 50.5% of all kWh consumed."),
        (RED,      "Bulk Consumer Distortion",
                   "'DOMESTIC HT/HV/CT' rows show 3,000–11,500 kWh/service — likely bulk, not household."),
    ]

    for i, (ac, title_, body_) in enumerate(insights):
        col = i % 3; row = i // 2
        x = Inches(0.3) + col * Inches(4.35)
        y = Inches(1.6)  + (i // 3) * Inches(2.55)

        rect(sl, x, y, Inches(4.1), Inches(2.28), WHITE, ac, Pt(1.5))
        rect(sl, x, y, Inches(4.1), Inches(0.52), ac)
        txt(sl, title_, x + Inches(0.15), y + Inches(0.1), Inches(3.8), Inches(0.38),
            size=Pt(13), bold=True, color=WHITE, name="Calibri")
        txt(sl, body_,  x + Inches(0.15), y + Inches(0.65), Inches(3.8), Inches(1.5),
            size=Pt(12.5), color=DARK_TEXT, name="Calibri")


    # ══════════════════════════════════════════════════════════════════════════════
    # SLIDE 15 — LIMITATIONS (visual table)
    # ══════════════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, LIGHT_BG)
    header_bar(sl, "Dataset Limitations", "Boundaries of the analysis — what the data cannot tell us")
    footer(sl, 15)

    lims = [
        ("No Time-Series",           "Single July 2026 snapshot only — cannot track consumption trends over time.",              "Time"),
        ("Domestic Only",            "Commercial & industrial consumers excluded — incomplete total demand picture.",             "Scope"),
        ("No Weather Data",          "Cannot explain seasonal consumption peaks without temperature/climate variables.",          "Context"),
        ("No Household Details",     "Income, family size, appliance usage unknown — cannot explain why consumption differs.",    "Depth"),
        ("Repeated Area Names",      "1,999 area names appear in multiple sections — spatial deduplication needed.",             "Quality"),
        ("Negative Units",           "15 billing-correction rows cannot be distinguished from real negative consumption.",       "Quality"),
        ("HT/HV Domestic Rows",      "Bulk consumers billed as 'domestic' inflate category averages significantly.",            "Accuracy"),
        ("No Socioeconomic Data",    "Cannot correlate consumption with poverty index, HDI, or household density.",              "Context"),
    ]

    # Header
    for cx, cw, lbl in [(Inches(0.3),Inches(3.5),"LIMITATION"), (Inches(3.85),Inches(8.2),"IMPACT & EXPLANATION"), (Inches(12.1),Inches(1.1),"CATEGORY")]:
        rect(sl, cx, Inches(1.53), cw, Inches(0.44), NAVY)
        txt(sl, lbl, cx + Inches(0.1), Inches(1.58), cw - Inches(0.15), Inches(0.35),
            size=Pt(11), bold=True, color=WHITE, name="Calibri")

    for i, (lim, exp, cat) in enumerate(lims):
        y = Inches(1.97) + i * Inches(0.56)
        bg = CARD_BG if i % 2 == 0 else WHITE
        rect(sl, Inches(0.3),  y, Inches(12.9), Inches(0.55), bg)
        txt(sl, lim,  Inches(0.4),  y + Inches(0.1), Inches(3.3), Inches(0.38),
            size=Pt(11.5), bold=True, color=NAVY, name="Calibri")
        txt(sl, exp,  Inches(3.88), y + Inches(0.1), Inches(8.1), Inches(0.38),
            size=Pt(11), color=DARK_TEXT, name="Calibri")
        cc = TEAL if cat == "Scope" else GOLD_DARK if cat == "Context" else RED if cat == "Quality" else NAVY_MED
        rect(sl, Inches(12.15), y + Inches(0.1), Inches(0.95), Inches(0.32), cc, radius=0.5)
        txt(sl, cat, Inches(12.15), y + Inches(0.12), Inches(0.95), Inches(0.28),
            size=Pt(8.5), color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")


    # ══════════════════════════════════════════════════════════════════════════════
    # SLIDE 16 — TECH STACK (visual icons/boxes)
    # ══════════════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, LIGHT_BG)
    header_bar(sl, "Technology Stack", "Tools, libraries & technologies used in this project")
    footer(sl, 16)

    # Two-column: Python | Web
    for side_x, side_color, side_label, tools in [
        (Inches(0.3),  NAVY,  "PYTHON ANALYSIS STACK", [
            (TEAL,     "Pandas",      "Data loading, cleaning, groupby aggregation & feature engineering"),
            (NAVY_MED, "NumPy",       "Vectorized operations: billing rate, outlier detection, clipping"),
            (GOLD_DARK,"Matplotlib",  "Publication-quality dark-theme static charts (5 PNG outputs)"),
            (GREEN,    "Seaborn",     "Statistical plots: regplot, histplot+KDE, barplot with hue palette"),
            (SLATE,    "python-pptx", "Automated generation of this presentation programmatically"),
            (RED,      "json",        "Exports analysis_summary.json for live dashboard data feed"),
        ]),
        (Inches(6.85), TEAL_DARK, "WEB DASHBOARD STACK", [
            (NAVY,     "HTML5 + CSS3",    "Glassmorphism dark UI, CSS variables, responsive grid layout"),
            (TEAL,     "JavaScript",      "Dynamic filter slicers, search, pagination & CSV export logic"),
            (GOLD_DARK,"Chart.js",        "Interactive bar/line combo, scatter plots, doughnut charts"),
            (GREEN,    "KPI Cards",       "Real-time metric tiles recalculated on every filter change"),
            (TEAL_DARK,"Anomaly Alerts",  "Auto-flags sections with billing rate < 85% or high load intensity"),
            (RED,      "Gallery Modal",   "Click-to-expand viewer for all 5 Matplotlib PNG reports"),
        ]),
    ]:
        rect(sl, side_x, Inches(1.52), Inches(6.3), Inches(5.65), WHITE, side_color, Pt(1))
        rect(sl, side_x, Inches(1.52), Inches(6.3), Inches(0.5), side_color)
        txt(sl, side_label, side_x + Inches(0.2), Inches(1.58), Inches(5.9), Inches(0.38),
            size=Pt(12), bold=True, color=WHITE, name="Calibri")
        for j, (c, name_, desc) in enumerate(tools):
            y = Inches(2.15) + j * Inches(0.82)
            rect(sl, side_x + Inches(0.15), y, Inches(5.95), Inches(0.76), CARD_BG, c, Pt(0.8))
            rect(sl, side_x + Inches(0.15), y, Inches(0.08), Inches(0.76), c)
            txt(sl, name_, side_x + Inches(0.35), y + Inches(0.06), Inches(2.0), Inches(0.3),
                size=Pt(13), bold=True, color=c, name="Calibri")
            txt(sl, desc,  side_x + Inches(0.35), y + Inches(0.36), Inches(5.5), Inches(0.4),
                size=Pt(10.5), color=DARK_TEXT, name="Calibri")

    # Divider
    rect(sl, Inches(6.7), Inches(1.52), Inches(0.08), Inches(5.65), RGBColor(0xdd,0xe5,0xee))


    # ══════════════════════════════════════════════════════════════════════════════
    # SLIDE 17 — CONCLUSION & THANK YOU
    # ══════════════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, NAVY)

    # Background accents
    rect(sl, SW - Inches(4.5), 0, Inches(4.5), SH, NAVY_MED)
    rect(sl, SW - Inches(4.5), 0, Inches(0.055), SH, TEAL)
    rect(sl, 0, 0, SW, Inches(0.09), GOLD)
    rect(sl, 0, SH - Inches(0.09), SW, Inches(0.09), TEAL)

    txt(sl, "Conclusions & Summary",
        Inches(0.5), Inches(0.25), Inches(8.5), Inches(0.75),
        size=Pt(34), bold=True, color=WHITE, name="Calibri Light")
    rect(sl, Inches(0.5), Inches(1.0), Inches(4.5), Inches(0.055), GOLD)

    conclusions = [
        ("Urban Dominance",    "Top 10 Hyderabad circles = ~90% of all domestic electricity consumed"),
        ("Per-Service Gap",    "Banjara Hills: 240 kWh vs district circles: ~80 kWh/month (3x difference)"),
        ("Billing Gap",        "92.68% efficiency → ~649K unbilled connections = revenue recovery opportunity"),
        ("Strong Correlation", "Load–consumption r ≈ 0.95 → load data is a reliable demand proxy"),
        ("Demand Skew",        "Top 15% of sectors (by size) consume over 50% of all electricity"),
    ]
    for i, (title_, body_) in enumerate(conclusions):
        y = Inches(1.12) + i * Inches(0.84)
        rect(sl, Inches(0.5), y, Inches(8.1), Inches(0.72), RGBColor(0x12,0x25,0x52), TEAL, Pt(0.5))
        rect(sl, Inches(0.5), y, Inches(0.08), Inches(0.72), GOLD)
        txt(sl, title_, Inches(0.7), y + Inches(0.06), Inches(2.4), Inches(0.28),
            size=Pt(12), bold=True, color=GOLD, name="Calibri")
        txt(sl, body_,  Inches(3.15), y + Inches(0.12), Inches(5.35), Inches(0.48),
            size=Pt(12.5), color=WHITE, name="Calibri")

    # GitHub + run strip
    rect(sl, Inches(0.5), Inches(5.42), Inches(8.1), Inches(0.6), RGBColor(0x0a,0x14,0x2e), TEAL, Pt(1))
    txt(sl, "GitHub:", Inches(0.7), Inches(5.52), Inches(1.3), Inches(0.42),
        size=Pt(12), bold=True, color=TEAL, name="Calibri")
    txt(sl, "https://github.com/arjunkanukuntla/dav-project",
        Inches(2.0), Inches(5.52), Inches(6.5), Inches(0.42),
        size=Pt(12), color=WHITE, name="Calibri")

    rect(sl, Inches(0.5), Inches(6.08), Inches(8.1), Inches(0.5), RGBColor(0x08,0x0e,0x1e), NAVY_MED, Pt(0.5))
    txt(sl, "$ python main.py    |    $ python -m http.server 8000    →    localhost:8000",
        Inches(0.7), Inches(6.15), Inches(7.8), Inches(0.38),
        size=Pt(11), color=TEAL, name="Courier New")

    # Right panel — team
    rect(sl, SW - Inches(4.3), Inches(0.3), Inches(4.1), Inches(6.7), RGBColor(0x10,0x1f,0x48), TEAL, Pt(0.5))
    txt(sl, "TEAM", SW - Inches(4.1), Inches(0.55), Inches(3.7), Inches(0.45),
        size=Pt(10), bold=True, color=TEAL, align=PP_ALIGN.CENTER, name="Calibri")
    txt(sl, "Thank You", SW - Inches(4.1), Inches(0.95), Inches(3.7), Inches(0.65),
        size=Pt(26), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri Light")
    rect(sl, SW - Inches(4.1), Inches(1.6), Inches(3.7), Inches(0.04), GOLD)

    for i, (name_, roll, dept) in enumerate([
        ("D. Chander", "2503A52214", "B.Tech CSE"),
        ("K. Arjun",   "2503A52235", "B.Tech CSE"),
        ("P. Sachin",  "2503A52240", "B.Tech CSE"),
    ]):
        y = Inches(1.85) + i * Inches(1.4)
        rect(sl, SW - Inches(4.1), y, Inches(3.7), Inches(1.22), RGBColor(0x12,0x25,0x52), TEAL, Pt(0.5))
        txt(sl, name_, SW - Inches(3.95), y + Inches(0.1), Inches(3.4), Inches(0.42),
            size=Pt(17), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")
        txt(sl, roll,  SW - Inches(3.95), y + Inches(0.5), Inches(3.4), Inches(0.32),
            size=Pt(12), color=GOLD, align=PP_ALIGN.CENTER, name="Calibri")
        txt(sl, dept,  SW - Inches(3.95), y + Inches(0.82), Inches(3.4), Inches(0.28),
            size=Pt(10.5), color=MID_GRAY, align=PP_ALIGN.CENTER, name="Calibri")

    txt(sl, "Data Analytics & Visualization  |  July 2026",
        SW - Inches(4.1), Inches(6.08), Inches(3.7), Inches(0.42),
        size=Pt(10.5), color=MID_GRAY, align=PP_ALIGN.CENTER, italic=True, name="Calibri")

    # ── SAVE ─────────────────────────────────────────────────────────────────────
    out = "TG_SPDCL_Electricity_Analysis_v2.pptx"
    prs.save(out)
    print(f"DONE — Saved: {out}  |  17 slides  |  Widescreen 13.33x7.5\"")

if __name__ == "__main__":
    build()
