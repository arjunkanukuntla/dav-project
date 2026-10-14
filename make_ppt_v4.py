# -*- coding: utf-8 -*-
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

"""
TG-SPDCL Electricity Consumption & Monitoring
Academic Presentation Generator — Version 4 (v4)
High-Density Executive Edition — Strictly under 12 slides (11 Slides Total)
Complete coverage of the 8-Pillar Classroom Syllabus:
  1. Project Explanation
  2. Data Exploration
  3. Basic Statistics
  4. Data Wrangling
  5. Variable Connections (Correlation)
  6. Basic Visualization using Excel
  7. Advanced Visualization using Python (Matplotlib & Seaborn)
  8. Final Decision

Authors: D. Chander, K. Arjun, P. Sachin
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# ─── DESIGN PALETTE (PREMIUM CONSULTING THEME) ────────────────────────────────
NAVY        = RGBColor(0x0a, 0x19, 0x3b)   # Deep Midnight Navy
NAVY_MED    = RGBColor(0x16, 0x2b, 0x5a)   # Medium Navy
TEAL        = RGBColor(0x00, 0xb4, 0xd8)   # Electric Teal
TEAL_DARK   = RGBColor(0x00, 0x77, 0x91)   # Deep Teal
GOLD        = RGBColor(0xf5, 0x9e, 0x0b)   # Warm Gold / Amber
GOLD_DARK   = RGBColor(0xd9, 0x77, 0x06)
GREEN       = RGBColor(0x10, 0xb9, 0x81)   # Emerald Green
GREEN_DARK  = RGBColor(0x04, 0x78, 0x57)
RED         = RGBColor(0xef, 0x44, 0x44)   # Coral / Alert Red
SLATE       = RGBColor(0x64, 0x74, 0x8b)   # Secondary Text
SLATE_LIGHT = RGBColor(0x94, 0xa3, 0xb8)
DARK_TEXT   = RGBColor(0x0f, 0x17, 0x2a)   # Charcoal High-Contrast Body
LIGHT_BG    = RGBColor(0xf8, 0xfa, 0xfc)   # Clean Slide Canvas
WHITE       = RGBColor(0xff, 0xff, 0xff)
BORDER_GRAY = RGBColor(0xe2, 0xe8, 0xf0)   # Subtle Border
ROW_ALT     = RGBColor(0xf1, 0xf5, 0xf9)   # Alternate Table Row
EXCEL_GREEN = RGBColor(0x10, 0x7c, 0x41)

SW = Inches(13.333)
SH = Inches(7.5)

# ─── REUSABLE UI HELPERS ──────────────────────────────────────────────────────

def set_bg(slide, color=LIGHT_BG):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def rect(slide, l, t, w, h, fill, line=None, lw=Pt(0), radius=None):
    if radius:
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
        shp.adjustments[0] = radius
    else:
        shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line:
        shp.line.color.rgb = line
        shp.line.width = lw
    else:
        shp.line.fill.background()
    return shp

def txt(slide, text, l, t, w, h, size=Pt(13), bold=False, color=DARK_TEXT,
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

def header_bar(slide, pillar_tag, title, subtitle=None):
    """Compact high-density banner with syllabus pillar badge."""
    rect(slide, 0, 0, SW, Inches(1.18), NAVY)
    rect(slide, 0, Inches(1.18), SW, Inches(0.04), GOLD)
    
    # Pillar Tag Pill
    rect(slide, Inches(0.5), Inches(0.12), Inches(3.2), Inches(0.26), TEAL_DARK, radius=0.4)
    txt(slide, pillar_tag.upper(), Inches(0.55), Inches(0.15), Inches(3.1), Inches(0.22),
        size=Pt(9), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")
    
    # Title
    txt(slide, title, Inches(0.5), Inches(0.42), Inches(9.8), Inches(0.45),
        size=Pt(21), bold=True, color=WHITE, name="Calibri")
    
    # Subtitle
    if subtitle:
        txt(slide, subtitle, Inches(0.5), Inches(0.88), Inches(9.8), Inches(0.26),
            size=Pt(11), color=SLATE_LIGHT, italic=True, name="Calibri")
            
    # Right-side metadata badge
    txt(slide, "TG-SPDCL | DAV Project v4", Inches(9.8), Inches(0.42), Inches(3.0), Inches(0.26),
        size=Pt(10), bold=True, color=TEAL, align=PP_ALIGN.RIGHT, name="Calibri")
    txt(slide, "Compact Syllabus Edition (11 Slides)", Inches(9.8), Inches(0.72), Inches(3.0), Inches(0.24),
        size=Pt(8.5), color=SLATE_LIGHT, align=PP_ALIGN.RIGHT, name="Calibri")

def footer(slide, slide_num, total=11):
    rect(slide, 0, SH - Inches(0.36), SW, Inches(0.36), NAVY)
    txt(slide, "TG-SPDCL Domestic Electricity Consumption & Monitoring  |  Data Analytics & Visualization (DAV)  |  July 2026",
        Inches(0.5), SH - Inches(0.31), Inches(10.5), Inches(0.26),
        size=Pt(8.5), color=SLATE_LIGHT, name="Calibri")
    txt(slide, f"Slide {slide_num} of {total}", Inches(11.2), SH - Inches(0.31), Inches(1.6), Inches(0.26),
        size=Pt(9), bold=True, color=WHITE, align=PP_ALIGN.RIGHT, name="Calibri")

def card(slide, l, t, w, h, bg=WHITE, border=BORDER_GRAY, radius=0.04):
    return rect(slide, l, t, w, h, bg, border, Pt(0.75), radius=radius)

def bullet_card(slide, l, t, w, h, title, items, header_color=NAVY):
    card(slide, l, t, w, h)
    rect(slide, l, t, w, Inches(0.38), header_color, radius=0.04)
    txt(slide, title, l + Inches(0.18), t + Inches(0.07), w - Inches(0.36), Inches(0.26),
        size=Pt(11.5), bold=True, color=WHITE, name="Calibri")
    
    box = slide.shapes.add_textbox(l + Inches(0.18), t + Inches(0.42), w - Inches(0.36), h - Inches(0.48))
    tf = box.text_frame
    tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(3)
        p.space_after = Pt(2)
        if isinstance(it, tuple):
            r1 = p.add_run()
            r1.text = f"•  {it[0]}: "
            r1.font.bold = True
            r1.font.size = Pt(9.5)
            r1.font.color.rgb = DARK_TEXT
            r1.font.name = "Calibri"
            
            r2 = p.add_run()
            r2.text = it[1]
            r2.font.bold = False
            r2.font.size = Pt(9.5)
            r2.font.color.rgb = SLATE
            r2.font.name = "Calibri"
        else:
            r = p.add_run()
            r.text = f"•  {it}"
            r.font.size = Pt(9.5)
            r.font.color.rgb = DARK_TEXT
            r.font.name = "Calibri"

def add_image_card(slide, l, t, w, h, img_path, caption=""):
    card(slide, l, t, w, h)
    if os.path.exists(img_path):
        pad = Inches(0.08)
        img_h = h - Inches(0.32) if caption else h - (pad * 2)
        slide.shapes.add_picture(img_path, l + pad, t + pad, w - (pad * 2), img_h)
        if caption:
            txt(slide, caption, l + pad, t + h - Inches(0.28), w - (pad * 2), Inches(0.24),
                size=Pt(8.5), bold=True, color=NAVY_MED, align=PP_ALIGN.CENTER, name="Calibri")
    else:
        txt(slide, f"[Image File Not Found: {img_path}]", l + Inches(0.2), t + h/2 - Inches(0.2),
            w - Inches(0.4), Inches(0.4), size=Pt(10), bold=True, color=RED, align=PP_ALIGN.CENTER)

def create_table(slide, l, t, w, h, headers, rows, col_widths=None):
    num_rows = len(rows) + 1
    num_cols = len(headers)
    table_shape = slide.shapes.add_table(num_rows, num_cols, l, t, w, h)
    table = table_shape.table
    
    if col_widths:
        for idx, width in enumerate(col_widths):
            table.columns[idx].width = width
            
    for c_idx, head in enumerate(headers):
        cell = table.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = head
        run.font.bold = True
        run.font.size = Pt(9.5)
        run.font.color.rgb = WHITE
        run.font.name = "Calibri"
        
    for r_idx, row in enumerate(rows):
        bg = ROW_ALT if r_idx % 2 == 1 else WHITE
        for c_idx, val in enumerate(row):
            cell = table.cell(r_idx + 1, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if c_idx > 0 else PP_ALIGN.LEFT
            run = p.add_run()
            run.text = str(val)
            run.font.size = Pt(9)
            run.font.name = "Calibri"
            if c_idx == 0:
                run.font.bold = True
                run.font.color.rgb = NAVY
            else:
                run.font.color.rgb = DARK_TEXT

# ─── PRESENTATION BUILDER (11 SLIDES STRICTLY UNDER 12) ────────────────────────

def build_presentation_v4():
    prs = Presentation()
    prs.slide_width  = SW
    prs.slide_height = SH
    BLANK = prs.slide_layouts[6]

    # ==============================================================================
    # SLIDE 1: TITLE & EXECUTIVE LAUNCHPAD
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, NAVY)
    
    # Decorative backdrop panels
    rect(sl, Inches(8.3), Inches(0), Inches(5.033), SH, RGBColor(0x13, 0x24, 0x4e))
    rect(sl, Inches(9.5), Inches(1.2), Inches(3.833), Inches(5.4), RGBColor(0x18, 0x2c, 0x5e))
    rect(sl, Inches(0.8), Inches(3.2), Inches(4.5), Inches(0.06), GOLD)

    # Category Pill
    rect(sl, Inches(0.8), Inches(0.9), Inches(3.4), Inches(0.36), TEAL, radius=0.4)
    txt(sl, "DATA ANALYTICS & VISUALIZATION (DAV)", Inches(0.9), Inches(0.95), Inches(3.2), Inches(0.25),
        size=Pt(9.5), bold=True, color=WHITE, name="Calibri")

    # Titles
    txt(sl, "Electricity Consumption", Inches(0.8), Inches(1.35), Inches(7.5), Inches(0.85),
        size=Pt(38), bold=True, color=WHITE, name="Calibri")
    txt(sl, "& Monitoring Analysis", Inches(0.8), Inches(2.15), Inches(7.5), Inches(0.85),
        size=Pt(38), bold=True, color=TEAL, name="Calibri")
    txt(sl, "TG-SPDCL Domestic Consumer Dataset (July 2026) | 8-Pillar Syllabus Synthesis",
        Inches(0.8), Inches(3.35), Inches(7.5), Inches(0.35),
        size=Pt(13), color=SLATE_LIGHT, italic=True, name="Calibri")

    # Team Members Card
    card(sl, Inches(0.8), Inches(4.1), Inches(6.9), Inches(2.6), bg=RGBColor(0x13, 0x24, 0x4e), border=TEAL_DARK)
    rect(sl, Inches(0.8), Inches(4.1), Inches(6.9), Inches(0.06), TEAL)
    txt(sl, "PROJECT RESEARCH TEAM", Inches(1.0), Inches(4.25), Inches(4.0), Inches(0.24),
        size=Pt(10.5), bold=True, color=GOLD, name="Calibri")

    team = [
        ("D. Chander", "Roll: 2503A52214", "Data Cleaning, Wrangling & Power Query ETL"),
        ("K. Arjun",   "Roll: 2503A52235", "Seaborn Visualizations, Correlations & Heatmaps"),
        ("P. Sachin",  "Roll: 2503A52240", "Statistical Modeling, Pareto Analysis & Final Decision")
    ]
    for idx, (m_name, m_roll, m_role) in enumerate(team):
        y_pos = Inches(4.62) + idx * Inches(0.62)
        txt(sl, m_name, Inches(1.0), y_pos, Inches(2.0), Inches(0.3),
            size=Pt(14), bold=True, color=WHITE, name="Calibri")
        txt(sl, m_roll, Inches(3.0), y_pos + Inches(0.02), Inches(1.7), Inches(0.25),
            size=Pt(10.5), bold=True, color=TEAL, name="Calibri")
        txt(sl, m_role, Inches(4.7), y_pos + Inches(0.02), Inches(2.8), Inches(0.25),
            size=Pt(9), color=SLATE_LIGHT, name="Calibri")

    # 4 Key Quantitative Anchors
    stat_boxes = [
        ("15,947", "SECTOR RECORDS", "20 Circles | 49 Divisions | 476 Sections", TEAL),
        ("8.87M", "TOTAL SERVICES", "8.22M Active Billed Connections", GOLD),
        ("1,245.5 GWh", "ENERGY CONSUMED", "1,245,523,023 kWh Monthly Demand", GREEN),
        ("92.68%", "BILLING EFFICIENCY", "648,960 Unbilled Services Identified", RGBColor(0xf9, 0x73, 0x16))
    ]
    for idx, (val, title_lbl, desc_lbl, accent) in enumerate(stat_boxes):
        y_b = Inches(1.1) + idx * Inches(1.4)
        card(sl, Inches(8.8), y_b, Inches(4.0), Inches(1.22), bg=RGBColor(0x0f, 0x17, 0x2a), border=accent)
        rect(sl, Inches(8.8), y_b, Inches(0.08), Inches(1.22), accent)
        txt(sl, val, Inches(9.0), y_b + Inches(0.12), Inches(3.6), Inches(0.42),
            size=Pt(22), bold=True, color=WHITE, name="Calibri")
        txt(sl, title_lbl, Inches(9.0), y_b + Inches(0.56), Inches(3.6), Inches(0.24),
            size=Pt(9), bold=True, color=accent, name="Calibri")
        txt(sl, desc_lbl, Inches(9.0), y_b + Inches(0.82), Inches(3.6), Inches(0.25),
            size=Pt(8.5), color=SLATE_LIGHT, name="Calibri")

    footer(sl, 1)

    # ==============================================================================
    # SLIDE 2: PILLAR 1 — PROJECT EXPLANATION & SCOPE
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Pillar 1: Project Explanation", "Domain Context, Problem Statement & Network Hierarchy",
               "Analyzing electricity distribution dynamics across Greater Hyderabad and rural Telangana")

    # 5-Tier Distribution Hierarchy Flow (Top Strip)
    hierarchy = [
        ("20 Circles", "Apex Geographic Zones", "Medchal, Cybercity, Habsiguda...", NAVY),
        ("49 Divisions", "Operational Operations", "Kukatpally, Mehdipatnam...", NAVY_MED),
        ("131 Subdivisions", "Field Engineering", "Miyapur, Qutbullapur, Uppal...", TEAL_DARK),
        ("476 Sections", "Local Billing Stations", "Line stations & meter reading offices", TEAL),
        ("15,947 Areas", "Neighborhood Sectors", "Granular colony & feeder records", GREEN_DARK)
    ]
    for idx, (h_title, h_sub, h_desc, h_color) in enumerate(hierarchy):
        x_h = Inches(0.5) + idx * Inches(2.52)
        card(sl, x_h, Inches(1.45), Inches(2.35), Inches(1.6))
        rect(sl, x_h, Inches(1.45), Inches(2.35), Inches(0.35), h_color, radius=0.04)
        txt(sl, h_title, x_h + Inches(0.1), Inches(1.5), Inches(2.15), Inches(0.25),
            size=Pt(11), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")
        txt(sl, h_sub, x_h + Inches(0.1), Inches(1.85), Inches(2.15), Inches(0.22),
            size=Pt(8.5), bold=True, color=h_color, align=PP_ALIGN.CENTER, name="Calibri")
        txt(sl, h_desc, x_h + Inches(0.1), Inches(2.1), Inches(2.15), Inches(0.85),
            size=Pt(8), color=SLATE, align=PP_ALIGN.CENTER, name="Calibri")

    # 3 Analytical Dimension Cards Below
    bullet_card(sl, Inches(0.5), Inches(3.25), Inches(3.95), Inches(3.7),
                "1. Utility Context (TG-SPDCL)", [
                    ("State Distribution Entity", "Serves 15 southern districts of Telangana including the state capital."),
                    ("Domestic Focus (Cat 1)", "Exclusively covers households, apartments, and domestic residential connections."),
                    ("Demand Significance", "Domestic consumption constitutes over 40% of grid peak load during summer/monsoon.")
                ], NAVY)

    bullet_card(sl, Inches(4.7), Inches(3.25), Inches(3.95), Inches(3.7),
                "2. Core Operational Challenges", [
                    ("High Unbilled Deficit", "648,960 services unbilled in July 2026 (~7.32% of total services)."),
                    ("Severe Load Concentration", "Heavy air conditioning & villa loads strain local transformers in western corridors."),
                    ("Data Hygiene Gaps", "Missing records, negative unit reversals, and repetitive naming across sections.")
                ], RED)

    bullet_card(sl, Inches(8.9), Inches(3.25), Inches(3.95), Inches(3.7),
                "3. Core Research Objectives", [
                    ("Consumption Profiling", "Establish consumption intensity per household across metro vs rural zones."),
                    ("Correlation Modeling", "Quantify the mathematical coupling between connected load (kW) and units (kWh)."),
                    ("Data-Backed Interventions", "Deliver actionable transformer upgrade and revenue recovery roadmaps.")
                ], TEAL_DARK)

    footer(sl, 2)

    # ==============================================================================
    # SLIDE 3: PILLAR 2 — DATA EXPLORATION & QUALITY AUDIT
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Pillar 2: Data Exploration", "Dataset Architecture, Data Dictionary & Data Quality Audit",
               "Comprehensive structural audit of 15,947 records × 11 active features")

    # Left: Column Data Dictionary Table
    headers = ["Feature Name", "Type", "Category", "Description / Analytical Purpose"]
    rows = [
        ["CIRCLE", "String", "Categorical", "Highest administrative circle (20 unique)"],
        ["DIVISION", "String", "Categorical", "Operational district subdivision (49 unique)"],
        ["SUBDIVISION", "String", "Categorical", "Field maintenance zone (131 unique)"],
        ["SECTION", "String", "Categorical", "Local billing station (476 unique)"],
        ["AREA", "String", "Categorical", "Neighborhood sector name (15,947 records)"],
        ["CATDESC", "String", "Categorical", "Tariff category (DOMESTIC)"],
        ["CATCODE", "Integer", "Numerical", "Tariff code identifier (Code = 1)"],
        ["TOTSERVICES", "Integer", "Numerical", "Total registered domestic connections"],
        ["BILLDSERVICES", "Float/Int", "Numerical", "Active connections billed in July 2026"],
        ["UNITS", "Float", "Numerical", "Metered electricity consumed in kWh"],
        ["LOAD", "Float", "Numerical", "Sanctioned connected load in Kilowatts"]
    ]
    col_w = [Inches(1.5), Inches(1.0), Inches(1.2), Inches(3.2)]
    create_table(sl, Inches(0.5), Inches(1.45), Inches(6.9), Inches(5.5), headers, rows, col_w)

    # Right: Data Quality & Anomaly Diagnostic Cards
    bullet_card(sl, Inches(7.6), Inches(1.45), Inches(5.2), Inches(2.65),
                "Missing Values & Root-Cause Audit", [
                    ("247 Missing BILLDSERVICES", "Crucial Discovery: 100% of these 247 rows have exactly 0 metered units. They represent uncommissioned or inactive feeders."),
                    ("47 Missing AREA Names", "Handled via string imputation as 'UNKNOWN AREA' without loss of geographic hierarchy."),
                    ("Zero Removal Strategy", "Retained rather than dropped to prevent artificially biasing average consumption upward.")
                ], RED)

    bullet_card(sl, Inches(7.6), Inches(4.3), Inches(5.2), Inches(2.65),
                "Numerical Outliers & Duplicate Anomalies", [
                    ("15 Negative UNIT Records", "Min value -1,703.0 kWh. Represent utility accounting adjustments and meter recalibrations; clamped to 0.0 in analytics."),
                    ("344 Zero-Unit Sectors", "Valid vacant homes, disconnected lines, or rural tribal thandas in Nalgonda/Suryapet."),
                    ("1,999 Repeated Area Names", "'MAIN ROAD' & 'GANDHI NAGAR' repeat across sections; composite keys (`section_area`) enforce unique identity.")
                ], GOLD_DARK)

    footer(sl, 3)

    # ==============================================================================
    # SLIDE 4: PILLAR 3 — BASIC STATISTICS & THE PARETO RULE
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Pillar 3: Basic Statistics", "Statistical Dispersion, Skewness & The 15% / 50.5% Pareto Rule",
               "Parametric vs non-parametric summary, urban-rural disparity, and heavy-consumer concentration")

    # Top Half: Five-Number Statistical Summary Table
    headers = ["Metric / Variable", "Registered Services", "Billed Services", "Consumption (kWh)", "Connected Load (kW)"]
    rows = [
        ["Mean (Average)", "556.1 services", "523.0 services", "78,103.6 kWh", "864.04 kW"],
        ["Std Deviation", "626.2", "605.8", "112,930.4", "1,529.57"],
        ["Minimum", "1", "1", "-1,703.0 (Anomalous)", "0.12 kW"],
        ["25th Percentile (Q1)", "152.0", "138.0", "11,312.0 kWh", "86.00 kW"],
        ["50th Percentile (Median)", "382.0", "345.0", "36,826.0 kWh", "316.00 kW"],
        ["75th Percentile (Q3)", "735.5", "690.0", "102,405.5 kWh", "1,058.15 kW"],
        ["Maximum Value", "12,182", "11,886", "2,319,144.0 kWh", "33,617.69 kW"]
    ]
    col_w = [Inches(2.5), Inches(2.3), Inches(2.3), Inches(2.7), Inches(2.5)]
    create_table(sl, Inches(0.5), Inches(1.45), Inches(12.33), Inches(2.7), headers, rows, col_w)

    # Bottom Left: The 15% / 50.5% Pareto Rule
    bullet_card(sl, Inches(0.5), Inches(4.35), Inches(6.0), Inches(2.6),
                "The Pareto Principle: Area-Size Segmentation", [
                    ("Very Large Areas (>1,000 services)", "2,452 areas (15.4% of total) account for 50.5% (629.0 GWh) of all electricity consumed."),
                    ("Large Areas (501–1,000 services)", "3,810 areas (23.9% count) account for 24.2% (301.4 GWh) of total demand."),
                    ("Medium Areas (201–500 services)", "5,120 areas (32.1% count) account for 18.1% (225.4 GWh)."),
                    ("Small Areas (1–200 services)", "4,565 areas (28.6% count) account for only 7.2% (89.7 GWh).")
                ], RED)

    # Bottom Right: Geographic Disparity Insights
    bullet_card(sl, Inches(6.8), Inches(4.35), Inches(6.0), Inches(2.6),
                "Geographic Rankings & The 3× Lifestyle Gap", [
                    ("Top 5 Circles Consume 50.6%", "Medchal (141.7M), Habsiguda (138.1M), Cybercity (125.4M), Hyd Central (112.5M), Saroornagar (112.3M)."),
                    ("Bottom 5 Circles < 1.5% each", "Gadwal (10.0M), Wanaparthy (10.1M), Nagarkurnool (13.1M), Medak (15.9M), Vikarabad (17.9M)."),
                    ("Household Intensity Divide", "Banjara Hills (240 kWh) & Cybercity (219 kWh) vs Rural Districts (75–83 kWh) — a 3× consumption lifestyle gap."),
                    ("Extreme Right Skewness", "Max sector Hydernagar (2.32M kWh) is 63× greater than the state median (36.8k kWh).")
                ], NAVY)

    footer(sl, 4)

    # ==============================================================================
    # SLIDE 5: PILLAR 4 — DATA WRANGLING & FEATURE ENGINEERING
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Pillar 4: Data Wrangling", "Automated ETL Pipeline & Synthesized Feature Engineering",
               "Systematic transformations in Power Query and Pandas to establish clean analytical variables")

    # Top Half: 6-Step ETL Workflow Horizontal Strip
    etl_flow = [
        ("1. Ingestion", "Folder connectors for CSV/Excel", NAVY),
        ("2. Null Treatment", "Impute Area & 247 missing Billed", TEAL_DARK),
        ("3. Anomaly Clamping", "Clamp 15 negatives to 0.0 kWh", RED),
        ("4. Text Cleansing", "Trim whitespace & composite keys", GOLD_DARK),
        ("5. Type Casting", "Uint32 counts & Float64 units", NAVY_MED),
        ("6. Validation", "Enforce (Billed <= Total) logic", GREEN_DARK)
    ]
    for idx, (step_title, step_desc, step_col) in enumerate(etl_flow):
        x_e = Inches(0.5) + idx * Inches(2.08)
        card(sl, x_e, Inches(1.45), Inches(1.95), Inches(1.3))
        rect(sl, x_e, Inches(1.45), Inches(1.95), Inches(0.32), step_col, radius=0.04)
        txt(sl, step_title, x_e + Inches(0.08), Inches(1.5), Inches(1.8), Inches(0.22),
            size=Pt(9.5), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")
        txt(sl, step_desc, x_e + Inches(0.08), Inches(1.85), Inches(1.8), Inches(0.85),
            size=Pt(8.5), color=DARK_TEXT, align=PP_ALIGN.CENTER, name="Calibri")

    # Bottom Half: 4 Feature Engineered Domain Indicators
    features = [
        ("Billing Efficiency Rate (%)",
         "billing_rate_pct = (BILLDSERVICES / TOTSERVICES) * 100",
         "Measures revenue capture efficiency. Identifies non-functioning meters, unmetered services, or delayed reading cycles. State benchmark is 95%.",
         GREEN),
        ("Consumption Intensity per Service",
         "units_per_billed_service = UNITS / BILLDSERVICES",
         "Normalizes consumption volume against active connections. Removes size bias, exposing true living-standard intensity across areas.",
         TEAL),
        ("Load Utilization Ratio (kWh / kW)",
         "units_per_kw_load = UNITS / LOAD",
         "Measures how intensively connected power capacity is utilized. Sectors > 250 kWh/kW flag potential transformer strain or commercial misuse.",
         GOLD_DARK),
        ("Unbilled Services Gap",
         "unbilled_services = TOTSERVICES - BILLDSERVICES",
         "Direct head-count of unregistered or unbilled domestic connections (total 648,960). Powers prioritized field recovery squads.",
         RED)
    ]
    for idx, (f_name, f_formula, f_rationale, f_col) in enumerate(features):
        x_pos = Inches(0.5) + (idx % 2) * Inches(6.3)
        y_pos = Inches(2.95) + (idx // 2) * Inches(2.0)
        card(sl, x_pos, y_pos, Inches(6.0), Inches(1.85))
        rect(sl, x_pos, y_pos, Inches(6.0), Inches(0.35), f_col, radius=0.04)
        txt(sl, f_name, x_pos + Inches(0.18), y_pos + Inches(0.06), Inches(5.6), Inches(0.24),
            size=Pt(11), bold=True, color=WHITE, name="Calibri")
        rect(sl, x_pos + Inches(0.18), y_pos + Inches(0.42), Inches(5.6), Inches(0.35), ROW_ALT, radius=0.04)
        txt(sl, f_formula, x_pos + Inches(0.25), y_pos + Inches(0.48), Inches(5.4), Inches(0.24),
            size=Pt(9), bold=True, color=NAVY, name="Courier New")
        txt(sl, f_rationale, x_pos + Inches(0.18), y_pos + Inches(0.85), Inches(5.6), Inches(0.9),
            size=Pt(9), color=DARK_TEXT, name="Calibri")

    footer(sl, 5)

    # ==============================================================================
    # SLIDE 6: PILLAR 5 — VARIABLE CONNECTIONS: CORRELATION & REGRESSION
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Pillar 5: Variable Connections", "Pearson Correlation Analysis & Regression Dynamics",
               "Investigating mathematical dependencies between connected load, connections, and energy demand")

    # Left: Seaborn Load vs Units Regression Plot
    add_image_card(sl, Inches(0.5), Inches(1.45), Inches(6.2), Inches(5.5),
                   "load_vs_units.png", "Figure: Connected Load (kW) vs Consumed Energy (MWh) with Seaborn Regplot")

    # Right: Correlation Matrix Table & Physical Takeaways
    headers = ["Variable Pair", "r", "Strength", "Operational Interpretation"]
    rows = [
        ["LOAD ↔ UNITS", "+0.946", "Very Strong", "Connected load capacity directly governs consumption."],
        ["SERVICES ↔ BILLED", "+0.995", "Near Perfect", "Registered connections dictate billed services."],
        ["BILLED ↔ UNITS", "+0.884", "Strong", "Larger customer base reliably scales energy usage."],
        ["SERVICES ↔ LOAD", "+0.763", "Moderate-High", "Sanctioned capacity scales with household counts."]
    ]
    col_w = [Inches(1.8), Inches(0.8), Inches(1.1), Inches(2.1)]
    create_table(sl, Inches(6.9), Inches(1.45), Inches(5.9), Inches(2.4), headers, rows, col_w)

    bullet_card(sl, Inches(6.9), Inches(4.05), Inches(5.9), Inches(2.9),
                "Regression Slope & Association vs Causation", [
                    ("Linear Regression Trend (Red Line)", "Empirically validates strong linear scaling. For every 1,000 kW of connected capacity, monthly energy increases by ~90 MWh."),
                    ("Dense Baseload Band", "The primary cluster between 0–2,000 kW displays tightly bounded, highly predictable residential behavior."),
                    ("High Vertical Scatter", "Sectors located high above the line reflect intense air conditioning or commercial activity billed on domestic tariffs."),
                    ("Correlation != Causation", "Load provides the physical pipeline, but summer temperature and affluence drive actual kilowatt-hour draw.")
                ], NAVY)

    footer(sl, 6)

    # ==============================================================================
    # SLIDE 7: PILLAR 6 — BASIC VISUALIZATION USING EXCEL
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Pillar 6: Excel Visualizations", "Microsoft Excel Analytics: Pivot Tables & Visual Dashboards",
               "Rapid ad-hoc exploratory analytics, slicing, clustered columns, and tool capability boundaries")

    # Left: Excel Visualization Image (Clustered Column + Donut)
    add_image_card(sl, Inches(0.5), Inches(1.45), Inches(6.2), Inches(5.5),
                   "excel_visualization.png", "Figure: Microsoft Excel Clustered Column & 2D Donut Visualizations")

    # Right Top: Pivot Table Features
    bullet_card(sl, Inches(6.9), Inches(1.45), Inches(5.9), Inches(2.6),
                "Excel Pivot Table & Slicer Workflows", [
                    ("Hierarchical Grouping", "Rows: CIRCLE > DIVISION > SECTION. Values: Sum of Units, Sum of Load, Average Units per Service."),
                    ("Interactive Slicers", "Added slicers for CIRCLE and DIVISION for instant zero-code field station filtering."),
                    ("Calculated Fields", "Created 'Billing Efficiency %' field within the Pivot Table to benchmark sections against 95%."),
                    ("Conditional Formatting", "Applied soft red data bars and cell highlights to flag low-billing sections in seconds.")
                ], EXCEL_GREEN)

    # Right Bottom: Excel Strengths vs Python Limitations
    bullet_card(sl, Inches(6.9), Inches(4.2), Inches(5.9), Inches(2.75),
                "Excel Strengths vs Tool Boundaries", [
                    ("Excel Advantages", "Accessible to all utility managers; standard for tariff regulatory filings and rapid ad-hoc slicing."),
                    ("Scale Constraints", "Performance degrades on multi-million row smart meter telemetry datasets."),
                    ("Lack of Advanced Statistics", "Cannot natively generate KDE distribution curves, regression confidence intervals, or automated heatmaps."),
                    ("Reproducibility Risk", "Vulnerable to accidental cell formula overwrites; Python automation guarantees 100% reproducible pipelines.")
                ], NAVY)

    footer(sl, 7)

    # ==============================================================================
    # SLIDE 8: PILLAR 7 — PYTHON VISUALIZATIONS: INFRASTRUCTURE & HOTSPOTS
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Pillar 7: Python Visualizations", "Dual-Axis Infrastructure Analysis & Top 20 Demand Hotspots",
               "High-resolution Matplotlib & Seaborn plots mapping circle capacity and neighborhood epicenters")

    # Left: Circle Consumption & Load Dual-Axis Plot
    add_image_card(sl, Inches(0.5), Inches(1.45), Inches(6.0), Inches(4.3),
                   "circle_consumption.png", "Figure 1: Energy Consumed (GWh) & Sanctioned Load (MW) by Circle (Matplotlib & Seaborn)")

    # Right: Top 20 Areas Seaborn Barplot
    add_image_card(sl, Inches(6.8), Inches(1.45), Inches(6.0), Inches(4.3),
                   "top20_areas.png", "Figure 2: Top 20 High-Demand Areas in TG-SPDCL (Seaborn Barplot)")

    # Bottom Synthesis Banner
    bullet_card(sl, Inches(0.5), Inches(5.9), Inches(12.33), Inches(1.05),
                "Key Infrastructure Takeaways", [
                    ("Proportional Grid Symmetry", "Consumption (cyan bars, GWh) and Connected Load (purple bars, MW) match with near-perfect alignment across all 20 circles, validating infrastructure sizing."),
                    ("Top Demand Epicenters", "Hydernagar (2.32M kWh), Moosapet (1.81M), and Vinayaknagar (1.47M) lead state-wide consumption, requiring targeted transformer upgrades.")
                ], NAVY)

    footer(sl, 8)

    # ==============================================================================
    # SLIDE 9: PILLAR 7 — PYTHON VISUALIZATIONS: EFFICIENCY & DISTRIBUTIONS
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Pillar 7: Python Visualizations", "Section Billing Efficiency Benchmark & Distribution Metrics",
               "Evaluating circle compliance against the 95% target, household KDE spread, and feature correlations")

    # Left: Billing Efficiency Barplot
    add_image_card(sl, Inches(0.5), Inches(1.45), Inches(6.0), Inches(4.3),
                   "billing_efficiency.png", "Figure 3: Circle Billing Efficiency (%) vs 95% Benchmark (Seaborn Viridis)")

    # Right: Seaborn Correlation Heatmap
    add_image_card(sl, Inches(6.8), Inches(1.45), Inches(6.0), Inches(4.3),
                   "seaborn_correlation_heatmap.png", "Figure 4: Seaborn Indicator Correlation Heatmap (Spectral Palette)")

    # Bottom Synthesis Banner
    bullet_card(sl, Inches(0.5), Inches(5.9), Inches(12.33), Inches(1.05),
                "Distribution & Compliance Synthesis", [
                    ("5 Circles Below 95% Benchmark", "Saroornagar and rural district circles show significant unbilled gaps requiring prioritized field inspection squads."),
                    ("Strong Multi-Collinearity Confirmed", "Spectral heatmap verifies heavy interdependence between Load, TotServices, and Units (+0.95), reinforcing regression findings.")
                ], NAVY)

    footer(sl, 9)

    # ==============================================================================
    # SLIDE 10: PILLAR 8 — FINAL DECISION: STRATEGIC ACTION PLAN & ROADMAP
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Pillar 8: Final Decision", "Strategic Decision Framework: 4 Pillars & Phased Implementation",
               "Data-driven operational interventions to maximize utility revenue and grid stability")

    # Top: 4 Strategic Pillars
    action_pillars = [
        ("Pillar 1: Infrastructure Augmentation",
         "Prioritize Top 15% Areas",
         "Target distribution transformer (DTR) capacity doubling on the 2,452 'Very Large' areas driving 50.5% of demand (Hydernagar, Moosapet, Kukatpally).",
         TEAL_DARK),
        ("Pillar 2: Revenue Protection & Audits",
         "Resolve 648k Unbilled Services",
         "Deploy field squads to audit the 247 zero-billed sections and 15 negative-unit areas. Enforce billing compliance in circles lagging below 90%.",
         RED),
        ("Pillar 3: Smart Grid Modernization",
         "Automated Smart Metering (AMI)",
         "Roll out IoT smart meters in high-intensity areas (>250 kWh/kW) and resolve area name ambiguity via GIS composite key tagging.",
         NAVY_MED),
        ("Pillar 4: Tariff & Peak Management",
         "Time-of-Day (ToD) Pricing",
         "Introduce dynamic ToD tariffs for domestic connections > 10 kW to shave peak evening residential loads and incentivize off-peak charging.",
         GOLD_DARK)
    ]
    for idx, (p_title, p_sub, p_desc, p_col) in enumerate(action_pillars):
        x_p = Inches(0.5) + idx * Inches(3.15)
        card(sl, x_p, Inches(1.45), Inches(3.0), Inches(2.55))
        rect(sl, x_p, Inches(1.45), Inches(3.0), Inches(0.35), p_col, radius=0.04)
        txt(sl, p_title, x_p + Inches(0.1), Inches(1.5), Inches(2.8), Inches(0.24),
            size=Pt(10), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")
        txt(sl, f"Focus: {p_sub}", x_p + Inches(0.12), Inches(1.85), Inches(2.75), Inches(0.22),
            size=Pt(8.5), bold=True, color=p_col, align=PP_ALIGN.CENTER, name="Calibri")
        txt(sl, p_desc, x_p + Inches(0.15), Inches(2.1), Inches(2.7), Inches(1.8),
            size=Pt(9), color=DARK_TEXT, name="Calibri")

    # Bottom: 3-Phase Implementation Roadmap Flow
    phases = [
        ("Phase 1: Immediate Sprint (Days 1–30)", "Audit 247 zero-billed sections, verify 15 negative-unit areas, and deploy GIS keys.", RED),
        ("Phase 2: Targeted Deployment (Days 31–180)", "Upgrade transformers in top 15% areas, install 250k smart meters, launch live web dashboard.", GOLD_DARK),
        ("Phase 3: Utility Transformation (Year 1)", "Achieve 97%+ billing efficiency across all circles, rollout ToD tariffs, and predictive ML failure models.", GREEN_DARK)
    ]
    for idx, (ph_name, ph_desc, ph_col) in enumerate(phases):
        x_ph = Inches(0.5) + idx * Inches(4.2)
        card(sl, x_ph, Inches(4.15), Inches(3.95), Inches(2.75))
        rect(sl, x_ph, Inches(4.15), Inches(3.95), Inches(0.35), ph_col, radius=0.04)
        txt(sl, ph_name, x_ph + Inches(0.1), Inches(4.2), Inches(3.75), Inches(0.24),
            size=Pt(10.5), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")
        txt(sl, ph_desc, x_ph + Inches(0.18), Inches(4.6), Inches(3.6), Inches(2.1),
            size=Pt(9.5), color=DARK_TEXT, name="Calibri")

    footer(sl, 10)

    # ==============================================================================
    # SLIDE 11: PILLAR 8 — BUSINESS IMPACT, FINANCIAL ROI & CONCLUSION
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Pillar 8: Final Decision", "Business Impact, Financial ROI & Project Conclusion",
               "Quantifiable returns, system reliability improvements, and summary conclusions")

    # 4 Financial ROI Badges (Top Half)
    roi_metrics = [
        ("+4.5% Billing Efficiency", "Unlocks ₹85–110 Crore annually by converting unbilled connections to active revenue.", TEAL),
        ("-35% Transformer Failures", "Targeted DTR load doubling prevents summer overheating across top 15% sectors.", GREEN),
        ("100% Feeder Traceability", "Complete elimination of repeated area name ambiguity via GIS composite key tagging.", GOLD),
        ("Real-Time Telemetry", "Interactive monitoring portal enables daily loss tracking rather than 30-day manual delays.", RGBColor(0x8b, 0x5c, 0xf6))
    ]
    for idx, (r_title, r_desc, r_col) in enumerate(roi_metrics):
        x_r = Inches(0.5) + idx * Inches(3.15)
        card(sl, x_r, Inches(1.45), Inches(3.0), Inches(2.0), bg=NAVY, border=r_col)
        rect(sl, x_r, Inches(1.45), Inches(3.0), Inches(0.06), r_col)
        txt(sl, r_title, x_r + Inches(0.15), Inches(1.6), Inches(2.7), Inches(0.4),
            size=Pt(14), bold=True, color=r_col, name="Calibri")
        txt(sl, r_desc, x_r + Inches(0.15), Inches(2.1), Inches(2.7), Inches(1.2),
            size=Pt(9.5), color=WHITE, name="Calibri")

    # Bottom Half: Project Synthesis & Q&A Card
    bullet_card(sl, Inches(0.5), Inches(3.65), Inches(12.33), Inches(3.3),
                "Executive Summary & Analytical Conclusions", [
                    ("Complete 8-Pillar Alignment", "Demonstrated end-to-end data lifecycle from raw multi-file ingestion to statistical modeling, dual-axis visualization, and strategic policy decisions."),
                    ("High Urban Demand Concentration", "Over 50.5% of electricity is driven by just 15.4% of neighborhood sectors in the Greater Hyderabad ring, requiring focused capital allocation."),
                    ("Strong Predictive Reliability", "Connected Load (kW) and Units (kWh) show near-perfect linear correlation (r = 0.946), enabling accurate forward demand forecasting."),
                    ("Actionable Revenue Gains", "Addressing the 648,960 unbilled connections offers immediate high-ROI financial recovery for TG-SPDCL."),
                    ("Academic & Professional Quality", "Completed by Research Team: D. Chander (2503A52214), K. Arjun (2503A52235), P. Sachin (2503A52240). Open for Q&A.")
                ], NAVY_MED)

    footer(sl, 11)

    # ─── SAVE PRESENTATION ────────────────────────────────────────────────────────
    out_file = "TG_SPDCL_Electricity_Analysis_v4.pptx"
    prs.save(out_file)
    print(f"Successfully generated Version 4 presentation: '{out_file}' (11 Slides Total - Strictly under 12)")

if __name__ == "__main__":
    build_presentation_v4()
