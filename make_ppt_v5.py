# -*- coding: utf-8 -*-
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

"""
TG-SPDCL Electricity Consumption & Monitoring
Academic Presentation Generator — Version 5 (v5)
Ultra-High-Density Executive Edition — EXACTLY 7 SLIDES
Synthesizes the Complete 8-Pillar Classroom Syllabus:
  Slide 1: Title & Executive Launchpad (Team, Headline KPIs & 8-Pillar Framework)
  Slide 2: Pillars 1 & 2 — Project Explanation, Grid Hierarchy & Data Exploration
  Slide 3: Pillar 3 — Basic Statistics, Distribution Skewness & The 15%/50.5% Pareto Rule
  Slide 4: Pillars 4 & 5 — Data Wrangling, Feature Engineering & Correlation Regression
  Slide 5: Pillar 6 — Basic Visualization using Excel (Pivot Tables & Dashboards)
  Slide 6: Pillar 7 — Advanced Visualization using Python (Matplotlib & Seaborn Multi-Panel)
  Slide 7: Pillar 8 — Final Decision, 4-Pillar Action Plan, Roadmap & Financial ROI

Authors: D. Chander, K. Arjun, P. Sachin
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# ─── DESIGN TOKENS (PREMIUM CONSULTING PALETTE) ───────────────────────────────
NAVY        = RGBColor(0x0a, 0x19, 0x3b)   # Deep Midnight Navy
NAVY_MED    = RGBColor(0x16, 0x2b, 0x5a)   # Medium Navy
TEAL        = RGBColor(0x00, 0xb4, 0xd8)   # Electric Teal
TEAL_DARK   = RGBColor(0x00, 0x77, 0x91)   # Deep Cyan / Teal
GOLD        = RGBColor(0xf5, 0x9e, 0x0b)   # Warm Gold / Amber Accent
GOLD_DARK   = RGBColor(0xd9, 0x77, 0x06)
GREEN       = RGBColor(0x10, 0xb9, 0x81)   # Emerald Success Green
GREEN_DARK  = RGBColor(0x04, 0x78, 0x57)
RED         = RGBColor(0xef, 0x44, 0x44)   # Alert / Anomaly Red
SLATE       = RGBColor(0x64, 0x74, 0x8b)   # Secondary Text
SLATE_LIGHT = RGBColor(0x94, 0xa3, 0xb8)
DARK_TEXT   = RGBColor(0x0f, 0x17, 0x2a)   # Charcoal High-Contrast Body
LIGHT_BG    = RGBColor(0xf8, 0xfa, 0xfc)   # Clean Slide Canvas
WHITE       = RGBColor(0xff, 0xff, 0xff)
BORDER_GRAY = RGBColor(0xe2, 0xe8, 0xf0)   # Clean Border
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
    """Ultra-clean header banner with syllabus pillar badge."""
    rect(slide, 0, 0, SW, Inches(1.15), NAVY)
    rect(slide, 0, Inches(1.15), SW, Inches(0.04), GOLD)
    
    # Pillar Tag Pill
    rect(slide, Inches(0.5), Inches(0.12), Inches(3.4), Inches(0.25), TEAL_DARK, radius=0.4)
    txt(slide, pillar_tag.upper(), Inches(0.55), Inches(0.14), Inches(3.3), Inches(0.22),
        size=Pt(8.5), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")
    
    # Title
    txt(slide, title, Inches(0.5), Inches(0.40), Inches(9.8), Inches(0.42),
        size=Pt(20), bold=True, color=WHITE, name="Calibri")
    
    # Subtitle
    if subtitle:
        txt(slide, subtitle, Inches(0.5), Inches(0.84), Inches(9.8), Inches(0.25),
            size=Pt(10.5), color=SLATE_LIGHT, italic=True, name="Calibri")
            
    # Right-side branding
    txt(slide, "TG-SPDCL | DAV Project v5", Inches(9.8), Inches(0.40), Inches(3.0), Inches(0.25),
        size=Pt(9.5), bold=True, color=TEAL, align=PP_ALIGN.RIGHT, name="Calibri")
    txt(slide, "Executive 7-Slide Master Edition", Inches(9.8), Inches(0.68), Inches(3.0), Inches(0.24),
        size=Pt(8.5), color=SLATE_LIGHT, align=PP_ALIGN.RIGHT, name="Calibri")

def footer(slide, slide_num, total=7):
    rect(slide, 0, SH - Inches(0.35), SW, Inches(0.35), NAVY)
    txt(slide, "TG-SPDCL Domestic Electricity Consumption & Monitoring  |  Data Analytics & Visualization (DAV)  |  July 2026",
        Inches(0.5), SH - Inches(0.30), Inches(10.5), Inches(0.25),
        size=Pt(8.5), color=SLATE_LIGHT, name="Calibri")
    txt(slide, f"Slide {slide_num} of {total}", Inches(11.2), SH - Inches(0.30), Inches(1.6), Inches(0.25),
        size=Pt(9), bold=True, color=WHITE, align=PP_ALIGN.RIGHT, name="Calibri")

def card(slide, l, t, w, h, bg=WHITE, border=BORDER_GRAY, radius=0.04):
    return rect(slide, l, t, w, h, bg, border, Pt(0.75), radius=radius)

def bullet_card(slide, l, t, w, h, title, items, header_color=NAVY):
    card(slide, l, t, w, h)
    rect(slide, l, t, w, Inches(0.35), header_color, radius=0.04)
    txt(slide, title, l + Inches(0.15), t + Inches(0.06), w - Inches(0.3), Inches(0.24),
        size=Pt(11), bold=True, color=WHITE, name="Calibri")
    
    box = slide.shapes.add_textbox(l + Inches(0.15), t + Inches(0.38), w - Inches(0.3), h - Inches(0.42))
    tf = box.text_frame
    tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(2.5)
        p.space_after = Pt(1.5)
        if isinstance(it, tuple):
            r1 = p.add_run()
            r1.text = f"•  {it[0]}: "
            r1.font.bold = True
            r1.font.size = Pt(9)
            r1.font.color.rgb = DARK_TEXT
            r1.font.name = "Calibri"
            
            r2 = p.add_run()
            r2.text = it[1]
            r2.font.bold = False
            r2.font.size = Pt(9)
            r2.font.color.rgb = SLATE
            r2.font.name = "Calibri"
        else:
            r = p.add_run()
            r.text = f"•  {it}"
            r.font.size = Pt(9)
            r.font.color.rgb = DARK_TEXT
            r.font.name = "Calibri"

def add_image_card(slide, l, t, w, h, img_path, caption=""):
    card(slide, l, t, w, h)
    if os.path.exists(img_path):
        pad = Inches(0.06)
        img_h = h - Inches(0.28) if caption else h - (pad * 2)
        slide.shapes.add_picture(img_path, l + pad, t + pad, w - (pad * 2), img_h)
        if caption:
            txt(slide, caption, l + pad, t + h - Inches(0.25), w - (pad * 2), Inches(0.22),
                size=Pt(8), bold=True, color=NAVY_MED, align=PP_ALIGN.CENTER, name="Calibri")
    else:
        txt(slide, f"[Image File Not Found: {img_path}]", l + Inches(0.2), t + h/2 - Inches(0.2),
            w - Inches(0.4), Inches(0.4), size=Pt(9.5), bold=True, color=RED, align=PP_ALIGN.CENTER)

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
        run.font.size = Pt(9)
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
            run.font.size = Pt(8.5)
            run.font.name = "Calibri"
            if c_idx == 0:
                run.font.bold = True
                run.font.color.rgb = NAVY
            else:
                run.font.color.rgb = DARK_TEXT

# ─── BUILD PRESENTATION (EXACTLY 7 SLIDES) ────────────────────────────────────

def build_presentation_v5():
    prs = Presentation()
    prs.slide_width  = SW
    prs.slide_height = SH
    BLANK = prs.slide_layouts[6]

    # ==============================================================================
    # SLIDE 1: TITLE & EXECUTIVE LAUNCHPAD
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, NAVY)
    
    # Modern abstract framing
    rect(sl, Inches(8.3), Inches(0), Inches(5.033), SH, RGBColor(0x13, 0x24, 0x4e))
    rect(sl, Inches(9.5), Inches(1.2), Inches(3.833), Inches(5.4), RGBColor(0x18, 0x2c, 0x5e))
    rect(sl, Inches(0.8), Inches(3.2), Inches(4.5), Inches(0.06), GOLD)

    # Category Pill
    rect(sl, Inches(0.8), Inches(0.85), Inches(3.4), Inches(0.35), TEAL, radius=0.4)
    txt(sl, "DATA ANALYTICS & VISUALIZATION (DAV)", Inches(0.9), Inches(0.90), Inches(3.2), Inches(0.25),
        size=Pt(9.5), bold=True, color=WHITE, name="Calibri")

    # Presentation Titles
    txt(sl, "Electricity Consumption", Inches(0.8), Inches(1.30), Inches(7.5), Inches(0.85),
        size=Pt(38), bold=True, color=WHITE, name="Calibri")
    txt(sl, "& Monitoring Analysis", Inches(0.8), Inches(2.10), Inches(7.5), Inches(0.85),
        size=Pt(38), bold=True, color=TEAL, name="Calibri")
    txt(sl, "TG-SPDCL Domestic Consumer Dataset (July 2026) | Complete 8-Pillar Syllabus Synthesis",
        Inches(0.8), Inches(3.35), Inches(7.5), Inches(0.35),
        size=Pt(12.5), color=SLATE_LIGHT, italic=True, name="Calibri")

    # Team Box
    card(sl, Inches(0.8), Inches(3.95), Inches(7.0), Inches(2.7), bg=RGBColor(0x13, 0x24, 0x4e), border=TEAL_DARK)
    rect(sl, Inches(0.8), Inches(3.95), Inches(7.0), Inches(0.05), TEAL)
    txt(sl, "PROJECT RESEARCH TEAM", Inches(1.0), Inches(4.10), Inches(4.0), Inches(0.24),
        size=Pt(10.5), bold=True, color=GOLD, name="Calibri")

    team = [
        ("D. Chander", "Roll: 2503A52214", "Data Cleaning, Wrangling & Power Query ETL"),
        ("K. Arjun",   "Roll: 2503A52235", "Seaborn Visualizations, Correlations & Heatmaps"),
        ("P. Sachin",  "Roll: 2503A52240", "Statistical Modeling, Pareto Analysis & Final Decision")
    ]
    for idx, (m_name, m_roll, m_role) in enumerate(team):
        y_pos = Inches(4.45) + idx * Inches(0.65)
        txt(sl, m_name, Inches(1.0), y_pos, Inches(2.0), Inches(0.3),
            size=Pt(14), bold=True, color=WHITE, name="Calibri")
        txt(sl, m_roll, Inches(3.0), y_pos + Inches(0.02), Inches(1.7), Inches(0.25),
            size=Pt(10.5), bold=True, color=TEAL, name="Calibri")
        txt(sl, m_role, Inches(4.7), y_pos + Inches(0.02), Inches(2.9), Inches(0.25),
            size=Pt(9), color=SLATE_LIGHT, name="Calibri")

    # 4 Key Stat Badges
    stat_boxes = [
        ("15,947", "SECTOR RECORDS", "20 Circles | 49 Divisions | 476 Sections", TEAL),
        ("8.87M", "TOTAL SERVICES", "8.22M Active Billed Connections", GOLD),
        ("1,245.5 GWh", "ENERGY CONSUMED", "1.245 Billion kWh Total Monthly Demand", GREEN),
        ("92.68%", "BILLING EFFICIENCY", "648,960 Unbilled Services Identified", RGBColor(0xf9, 0x73, 0x16))
    ]
    for idx, (val, title_lbl, desc_lbl, accent) in enumerate(stat_boxes):
        y_b = Inches(1.05) + idx * Inches(1.4)
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
    # SLIDE 2: PILLARS 1 & 2 — PROJECT EXPLANATION & DATA EXPLORATION
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Pillars 1 & 2: Project Explanation & Exploration",
               "Utility Context, 5-Tier Grid Hierarchy & Complete Column Data Dictionary",
               "Structural mapping of TG-SPDCL domestic distribution network and data quality audit")

    # Left Column: Pillar 1 (Context, Hierarchy & Challenges)
    bullet_card(sl, Inches(0.5), Inches(1.35), Inches(5.8), Inches(2.7),
                "Pillar 1: Utility Context & Grid Hierarchy", [
                    ("Organization", "Telangana Southern Power Distribution Company Limited (TG-SPDCL)."),
                    ("5-Tier Network Flow", "20 Circles → 49 Divisions → 131 Subdivisions → 476 Sections → 15,947 Areas."),
                    ("Consumer Category", "Domestic Category 1 (`CATCODE: 1`, `CATDESC: DOMESTIC`) for July 2026 cycle."),
                    ("Grid Relevance", "Domestic power accounts for >40% of urban peak load during summer/monsoon.")
                ], NAVY)

    bullet_card(sl, Inches(0.5), Inches(4.2), Inches(5.8), Inches(2.8),
                "Operational Problem Statement", [
                    ("High Unbilled Gap", "648,960 services unbilled in a single month (~7.32% unbilled rate), leaking revenue."),
                    ("Severe Load Concentration", "Heavy air-conditioning clusters in western Hyderabad strain local distribution transformers."),
                    ("Data Hygiene Challenges", "Missing records, negative unit reversals, and repeated naming across sections.")
                ], RED)

    # Right Column: Pillar 2 (Data Dictionary & Quality Audit)
    headers = ["Feature", "Type", "Role / Meaning in Analysis"]
    rows = [
        ["CIRCLE", "String", "Administrative zone (20 unique zones)"],
        ["DIVISION", "String", "Operational district division (49 unique)"],
        ["SUBDIVISION", "String", "Field maintenance engineering zone (131)"],
        ["SECTION", "String", "Local billing station & line office (476)"],
        ["AREA", "String", "Micro neighborhood colony (15,947 rows)"],
        ["CATDESC/CODE", "Cat/Int", "Category 1 DOMESTIC tariff identifier"],
        ["TOTSERVICES", "Integer", "Registered domestic consumer connections"],
        ["BILLDSERVICES", "Float/Int", "Active billed connections in July 2026"],
        ["UNITS", "Float", "Metered energy consumption in kWh"],
        ["LOAD", "Float", "Sanctioned connected load capacity in kW"]
    ]
    col_w = [Inches(1.5), Inches(1.0), Inches(3.8)]
    create_table(sl, Inches(6.6), Inches(1.35), Inches(6.3), Inches(3.2), headers, rows, col_w)

    bullet_card(sl, Inches(6.6), Inches(4.7), Inches(6.233), Inches(2.3),
                "Pillar 2: Data Quality Audit & Anomaly Profiling", [
                    ("247 Missing BILLDSERVICES", "Discovery: 100% of these 247 rows have exactly 0 units (dormant/uncommissioned lines)."),
                    ("47 Missing AREA Names", "Handled via string imputation as 'UNKNOWN AREA' without losing geographic hierarchy."),
                    ("15 Negative UNIT Records", "Min -1,703 kWh (billing reversal adjustments); clamped to 0.0 in analytics pipeline."),
                    ("1,999 Repeated Area Names", "'MAIN ROAD' & 'GANDHI NAGAR' repeat; composite keys (`section_area`) enforce identity.")
                ], GOLD_DARK)

    footer(sl, 2)

    # ==============================================================================
    # SLIDE 3: PILLAR 3 — BASIC STATISTICS & THE PARETO RULE
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Pillar 3: Basic Statistics",
               "Statistical Dispersion, Regional Rankings & The 15% / 50.5% Pareto Rule",
               "Parametric vs median skewness, urban-rural lifestyle gap, and demand concentration")

    # Top: Five-Number Summary Table
    headers = ["Statistical Indicator", "Total Services", "Billed Services", "Consumption (kWh)", "Connected Load (kW)"]
    rows = [
        ["Mean (Average)", "556.1 services", "523.0 services", "78,103.6 kWh", "864.04 kW"],
        ["Std Deviation", "626.2", "605.8", "112,930.4", "1,529.57"],
        ["Minimum", "1", "1", "-1,703.0 (Anomalous)", "0.12 kW"],
        ["25th Percentile (Q1)", "152.0", "138.0", "11,312.0 kWh", "86.00 kW"],
        ["50th Percentile (Median)", "382.0", "345.0", "36,826.0 kWh", "316.00 kW"],
        ["75th Percentile (Q3)", "735.5", "690.0", "102,405.5 kWh", "1,058.15 kW"],
        ["Maximum", "12,182", "11,886", "2,319,144.0 kWh", "33,617.69 kW"]
    ]
    col_w = [Inches(2.5), Inches(2.3), Inches(2.3), Inches(2.7), Inches(2.5)]
    create_table(sl, Inches(0.5), Inches(1.35), Inches(12.33), Inches(2.7), headers, rows, col_w)

    # Bottom Left: The 15% / 50.5% Pareto Principle
    bullet_card(sl, Inches(0.5), Inches(4.25), Inches(6.0), Inches(2.75),
                "The Pareto Principle: Area-Size Tier Segmentation", [
                    ("Very Large Areas (>1,000 services)", "2,452 areas (15.4% of total) account for 50.5% (629.0 GWh) of all electricity consumed."),
                    ("Large Areas (501–1,000 services)", "3,810 areas (23.9% count) account for 24.2% (301.4 GWh) of total demand."),
                    ("Medium Areas (201–500 services)", "5,120 areas (32.1% count) account for 18.1% (225.4 GWh)."),
                    ("Small Areas (1–200 services)", "4,565 areas (28.6% count) account for only 7.2% (89.7 GWh)."),
                    ("Utility Takeaway", "Focusing grid upgrades on the top 15% sectors stabilizes more than half the entire utility load.")
                ], RED)

    # Bottom Right: Geographic Disparity & 3x Lifestyle Gap
    bullet_card(sl, Inches(6.8), Inches(4.25), Inches(6.033), Inches(2.75),
                "Geographic Rankings & The 3× Lifestyle Consumption Gap", [
                    ("Top 5 Circles Consume 50.6%", "Medchal (141.7M), Habsiguda (138.1M), Cybercity (125.4M), Hyd Central (112.5M), Saroornagar (112.3M)."),
                    ("Bottom 5 Circles < 1.5% Each", "Gadwal (10.0M), Wanaparthy (10.1M), Nagarkurnool (13.1M), Medak (15.9M), Vikarabad (17.9M)."),
                    ("Lifestyle Consumption Disparity", "Banjara Hills (240 kWh) & Cybercity (219 kWh) vs Rural Districts (75–83 kWh) — a 3× living standard gap."),
                    ("Extreme Right Skewness", "Max sector Hydernagar (2.32M kWh) is 63× greater than the state median (36.8k kWh).")
                ], NAVY)

    footer(sl, 3)

    # ==============================================================================
    # SLIDE 4: PILLARS 4 & 5 — DATA WRANGLING & VARIABLE CONNECTIONS
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Pillars 4 & 5: Wrangling & Variable Connections",
               "6-Step ETL Pipeline, Feature Engineering & Load-Units Regression Dynamics",
               "Transforming raw billing data and modeling physical dependencies between load and consumption")

    # Left Column: Pillar 4 (ETL Flow & Feature Formulas)
    bullet_card(sl, Inches(0.5), Inches(1.35), Inches(5.8), Inches(2.3),
                "Pillar 4: 6-Step Data Wrangling (ETL) Pipeline", [
                    ("1. Ingestion & Merging", "Multi-file CSV/Excel loading via Power Query & Pandas read_csv."),
                    ("2. Null Treatment", "Impute 47 area names; replace 247 missing billed services with 0."),
                    ("3. Anomaly Clamping", "Clamp 15 negative units to 0.0; clamp minimum load to 0.01 kW."),
                    ("4. Standardize & Validate", "Trim whitespace, cast types (uint32/float64), enforce (Billed <= Total).")
                ], TEAL_DARK)

    bullet_card(sl, Inches(0.5), Inches(3.8), Inches(5.8), Inches(3.2),
                "Feature Engineered Domain Indicators", [
                    ("Billing Efficiency Rate (%)", "`billing_rate_pct = (BILLDSERVICES / TOTSERVICES) * 100` — Tracks revenue capture against 95% target."),
                    ("Consumption per Connection", "`units_per_billed_service = UNITS / BILLDSERVICES` — Normalizes household living standards without size bias."),
                    ("Load Utilization Ratio", "`units_per_kw_load = UNITS / LOAD` — Measures transformer strain (>250 kWh/kW flags commercial misuse)."),
                    ("Unbilled Services Gap", "`unbilled_services = TOTSERVICES - BILLDSERVICES` — Identifies 648,960 non-billing accounts for field inspection.")
                ], GREEN_DARK)

    # Right Column: Pillar 5 (Regression Image + Correlation Table)
    add_image_card(sl, Inches(6.6), Inches(1.35), Inches(6.233), Inches(3.5),
                   "load_vs_units.png", "Figure: Connected Load (kW) vs Energy Consumed (MWh) with Seaborn Regplot")

    headers = ["Variable Pair", "Correlation (r)", "Operational Significance"]
    rows = [
        ["LOAD ↔ UNITS", "+0.946 (Very Strong)", "Connected load capacity directly governs electricity draw."],
        ["SERVICES ↔ BILLED", "+0.995 (Near Perfect)", "Registered connections dictate billed services linearly."],
        ["BILLED ↔ UNITS", "+0.884 (Strong)", "Customer base volume reliably scales gross energy usage."],
        ["SERVICES ↔ LOAD", "+0.763 (Moderate-High)", "Sanctioned capacity expands proportionally with connections."]
    ]
    col_w = [Inches(1.8), Inches(1.8), Inches(2.6)]
    create_table(sl, Inches(6.6), Inches(5.0), Inches(6.233), Inches(2.0), headers, rows, col_w)

    footer(sl, 4)

    # ==============================================================================
    # SLIDE 5: PILLAR 6 — BASIC VISUALIZATION USING EXCEL
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Pillar 6: Excel Visualizations",
               "Microsoft Excel Analytics: Pivot Tables, Slicers & Chart Visualizations",
               "Rapid ad-hoc exploratory analytics, slicing, clustered columns, and tool capability boundaries")

    # Left: Embedded Excel Chart (Clustered Column + Donut)
    add_image_card(sl, Inches(0.5), Inches(1.35), Inches(6.0), Inches(5.65),
                   "excel_visualization.png", "Figure: Microsoft Excel Clustered Column & 2D Donut Visualizations")

    # Right Top: Pivot Table Architecture
    bullet_card(sl, Inches(6.8), Inches(1.35), Inches(6.033), Inches(2.7),
                "Excel Pivot Table & Slicer Workflows", [
                    ("Hierarchical Drill-Down", "Configured Pivot Table rows as CIRCLE > DIVISION > SECTION for fast regional inspection."),
                    ("Aggregated Indicators", "Sum of UNITS, Sum of LOAD, and Average Units per Billed Service."),
                    ("Interactive Slicers", "Added slicers for CIRCLE and DIVISION for instant zero-code field station filtering."),
                    ("Calculated Fields", "Created 'Billing Efficiency %' field within the Pivot Table to benchmark sections against 95%."),
                    ("Conditional Formatting", "Applied soft red data bars and cell highlights to flag low-billing sections in seconds.")
                ], EXCEL_GREEN)

    # Right Bottom: Excel Strengths vs Python Limitations
    bullet_card(sl, Inches(6.8), Inches(4.2), Inches(6.033), Inches(2.8),
                "Excel Practical Advantages vs Analytical Boundaries", [
                    ("Excel Advantages", "Accessible to all utility managers; standard for tariff regulatory filings and rapid ad-hoc slicing."),
                    ("Scale Constraints", "Performance degrades on multi-million row smart meter telemetry datasets."),
                    ("Lack of Advanced Statistics", "Cannot natively generate KDE distribution curves, regression confidence intervals, or automated heatmaps."),
                    ("Reproducibility Risk", "Vulnerable to accidental cell formula overwrites; Python automation guarantees 100% reproducible pipelines.")
                ], NAVY)

    footer(sl, 5)

    # ==============================================================================
    # SLIDE 6: PILLAR 7 — PYTHON VISUALIZATIONS (MATPLOTLIB & SEABORN MULTI-PANEL)
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Pillar 7: Python Visualizations",
               "Advanced Python Data Stack: Dual-Axis, Hotspots, Benchmark & Heatmaps",
               "Four publication-grade Matplotlib & Seaborn figures demonstrating complete grid telemetry")

    # 4-Panel Grid of Visualizations
    add_image_card(sl, Inches(0.5), Inches(1.35), Inches(5.9), Inches(2.55),
                   "circle_consumption.png", "Fig 1: Circle Consumption (GWh) & Sanctioned Load (MW)")

    add_image_card(sl, Inches(6.8), Inches(1.35), Inches(6.033), Inches(2.55),
                   "top20_areas.png", "Fig 2: Top 20 Consumption Epicenters (Seaborn Barplot)")

    add_image_card(sl, Inches(0.5), Inches(4.05), Inches(5.9), Inches(2.55),
                   "billing_efficiency.png", "Fig 3: Circle Billing Efficiency vs 95% Benchmark")

    add_image_card(sl, Inches(6.8), Inches(4.05), Inches(6.033), Inches(2.55),
                   "seaborn_correlation_heatmap.png", "Fig 4: Seaborn Indicator Correlation Heatmap")

    # Center Takeaway Strip
    rect(sl, Inches(0.5), Inches(6.65), Inches(12.333), Inches(0.42), NAVY, radius=0.04)
    txt(sl, "SYNTHESIS: Near-perfect load/consumption symmetry | Hydernagar (2.32M kWh) leads demand | 5 circles lag below 95% target | Multi-collinear load coupling (+0.95)",
        Inches(0.6), Inches(6.72), Inches(12.1), Inches(0.26),
        size=Pt(9), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")

    footer(sl, 6)

    # ==============================================================================
    # SLIDE 7: PILLAR 8 — FINAL DECISION, ROADMAP & BUSINESS ROI
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Pillar 8: Final Decision",
               "Strategic 4-Pillar Action Plan, 3-Phase Implementation & Measurable Business ROI",
               "Operational utility interventions delivering revenue recovery, transformer longevity, and project conclusion")

    # Top: 4 Strategic Pillars
    action_pillars = [
        ("Pillar 1: Infrastructure", "Target Top 15% Areas",
         "Double transformer capacity in the 2,452 heavy areas driving 50.5% load (Hydernagar, Moosapet).", TEAL_DARK),
        ("Pillar 2: Revenue Audits", "Resolve 648k Unbilled",
         "Deploy field squads to audit 247 zero-billed sections and enforce compliance in circles < 90%.", RED),
        ("Pillar 3: Smart Grid AMI", "IoT Telemetry",
         "Install smart meters in high-intensity areas (>250 kWh/kW); resolve duplicate names via GIS keys.", NAVY_MED),
        ("Pillar 4: Tariff Reform", "Time-of-Day (ToD)",
         "Implement dynamic ToD pricing for connections > 10 kW to flatten evening residential peaks.", GOLD_DARK)
    ]
    for idx, (p_title, p_sub, p_desc, p_col) in enumerate(action_pillars):
        x_p = Inches(0.5) + idx * Inches(3.15)
        card(sl, x_p, Inches(1.35), Inches(3.0), Inches(2.2))
        rect(sl, x_p, Inches(1.35), Inches(3.0), Inches(0.32), p_col, radius=0.04)
        txt(sl, p_title, x_p + Inches(0.1), Inches(1.40), Inches(2.8), Inches(0.22),
            size=Pt(9.5), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")
        txt(sl, f"Focus: {p_sub}", x_p + Inches(0.12), Inches(1.72), Inches(2.75), Inches(0.20),
            size=Pt(8.5), bold=True, color=p_col, align=PP_ALIGN.CENTER, name="Calibri")
        txt(sl, p_desc, x_p + Inches(0.15), Inches(1.95), Inches(2.7), Inches(1.5),
            size=Pt(8.5), color=DARK_TEXT, name="Calibri")

    # Middle: 3-Phase Roadmap Flow
    phases = [
        ("Phase 1: 30-Day Sprint", "Audit 247 zero-billed sections, verify 15 negative units, deploy GIS primary keys.", RED),
        ("Phase 2: 90-Day Expansion", "Upgrade transformers in top 15% areas, install 250k smart meters, launch live dashboard.", GOLD_DARK),
        ("Phase 3: 1-Year Transformation", "Achieve 97%+ billing efficiency across all circles, launch ToD tariffs, and predictive ML.", GREEN_DARK)
    ]
    for idx, (ph_name, ph_desc, ph_col) in enumerate(phases):
        x_ph = Inches(0.5) + idx * Inches(4.2)
        card(sl, x_ph, Inches(3.70), Inches(3.95), Inches(1.35))
        rect(sl, x_ph, Inches(3.70), Inches(3.95), Inches(0.28), ph_col, radius=0.04)
        txt(sl, ph_name, x_ph + Inches(0.1), Inches(3.74), Inches(3.75), Inches(0.20),
            size=Pt(9.5), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")
        txt(sl, ph_desc, x_ph + Inches(0.15), Inches(4.02), Inches(3.65), Inches(0.95),
            size=Pt(8.5), color=DARK_TEXT, name="Calibri")

    # Bottom: Financial ROI Badges & Project Sign-off
    card(sl, Inches(0.5), Inches(5.20), Inches(12.333), Inches(1.85), bg=NAVY, border=GOLD)
    rect(sl, Inches(0.5), Inches(5.20), Inches(12.333), Inches(0.05), GOLD)
    txt(sl, "PROJECTED BUSINESS IMPACT, FINANCIAL ROI & PROJECT CONCLUSION", Inches(0.8), Inches(5.32), Inches(11.5), Inches(0.22),
        size=Pt(10), bold=True, color=GOLD, name="Calibri")

    roi_stats = [
        ("+4.5% Billing Efficiency", "Unlocks ₹85–110 Crore annually by converting unbilled connections to active revenue.", TEAL),
        ("-35% Transformer Failures", "Targeted DTR load doubling prevents summer overloads in top 15% demand sectors.", GREEN),
        ("100% Feeder Traceability", "Complete elimination of repeated area name ambiguity via GIS composite key tagging.", GOLD),
        ("Research Team Sign-off", "D. Chander (2503A52214), K. Arjun (2503A52235), P. Sachin (2503A52240). Open for Q&A.", WHITE)
    ]
    for idx, (roi_title, roi_sub, roi_color) in enumerate(roi_stats):
        x_r = Inches(0.8) + idx * Inches(2.95)
        txt(sl, roi_title, x_r, Inches(5.60), Inches(2.8), Inches(0.26),
            size=Pt(10.5), bold=True, color=roi_color, name="Calibri")
        txt(sl, roi_sub, x_r, Inches(5.90), Inches(2.75), Inches(1.0),
            size=Pt(8), color=RGBColor(0xcf, 0xd8, 0xdc), name="Calibri")

    footer(sl, 7)

    # ─── SAVE PRESENTATION ────────────────────────────────────────────────────────
    out_file = "TG_SPDCL_Electricity_Analysis_v5.pptx"
    prs.save(out_file)
    print(f"Successfully generated Version 5 presentation: '{out_file}' (EXACTLY 7 SLIDES TOTAL)")

if __name__ == "__main__":
    build_presentation_v5()
