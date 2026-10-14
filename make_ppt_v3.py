# -*- coding: utf-8 -*-
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

"""
TG-SPDCL Electricity Consumption & Monitoring
Academic Presentation Generator — Version 3 (v3)
Structured strictly according to the 8-Part Classroom Syllabus:
  1. Project Explanation
  2. Data Exploration
  3. Basic Statistics
  4. Data Wrangling
  5. Variable Connections (Correlation)
  6. Basic Visualization using Excel
  7. Advanced Visualization using Python (Matplotlib, Seaborn, etc.)
  8. Final Decision

Authors: D. Chander, K. Arjun, P. Sachin
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_CONNECTOR_TYPE
from pptx.enum.shapes import MSO_SHAPE

# ─── COLOR PALETTE (EXECUTIVE CONSULTING THEME) ───────────────────────────────
NAVY        = RGBColor(0x0d, 0x1b, 0x3e)   # Primary brand dark navy
NAVY_MED    = RGBColor(0x1a, 0x31, 0x6b)   # Secondary navy
TEAL        = RGBColor(0x00, 0xb4, 0xd8)   # Accent electric teal
TEAL_DARK   = RGBColor(0x00, 0x77, 0x91)   # Deep cyan / teal
GOLD        = RGBColor(0xf5, 0xa6, 0x23)   # Amber / gold accent
GOLD_DARK   = RGBColor(0xd9, 0x77, 0x06)   # Deep amber
GREEN       = RGBColor(0x10, 0xb9, 0x81)   # Success emerald green
GREEN_DARK  = RGBColor(0x05, 0x96, 0x69)
RED         = RGBColor(0xef, 0x44, 0x44)   # Alert / anomaly red
SLATE       = RGBColor(0x64, 0x74, 0x8b)   # Secondary muted text
DARK_TEXT   = RGBColor(0x0f, 0x17, 0x2a)   # High-contrast body text
LIGHT_BG    = RGBColor(0xf8, 0xfa, 0xfc)   # Slide background
WHITE       = RGBColor(0xff, 0xff, 0xff)   # Card background
BORDER_GRAY = RGBColor(0xe2, 0xe8, 0xf0)   # Clean subtle card border
ROW_ALT     = RGBColor(0xf1, 0xf5, 0xf9)   # Alternate table row

SW = Inches(13.333)
SH = Inches(7.5)

# ─── CORE LAYOUT UTILITIES ────────────────────────────────────────────────────

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

def header_bar(slide, section_tag, title, subtitle=None):
    """Standardized top banner with colored section tag pill, title, and metadata."""
    rect(slide, 0, 0, SW, Inches(1.3), NAVY)
    rect(slide, 0, Inches(1.3), SW, Inches(0.04), GOLD)
    
    # Section Tag Pill
    rect(slide, Inches(0.5), Inches(0.15), Inches(2.8), Inches(0.28), TEAL_DARK, radius=0.4)
    txt(slide, section_tag.upper(), Inches(0.55), Inches(0.18), Inches(2.7), Inches(0.24),
        size=Pt(9), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")
    
    # Main Title
    txt(slide, title, Inches(0.5), Inches(0.46), Inches(9.8), Inches(0.48),
        size=Pt(22), bold=True, color=WHITE, name="Calibri")
    
    # Subtitle
    if subtitle:
        txt(slide, subtitle, Inches(0.5), Inches(0.94), Inches(9.8), Inches(0.32),
            size=Pt(11.5), color=RGBColor(0x94, 0xa3, 0xb8), italic=True, name="Calibri")
            
    # Right-hand branding
    txt(slide, "TG-SPDCL | DAV Project v3", Inches(9.8), Inches(0.48), Inches(3.0), Inches(0.3),
        size=Pt(10), bold=True, color=TEAL, align=PP_ALIGN.RIGHT, name="Calibri")
    txt(slide, "Classroom Syllabus Structure", Inches(9.8), Inches(0.8), Inches(3.0), Inches(0.3),
        size=Pt(9), color=RGBColor(0x94, 0xa3, 0xb8), align=PP_ALIGN.RIGHT, name="Calibri")

def footer(slide, slide_num, total=20):
    rect(slide, 0, SH - Inches(0.38), SW, Inches(0.38), NAVY)
    txt(slide, "TG-SPDCL Electricity Consumption & Monitoring  |  Data Analytics & Visualization (DAV)  |  July 2026",
        Inches(0.5), SH - Inches(0.32), Inches(10.5), Inches(0.28),
        size=Pt(9), color=RGBColor(0x94, 0xa3, 0xb8), name="Calibri")
    txt(slide, f"Slide {slide_num} of {total}", Inches(11.2), SH - Inches(0.32), Inches(1.6), Inches(0.28),
        size=Pt(9), bold=True, color=WHITE, align=PP_ALIGN.RIGHT, name="Calibri")

def card(slide, l, t, w, h, bg=WHITE, border=BORDER_GRAY, radius=0.04):
    """Clean white card with subtle border and optional rounded corners."""
    return rect(slide, l, t, w, h, bg, border, Pt(0.75), radius=radius)

def stat_badge(slide, l, t, w, h, label, value, sub, accent=TEAL):
    """KPI summary card with vertical accent stripe."""
    card(slide, l, t, w, h)
    rect(slide, l, t, Inches(0.08), h, accent, radius=0.04)
    txt(slide, label.upper(), l + Inches(0.18), t + Inches(0.1), w - Inches(0.25), Inches(0.22),
        size=Pt(8.5), bold=True, color=SLATE, name="Calibri")
    txt(slide, value, l + Inches(0.18), t + Inches(0.32), w - Inches(0.25), Inches(0.45),
        size=Pt(20), bold=True, color=NAVY, name="Calibri")
    txt(slide, sub, l + Inches(0.18), t + h - Inches(0.3), w - Inches(0.25), Inches(0.25),
        size=Pt(9), color=SLATE, name="Calibri")

def bullet_card(slide, l, t, w, h, title, items, header_color=NAVY):
    """Framed card with bullet points and bold headers."""
    card(slide, l, t, w, h)
    rect(slide, l, t, w, Inches(0.42), header_color, radius=0.04)
    txt(slide, title, l + Inches(0.2), t + Inches(0.08), w - Inches(0.4), Inches(0.3),
        size=Pt(12), bold=True, color=WHITE, name="Calibri")
    
    box = slide.shapes.add_textbox(l + Inches(0.2), t + Inches(0.48), w - Inches(0.4), h - Inches(0.55))
    tf = box.text_frame
    tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(4)
        p.space_after = Pt(2)
        if isinstance(it, tuple):
            r1 = p.add_run()
            r1.text = f"•  {it[0]}: "
            r1.font.bold = True
            r1.font.size = Pt(10.5)
            r1.font.color.rgb = DARK_TEXT
            r1.font.name = "Calibri"
            
            r2 = p.add_run()
            r2.text = it[1]
            r2.font.bold = False
            r2.font.size = Pt(10)
            r2.font.color.rgb = SLATE
            r2.font.name = "Calibri"
        else:
            r = p.add_run()
            r.text = f"•  {it}"
            r.font.size = Pt(10)
            r.font.color.rgb = DARK_TEXT
            r.font.name = "Calibri"

def add_image_card(slide, l, t, w, h, img_path, caption=""):
    """Embed an image with a clean border, white card background, and subtle caption."""
    card(slide, l, t, w, h)
    if os.path.exists(img_path):
        pad = Inches(0.08)
        img_h = h - Inches(0.35) if caption else h - (pad * 2)
        slide.shapes.add_picture(img_path, l + pad, t + pad, w - (pad * 2), img_h)
        if caption:
            txt(slide, caption, l + pad, t + h - Inches(0.3), w - (pad * 2), Inches(0.25),
                size=Pt(9), bold=True, color=NAVY_MED, align=PP_ALIGN.CENTER, name="Calibri")
    else:
        txt(slide, f"[Image File Not Found: {img_path}]", l + Inches(0.2), t + h/2 - Inches(0.2),
            w - Inches(0.4), Inches(0.4), size=Pt(11), bold=True, color=RED, align=PP_ALIGN.CENTER)

def create_table(slide, l, t, w, h, headers, rows, col_widths=None):
    """Draw a clean, professional corporate data table."""
    num_rows = len(rows) + 1
    num_cols = len(headers)
    table_shape = slide.shapes.add_table(num_rows, num_cols, l, t, w, h)
    table = table_shape.table
    
    if col_widths:
        for idx, width in enumerate(col_widths):
            table.columns[idx].width = width
            
    # Format Headers
    for c_idx, head in enumerate(headers):
        cell = table.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = head
        run.font.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = WHITE
        run.font.name = "Calibri"
        
    # Format Rows
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
            run.font.size = Pt(9.5)
            run.font.name = "Calibri"
            if c_idx == 0:
                run.font.bold = True
                run.font.color.rgb = NAVY
            else:
                run.font.color.rgb = DARK_TEXT

# ─── PRESENTATION BUILDER ──────────────────────────────────────────────────────

def build_presentation():
    prs = Presentation()
    prs.slide_width  = SW
    prs.slide_height = SH
    BLANK = prs.slide_layouts[6]

    # ==============================================================================
    # SLIDE 1: TITLE SLIDE
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl, NAVY)
    
    # Abstract geometric shapes
    rect(sl, Inches(8.5), Inches(0), Inches(4.833), SH, RGBColor(0x13, 0x24, 0x4e))
    rect(sl, Inches(9.8), Inches(1.5), Inches(3.533), Inches(5.0), RGBColor(0x19, 0x2e, 0x62))
    rect(sl, Inches(0.8), Inches(3.4), Inches(4.5), Inches(0.06), GOLD)

    # Category Pill
    rect(sl, Inches(0.8), Inches(1.0), Inches(3.5), Inches(0.38), TEAL, radius=0.4)
    txt(sl, "ACADEMIC DATA ANALYTICS PROJECT (DAV)", Inches(0.9), Inches(1.06), Inches(3.3), Inches(0.26),
        size=Pt(10), bold=True, color=WHITE, name="Calibri")

    # Presentation Titles
    txt(sl, "Electricity Consumption", Inches(0.8), Inches(1.5), Inches(7.5), Inches(0.85),
        size=Pt(40), bold=True, color=WHITE, name="Calibri")
    txt(sl, "& Monitoring Analysis", Inches(0.8), Inches(2.35), Inches(7.5), Inches(0.85),
        size=Pt(40), bold=True, color=TEAL, name="Calibri")
    txt(sl, "Comprehensive Analytical Report & Decision Framework (TG-SPDCL July 2026)",
        Inches(0.8), Inches(3.55), Inches(7.5), Inches(0.4),
        size=Pt(13.5), color=RGBColor(0x94, 0xa3, 0xb8), italic=True, name="Calibri")

    # Team Members Box
    card(sl, Inches(0.8), Inches(4.3), Inches(6.8), Inches(2.4), bg=RGBColor(0x13, 0x24, 0x4e), border=TEAL_DARK)
    rect(sl, Inches(0.8), Inches(4.3), Inches(6.8), Inches(0.06), TEAL)
    txt(sl, "PROJECT RESEARCH TEAM", Inches(1.0), Inches(4.45), Inches(4.0), Inches(0.25),
        size=Pt(11), bold=True, color=GOLD, name="Calibri")

    team = [
        ("D. Chander", "Roll No: 2503A52214", "Data Cleaning, Wrangling & Power Query"),
        ("K. Arjun",   "Roll No: 2503A52235", "Seaborn Visualizations & Correlation Analysis"),
        ("P. Sachin",  "Roll No: 2503A52240", "Statistical Modeling & Final Decision Framework")
    ]
    for idx, (m_name, m_roll, m_role) in enumerate(team):
        y_pos = Inches(4.8) + idx * Inches(0.58)
        txt(sl, m_name, Inches(1.0), y_pos, Inches(2.0), Inches(0.3),
            size=Pt(14), bold=True, color=WHITE, name="Calibri")
        txt(sl, m_roll, Inches(3.0), y_pos + Inches(0.02), Inches(1.8), Inches(0.25),
            size=Pt(10.5), bold=True, color=TEAL, name="Calibri")
        txt(sl, m_role, Inches(4.8), y_pos + Inches(0.02), Inches(2.6), Inches(0.25),
            size=Pt(9.5), color=RGBColor(0x94, 0xa3, 0xb8), name="Calibri")

    # Headline Stat Badges (Right Column)
    stat_boxes = [
        ("15,947", "SECTOR RECORDS", "20 Circles & 49 Divisions", TEAL),
        ("8.87M", "TOTAL SERVICES", "8.22M Billed Connections", GOLD),
        ("1,245.5 GWh", "ENERGY CONSUMED", "1.245 Billion kWh Total", GREEN),
        ("92.68%", "BILLING EFFICIENCY", "648,960 Unbilled Services", RGBColor(0xf9, 0x73, 0x16))
    ]
    for idx, (val, title_lbl, desc_lbl, accent) in enumerate(stat_boxes):
        y_b = Inches(1.2) + idx * Inches(1.35)
        card(sl, Inches(8.9), y_b, Inches(3.8), Inches(1.15), bg=RGBColor(0x0f, 0x17, 0x2a), border=accent)
        rect(sl, Inches(8.9), y_b, Inches(0.08), Inches(1.15), accent)
        txt(sl, val, Inches(9.1), y_b + Inches(0.12), Inches(3.4), Inches(0.42),
            size=Pt(22), bold=True, color=WHITE, name="Calibri")
        txt(sl, title_lbl, Inches(9.1), y_b + Inches(0.55), Inches(3.4), Inches(0.25),
            size=Pt(9), bold=True, color=accent, name="Calibri")
        txt(sl, desc_lbl, Inches(9.1), y_b + Inches(0.8), Inches(3.4), Inches(0.22),
            size=Pt(8.5), color=RGBColor(0x94, 0xa3, 0xb8), name="Calibri")

    footer(sl, 1)

    # ==============================================================================
    # SLIDE 2: THE 8-PILLAR PRESENTATION FRAMEWORK (Direct Syllabus Mapping)
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Executive Agenda", "The 8-Pillar Data Analytics Lifecycle Framework",
               "Structured end-to-end curriculum roadmap for TG-SPDCL electricity monitoring")

    syllabus_steps = [
        ("01", "Project Explanation", "Domain context, problem statement, TG-SPDCL objectives, and scope boundaries.", TEAL),
        ("02", "Data Exploration", "Column dictionary, variable categorization, missing value audit, and anomaly detection.", NAVY_MED),
        ("03", "Basic Statistics", "Five-number summary, high skewness, regional rankings, and the 15%/50.5% Pareto rule.", GOLD_DARK),
        ("04", "Data Wrangling", "6-stage Power Query/Pandas ETL pipeline, cleaning rules, and feature-engineered metrics.", GREEN_DARK),
        ("05", "Variable Connections", "Pearson correlation matrix, load vs units coupling (r=0.95), and regression trendlines.", RGBColor(0x8b, 0x5c, 0xf6)),
        ("06", "Excel Visualizations", "Multi-level Pivot Tables, Slicers, Clustered Columns, Donut charts, and tool boundaries.", RGBColor(0x02, 0x84, 0xc7)),
        ("07", "Python Visualizations", "Matplotlib & Seaborn advanced plots: Dual-axis, Top 20 areas, KDE curves, and Heatmap.", RGBColor(0xd9, 0x46, 0xef)),
        ("08", "Final Decision", "Strategic 4-pillar utility action plan, 3-phase roadmap, and measurable revenue ROI.", RGBColor(0xef, 0x44, 0x44)),
    ]

    for idx, (num, p_title, p_desc, col_accent) in enumerate(syllabus_steps):
        row = idx // 4
        col = idx % 4
        x_c = Inches(0.5) + col * Inches(3.15)
        y_c = Inches(1.6) + row * Inches(2.65)
        
        card(sl, x_c, y_c, Inches(3.0), Inches(2.45))
        rect(sl, x_c, y_c, Inches(3.0), Inches(0.48), col_accent, radius=0.04)
        
        txt(sl, f"{num}. {p_title}", x_c + Inches(0.12), y_c + Inches(0.1), Inches(2.75), Inches(0.3),
            size=Pt(12), bold=True, color=WHITE, name="Calibri")
            
        txt(sl, p_desc, x_c + Inches(0.15), y_c + Inches(0.65), Inches(2.7), Inches(1.15),
            size=Pt(10), color=DARK_TEXT, name="Calibri")
            
        rect(sl, x_c + Inches(0.15), y_c + Inches(1.95), Inches(2.7), Inches(0.32), ROW_ALT, radius=0.04)
        txt(sl, f"Phase {idx+1} Deliverable", x_c + Inches(0.2), y_c + Inches(2.0), Inches(2.6), Inches(0.22),
            size=Pt(8.5), bold=True, color=col_accent, align=PP_ALIGN.CENTER, name="Calibri")

    footer(sl, 2)

    # ==============================================================================
    # SLIDE 3: PART 1 — PROJECT EXPLANATION: CONTEXT & OBJECTIVES
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Part 1: Project Explanation", "Project Motivation, Problem Statement & Core Objectives",
               "Understanding the critical operational challenges of electricity distribution in Telangana")

    # 3 Strategic Dimension Cards
    bullet_card(sl, Inches(0.5), Inches(1.6), Inches(3.9), Inches(5.2),
                "1. Utility Background (TG-SPDCL)", [
                    ("Organization", "Telangana Southern Power Distribution Company Limited."),
                    ("Operational Footprint", "Serves 15 southern districts of Telangana including Greater Hyderabad."),
                    ("Consumer Category", "Domestic Category 1 (Households, apartments, residential colonies)."),
                    ("Period Covered", "Monthly billing cycle of July 2026."),
                    ("Data Scale", "Over 8.86 million domestic connections across 15,947 neighborhood sectors."),
                    ("Grid Relevance", "Domestic load constitutes over 40% of urban peak power demand during summer/monsoon.")
                ], NAVY)

    bullet_card(sl, Inches(4.7), Inches(1.6), Inches(3.9), Inches(5.2),
                "2. Core Problem Statement", [
                    ("Rapid Urban Load Growth", "Hyper-expansion in western Hyderabad creates severe distribution bottlenecks."),
                    ("High Unbilled Gap", "648,960 services unbilled in a single month (~7.32% unbilled rate)."),
                    ("Severe Load Concentration", "Massive disparity between high-rise urban clusters and rural tribal hamlets."),
                    ("Revenue Leakage", "Unmetered connections, delayed meter readings, and billing adjustments."),
                    ("Transformer Overheating", "High-intensity areas drawing over 250 kWh/kW causing local outages."),
                    ("Data Quality Gaps", "Missing records, negative unit anomalies, and unstandardized area nomenclatures.")
                ], RED)

    bullet_card(sl, Inches(8.9), Inches(1.6), Inches(3.9), Inches(5.2),
                "3. Key Project Objectives", [
                    ("Consumption Profiling", "Map total and per-connection energy usage across 20 administrative circles."),
                    ("Loss Identification", "Pinpoint divisions and sections lagging below the 95% billing efficiency threshold."),
                    ("Correlation Modeling", "Quantify the mathematical dependency between Sanctioned Load and Units Consumed."),
                    ("Comparative Analytics", "Demonstrate complementary analytical workflows between Excel and Python."),
                    ("Decision Support System", "Deliver data-driven interventions for transformer augmentation and billing audits."),
                    ("Interactive Dashboard", "Deploy a real-time web portal for operational grid monitoring.")
                ], TEAL_DARK)

    footer(sl, 3)

    # ==============================================================================
    # SLIDE 4: PART 1 — PROJECT EXPLANATION: HIERARCHY & SCOPE
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Part 1: Project Explanation", "Administrative Grid Hierarchy & Operational Scope",
               "Structural mapping of the 5-tier distribution network and operational boundaries")

    # Flowchart of Hierarchy
    levels = [
        ("20 Circles", "Apex Administrative Zone", "Medchal, Cybercity, Habsiguda, Banjara Hills, etc.", NAVY),
        ("49 Divisions", "Operational Management", "Kukatpally, Mehdipatnam, Saroornagar, Kondapur, etc.", NAVY_MED),
        ("131 Subdivisions", "Field Engineering Clusters", "Miyapur, Qutbullapur, Uppal, Gachibowli, etc.", TEAL_DARK),
        ("476 Sections", "Local Billing & Line Stations", "Billing offices responsible for meter reading cycles.", TEAL),
        ("15,947 Areas", "Granular Consumption Sectors", "Neighborhoods, colonies, streets, and rural hamlets.", GREEN)
    ]

    for idx, (l_title, l_role, l_desc, l_color) in enumerate(levels):
        x_p = Inches(0.5) + idx * Inches(2.52)
        card(sl, x_p, Inches(1.6), Inches(2.35), Inches(2.2))
        rect(sl, x_p, Inches(1.6), Inches(2.35), Inches(0.4), l_color, radius=0.04)
        txt(sl, l_title, x_p + Inches(0.1), Inches(1.68), Inches(2.15), Inches(0.25),
            size=Pt(11), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")
        txt(sl, l_role, x_p + Inches(0.1), Inches(2.08), Inches(2.15), Inches(0.25),
            size=Pt(9), bold=True, color=l_color, align=PP_ALIGN.CENTER, name="Calibri")
        txt(sl, l_desc, x_p + Inches(0.1), Inches(2.38), Inches(2.15), Inches(1.3),
            size=Pt(8.5), color=SLATE, align=PP_ALIGN.CENTER, name="Calibri")

    # Scope and Boundaries Section below
    bullet_card(sl, Inches(0.5), Inches(4.1), Inches(6.0), Inches(2.7),
                "Dataset Boundaries & Strengths", [
                    ("Comprehensive Domestic Coverage", "Complete census of all registered domestic consumer connections in July 2026."),
                    ("Granular Hierarchy", "Enables macro executive summaries (Circle) down to micro root-cause investigations (Section/Area)."),
                    ("Tri-Metric Accounting", "Simultaneously records connection counts, sanctioned power capacity (kW), and metered units (kWh)."),
                    ("Realistic Operational Nuance", "Contains true-to-life field artifacts including billing corrections and inactive connections.")
                ], NAVY)

    bullet_card(sl, Inches(6.8), Inches(4.1), Inches(6.0), Inches(2.7),
                "Analytical Limitations & Caveats", [
                    ("Domestic Only (Cat 1)", "Excludes commercial, industrial, and agricultural loads; does not describe total grid demand."),
                    ("Single Monthly Snapshot", "July 2026 cross-section precludes longitudinal seasonal time-series trend forecasting."),
                    ("No Socioeconomic Predictors", "Lacks household income, building square-footage, weather/temperature, and appliance counts."),
                    ("Aggregated Feeder Units", "Data is aggregated at the area level rather than individual smart meter intervals.")
                ], GOLD_DARK)

    footer(sl, 4)

    # ==============================================================================
    # SLIDE 5: PART 2 — DATA EXPLORATION: COLUMN DICTIONARY & SCOPE
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Part 2: Data Exploration", "Dataset Architecture & Complete Column Data Dictionary",
               "15,947 rows × 11 active features categorizing domestic electricity distribution")

    headers = ["Field Name", "Data Type", "Category", "Description / Analytical Purpose", "Sample Value"]
    rows = [
        ["CIRCLE", "String", "Categorical", "Highest administrative division in TG-SPDCL", "MEDCHAL"],
        ["DIVISION", "String", "Categorical", "Operational district subdivision managing sections", "KUKATPALLY"],
        ["SUBDIVISION", "String", "Categorical", "Engineering maintenance cluster", "MIYAPUR"],
        ["SECTION", "String", "Categorical", "Local billing station responsible for meter collection", "HYDERNAGAR"],
        ["AREA", "String", "Categorical", "Micro geographic neighborhood or settlement name", "VIVEKANANDA NAGAR"],
        ["CATDESC", "String", "Categorical", "Consumer category description (DOMESTIC)", "DOMESTIC"],
        ["CATCODE", "Integer", "Numerical", "Tariff category classification identifier (Code = 1)", "1"],
        ["TOTSERVICES", "Integer", "Numerical", "Total sanctioned/registered domestic connections", "1,240"],
        ["BILLDSERVICES", "Float/Int", "Numerical", "Active connections billed during July 2026 cycle", "1,185"],
        ["UNITS", "Float", "Numerical", "Total metered electrical consumption in kWh (Units)", "184,520.0"],
        ["LOAD", "Float", "Numerical", "Total sanctioned connected load capacity in Kilowatts", "1,850.50 kW"]
    ]
    col_w = [Inches(1.5), Inches(1.1), Inches(1.3), Inches(5.8), Inches(2.3)]
    create_table(sl, Inches(0.5), Inches(1.6), Inches(12.33), Inches(5.2), headers, rows, col_w)

    footer(sl, 5)

    # ==============================================================================
    # SLIDE 6: PART 2 — DATA EXPLORATION: QUALITY AUDIT & ANOMALIES
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Part 2: Data Exploration", "Data Quality Audit, Missing Value Profiling & Anomalies",
               "Rigorous diagnostic screening revealing null patterns, negative corrections, and duplicates")

    # 4 Anomaly Diagnostic Cards
    bullet_card(sl, Inches(0.5), Inches(1.6), Inches(2.9), Inches(5.2),
                "Missing Values (Nulls)", [
                    ("BILLDSERVICES Nulls", "247 records have missing billed service values."),
                    ("The Zero-Unit Pattern", "Crucially, 100% of these 247 rows have exactly 0 metered units."),
                    ("Root Cause Diagnosis", "Represents uncommissioned, inactive, or unmetered service feeders."),
                    ("AREA Name Nulls", "47 records feature missing area name strings."),
                    ("Remediation Strategy", "Impute AREA as 'UNKNOWN AREA'; impute missing billed services to 0 or total.")
                ], RED)

    bullet_card(sl, Inches(3.6), Inches(1.6), Inches(2.9), Inches(5.2),
                "Negative Consumption (15)", [
                    ("Observed Minimum", "-1,703.0 kWh observed in rural agricultural-border feeders."),
                    ("Total Occurrences", "15 records feature negative unit values."),
                    ("Accounting Meaning", "Utility accounting adjustments resulting from prior over-billing reversals or faulty meters."),
                    ("Impact Assessment", "Distorts local efficiency if left untreated; clamped to 0.0 in analytical pipelines.")
                ], GOLD_DARK)

    bullet_card(sl, Inches(6.7), Inches(1.6), Inches(2.9), Inches(5.2),
                "Zero Consumption (344)", [
                    ("Zero-Unit Areas", "344 records record exactly 0.0 kWh consumed despite active connections."),
                    ("Geographic Clustering", "Concentrated in rural tribal settlements ('THANDAS') in Nalgonda & Suryapet."),
                    ("Operational Nature", "Valid real-world state representing vacant homes, seasonal migration, or disconnected lines."),
                    ("Retained in Model", "Kept intact to avoid artificially inflating average consumption benchmarks.")
                ], TEAL_DARK)

    bullet_card(sl, Inches(9.8), Inches(1.6), Inches(3.0), Inches(5.2),
                "Nomenclature Duplication", [
                    ("Duplicate Area Labels", "1,999 area names repeat across different sections."),
                    ("Common Examples", "'MAIN ROAD', 'GANDHI NAGAR', 'SC COLONY', 'HOUSING BOARD'."),
                    ("Data Hygiene Lesson", "Area names cannot serve as primary keys without concatenating Section and Division IDs."),
                    ("Power Query Cleaning", "Standardized by trimming whitespace, uppercase formatting, and composite indexing.")
                ], NAVY)

    footer(sl, 6)

    # ==============================================================================
    # SLIDE 7: PART 3 — BASIC STATISTICS: FIVE-NUMBER SUMMARY & SKEWNESS
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Part 3: Basic Statistics", "Five-Number Statistical Summary & Extreme Skewness",
               "Parametric vs non-parametric metrics demonstrating extreme right-skewed distribution")

    # Summary Statistics Table
    headers = ["Metric / Variable", "Registered Services", "Billed Services", "Consumption (kWh)", "Connected Load (kW)"]
    rows = [
        ["Count (Records)", "15,947", "15,700 (247 null)", "15,947", "15,947"],
        ["Mean (Average)", "556.1 services", "523.0 services", "78,103.6 kWh", "864.04 kW"],
        ["Standard Deviation", "626.2", "605.8", "112,930.4", "1,529.57"],
        ["Minimum", "1", "1", "-1,703.0 (Anomalous)", "0.12 kW"],
        ["25th Percentile (Q1)", "152.0", "138.0", "11,312.0 kWh", "86.00 kW"],
        ["50th Percentile (Median)", "382.0", "345.0", "36,826.0 kWh", "316.00 kW"],
        ["75th Percentile (Q3)", "735.5", "690.0", "102,405.5 kWh", "1,058.15 kW"],
        ["Interquartile Range (IQR)", "583.5", "552.0", "91,093.5 kWh", "972.15 kW"],
        ["Maximum", "12,182", "11,886", "2,319,144.0 kWh", "33,617.69 kW"]
    ]
    col_w = [Inches(2.5), Inches(2.2), Inches(2.2), Inches(2.6), Inches(2.4)]
    create_table(sl, Inches(0.5), Inches(1.6), Inches(11.9), Inches(3.2), headers, rows, col_w)

    # Statistical Interpretation Cards below
    bullet_card(sl, Inches(0.5), Inches(5.0), Inches(5.8), Inches(1.8),
                "Parametric Skewness Analysis", [
                    ("Mean vs Median Disparity", "Mean consumption (78.1k kWh) is more than 2.1× higher than Median (36.8k kWh)."),
                    ("Long Right Tail", "Max value of 2.32M kWh (Hydernagar) is 63× greater than the median sector."),
                    ("Robust Statistics Rule", "Median and IQR must be reported alongside Mean to prevent distorted benchmarks.")
                ], NAVY)

    bullet_card(sl, Inches(6.6), Inches(5.0), Inches(6.2), Inches(1.8),
                "1.5 × IQR Outlier Detection (Tukey's Rule)", [
                    ("Upper Threshold for Units", "Q3 + (1.5 × IQR) = 102,405 + 136,640 = 239,045 kWh."),
                    ("Flagged Sectors", "Approximately 1,120 areas exceed this threshold as heavy load epicenters."),
                    ("Actionable Protocol", "Investigate extreme values against feeder ratings before treating as errors.")
                ], GOLD_DARK)

    footer(sl, 7)

    # ==============================================================================
    # SLIDE 8: PART 3 — BASIC STATISTICS: REGIONAL & CIRCLE PERFORMANCE
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Part 3: Basic Statistics", "Circle Rankings & The Urban-Rural Disparity",
               "Comparison of 20 electricity circles reveals profound geographical concentration of energy")

    # Table of Top 5 vs Bottom 5 Circles
    headers = ["Performance Tier", "Circle Name", "Total Units (GWh)", "Share (%)", "Billed Connections", "Avg kWh / Service"]
    rows = [
        ["Top 1 (Highest)", "MEDCHAL", "141.74 GWh", "11.38%", "806,897", "175.7 kWh"],
        ["Top 2", "HABSIGUDA", "138.12 GWh", "11.09%", "751,204", "183.9 kWh"],
        ["Top 3", "CYBERCITY", "125.43 GWh", "10.07%", "572,423", "219.1 kWh"],
        ["Top 4", "HYD CENTRAL", "112.51 GWh", "9.03%", "518,412", "217.0 kWh"],
        ["Top 5", "SAROORNAGAR", "112.30 GWh", "9.02%", "682,149", "164.6 kWh"],
        ["Bottom 5", "VIKARABAD", "17.92 GWh", "1.44%", "215,901", "83.0 kWh"],
        ["Bottom 4", "MEDAK", "15.91 GWh", "1.28%", "189,402", "84.0 kWh"],
        ["Bottom 3", "NAGARKURNOOL", "13.12 GWh", "1.05%", "174,930", "75.0 kWh"],
        ["Bottom 2", "WANAPARTHY", "10.11 GWh", "0.81%", "134,800", "75.0 kWh"],
        ["Bottom 1 (Lowest)", "GADWAL", "10.02 GWh", "0.80%", "123,700", "81.0 kWh"]
    ]
    col_w = [Inches(1.8), Inches(2.0), Inches(2.0), Inches(1.3), Inches(2.2), Inches(2.0)]
    create_table(sl, Inches(0.5), Inches(1.6), Inches(11.3), Inches(3.4), headers, rows, col_w)

    # Analytical Insights below
    bullet_card(sl, Inches(0.5), Inches(5.15), Inches(5.8), Inches(1.7),
                "Hyderabad Urban Dominance", [
                    ("Top 5 Circles Consume 50.6%", "The top 5 urban circles account for 630.1 GWh out of 1,245.5 GWh."),
                    ("Top 10 Circles Supply ~90%", "Virtually all major demand is concentrated in the Greater Hyderabad ring.")
                ], NAVY)

    bullet_card(sl, Inches(6.6), Inches(5.15), Inches(6.2), Inches(1.7),
                "The 3× Lifestyle Consumption Gap", [
                    ("Affluence & Appliance Usage", "Banjara Hills (240 kWh) & Cybercity (219 kWh) lead consumption intensity."),
                    ("District Baseline", "Rural circles average 75-83 kWh/connection — nearly one-third of tech corridor usage.")
                ], TEAL_DARK)

    footer(sl, 8)

    # ==============================================================================
    # SLIDE 9: PART 3 — BASIC STATISTICS: THE 15% / 50.5% PARETO RULE
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Part 3: Basic Statistics", "The Pareto Principle: Area-Size Tier Segmentation",
               "15.4% of neighborhood areas account for over half (50.5%) of total electricity consumed")

    tiers = [
        ("Very Large Areas", "> 1,000 Services", "2,452 Areas (15.4%)", "629.0 GWh (50.5%)",
         "High-density urban apartment corridors, IT corridors & townships. Highest priority for transformer health.", RED),
        ("Large Areas", "501 to 1,000 Services", "3,810 Areas (23.9%)", "301.4 GWh (24.2%)",
         "Established municipal colonies and commercial-residential mixed zones. High baseline consumption.", GOLD_DARK),
        ("Medium Areas", "201 to 500 Services", "5,120 Areas (32.1%)", "225.4 GWh (18.1%)",
         "Semi-urban and peri-urban wards. Moderate per-feeder loading with stable seasonal demand.", NAVY_MED),
        ("Small Areas", "1 to 200 Services", "4,565 Areas (28.6%)", "89.7 GWh (7.2%)",
         "Rural hamlets, agricultural borders, and tribal thandas. Extensive line length but minimal energy volume.", GREEN_DARK)
    ]

    for idx, (t_name, t_range, t_count, t_units, t_desc, accent_c) in enumerate(tiers):
        x_pos = Inches(0.5) + idx * Inches(3.15)
        card(sl, x_pos, Inches(1.6), Inches(3.0), Inches(5.2))
        rect(sl, x_pos, Inches(1.6), Inches(3.0), Inches(0.48), accent_c, radius=0.04)
        txt(sl, t_name, x_pos + Inches(0.1), Inches(1.7), Inches(2.8), Inches(0.3),
            size=Pt(12), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")
            
        txt(sl, t_range, x_pos + Inches(0.1), Inches(2.18), Inches(2.8), Inches(0.25),
            size=Pt(10), bold=True, color=accent_c, align=PP_ALIGN.CENTER, name="Calibri")
            
        rect(sl, x_pos + Inches(0.2), Inches(2.5), Inches(2.6), Inches(0.7), ROW_ALT, radius=0.04)
        txt(sl, "SECTOR COUNT", x_pos + Inches(0.25), Inches(2.55), Inches(2.5), Inches(0.2),
            size=Pt(8), bold=True, color=SLATE, align=PP_ALIGN.CENTER, name="Calibri")
        txt(sl, t_count, x_pos + Inches(0.25), Inches(2.75), Inches(2.5), Inches(0.38),
            size=Pt(11), bold=True, color=NAVY, align=PP_ALIGN.CENTER, name="Calibri")

        rect(sl, x_pos + Inches(0.2), Inches(3.3), Inches(2.6), Inches(0.7), ROW_ALT, radius=0.04)
        txt(sl, "ENERGY SHARE", x_pos + Inches(0.25), Inches(3.35), Inches(2.5), Inches(0.2),
            size=Pt(8), bold=True, color=SLATE, align=PP_ALIGN.CENTER, name="Calibri")
        txt(sl, t_units, x_pos + Inches(0.25), Inches(3.55), Inches(2.5), Inches(0.38),
            size=Pt(12), bold=True, color=accent_c, align=PP_ALIGN.CENTER, name="Calibri")

        txt(sl, t_desc, x_pos + Inches(0.2), Inches(4.15), Inches(2.6), Inches(2.4),
            size=Pt(9.5), color=DARK_TEXT, name="Calibri")

    footer(sl, 9)

    # ==============================================================================
    # SLIDE 10: PART 4 — DATA WRANGLING: 6-STEP PREPARATION PIPELINE
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Part 4: Data Wrangling", "The 6-Step Automated Data Wrangling & ETL Pipeline",
               "Systematic data transformations applied in Power Query and Python Pandas")

    etl_steps = [
        ("Step 1: Ingestion", "Connect to multi-source CSV & Excel workbooks via Power Query Folder connector / Pandas read_csv.", NAVY),
        ("Step 2: Null Treatment", "Impute 47 missing AREA names as 'UNKNOWN AREA'; replace missing BILLDSERVICES with 0 / registered services.", TEAL_DARK),
        ("Step 3: Anomaly Clamping", "Identify 15 negative consumption records and clamp units to min 0.0; clamp minimum load to 0.01 kW.", RED),
        ("Step 4: Text Normalization", "Trim leading/trailing whitespace, resolve capitalization mismatches, and build Section-Area keys.", GOLD_DARK),
        ("Step 5: Type Casting", "Cast service counts to uint32, metered units to float64, and administrative levels to categorical strings.", NAVY_MED),
        ("Step 6: Validation", "Enforce logical integrity constraints: (BILLDSERVICES <= TOTSERVICES), recalculate summaries.", GREEN_DARK)
    ]

    for idx, (s_title, s_desc, s_color) in enumerate(etl_steps):
        y_pos = Inches(1.6) + idx * Inches(0.86)
        card(sl, Inches(0.5), y_pos, Inches(12.33), Inches(0.78))
        rect(sl, Inches(0.5), y_pos, Inches(2.5), Inches(0.78), s_color, radius=0.04)
        txt(sl, s_title, Inches(0.6), y_pos + Inches(0.24), Inches(2.3), Inches(0.3),
            size=Pt(12), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")
        txt(sl, s_desc, Inches(3.2), y_pos + Inches(0.18), Inches(9.4), Inches(0.42),
            size=Pt(10.5), color=DARK_TEXT, name="Calibri")

    footer(sl, 10)

    # ==============================================================================
    # SLIDE 11: PART 4 — DATA WRANGLING: FEATURE ENGINEERING & METRICS
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Part 4: Data Wrangling", "Feature Engineering: Synthesizing Domain Indicators",
               "Mathematical formulations enabling multi-dimensional utility efficiency assessments")

    features = [
        ("Billing Efficiency Rate (%)",
         "billing_rate_pct = (BILLDSERVICES / TOTSERVICES) * 100",
         "Quantifies revenue capture performance. Highlights areas with severe metering delays, non-functional meters, or ghost connections. Benchmark is 95%.",
         GREEN),
        ("Consumption Intensity per Connection",
         "units_per_billed_service = UNITS / BILLDSERVICES",
         "Normalizes consumption against connection volume. Eliminates size bias, revealing true household lifestyle intensity (e.g., Banjara Hills vs Rural).",
         TEAL),
        ("Load Utilization Ratio (kWh / kW)",
         "units_per_kw_load = UNITS / LOAD",
         "Measures how heavily connected power capacity is utilized. Sectors > 250 kWh/kW indicate severe transformer strain or commercial misuse of domestic tariffs.",
         GOLD_DARK),
        ("Unbilled Service Gap",
         "unbilled_services = TOTSERVICES - BILLDSERVICES",
         "Direct head-count of meters that failed to register billing in July 2026. Feeds directly into field inspection priority queues (total 648,960 services).",
         RED)
    ]

    for idx, (f_name, f_formula, f_rationale, f_color) in enumerate(features):
        row = idx // 2
        col = idx % 2
        x_f = Inches(0.5) + col * Inches(6.3)
        y_f = Inches(1.6) + row * Inches(2.65)
        
        card(sl, x_f, y_f, Inches(6.0), Inches(2.45))
        rect(sl, x_f, y_f, Inches(6.0), Inches(0.42), f_color, radius=0.04)
        txt(sl, f_name, x_f + Inches(0.2), y_f + Inches(0.08), Inches(5.6), Inches(0.28),
            size=Pt(12), bold=True, color=WHITE, name="Calibri")
            
        # Formula container
        rect(sl, x_f + Inches(0.2), y_f + Inches(0.52), Inches(5.6), Inches(0.45), ROW_ALT, radius=0.04)
        txt(sl, f_formula, x_f + Inches(0.3), y_f + Inches(0.6), Inches(5.4), Inches(0.3),
            size=Pt(9.5), bold=True, color=NAVY, name="Courier New")
            
        txt(sl, f_rationale, x_f + Inches(0.2), y_f + Inches(1.1), Inches(5.6), Inches(1.2),
            size=Pt(10), color=DARK_TEXT, name="Calibri")

    footer(sl, 11)

    # ==============================================================================
    # SLIDE 12: PART 5 — VARIABLE CONNECTIONS: CORRELATION MATRIX
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Part 5: Variable Connections", "Correlation Analysis & Pairwise Mathematical Associations",
               "Evaluating Pearson correlation coefficients across operational electrical parameters")

    # Correlation Table
    headers = ["Variable Pair", "Correlation (r)", "Relationship Strength", "Analytical & Operational Interpretation"]
    rows = [
        ["TOTSERVICES ↔ BILLDSERVICES", "+0.995", "Near Perfect Positive", "Sanctioned connection count strongly dictates billed connections across all regions."],
        ["LOAD (kW) ↔ UNITS (kWh)", "+0.946", "Very Strong Positive", "Connected load capacity is the primary physical driver of metered energy consumption."],
        ["BILLDSERVICES ↔ UNITS", "+0.884", "Strong Positive", "Larger customer bases naturally generate higher aggregated energy demand."],
        ["TOTSERVICES ↔ UNITS", "+0.863", "Strong Positive", "Total registered connections reliably predict area-level gross consumption."],
        ["BILLDSERVICES ↔ LOAD", "+0.786", "Moderate-High Positive", "Sanctioned load accumulates linearly as household connection counts expand."],
        ["TOTSERVICES ↔ LOAD", "+0.763", "Moderate-High Positive", "Aggregate connected infrastructure tracks total neighborhood scale."]
    ]
    col_w = [Inches(2.8), Inches(1.5), Inches(2.2), Inches(5.8)]
    create_table(sl, Inches(0.5), Inches(1.6), Inches(12.3), Inches(3.2), headers, rows, col_w)

    # Correlation vs Causation Card
    bullet_card(sl, Inches(0.5), Inches(5.0), Inches(12.33), Inches(1.8),
                "Key Analytical Finding: Association vs Physical Causation", [
                    ("Physical Power Reality", "High connected load (kW) does not cause electricity usage by itself — it creates the capacity for appliances (ACs, pumps) to consume power."),
                    ("Absence of Tariff/Weather Covariates", "Because temperature, humidity, and building square footage are absent, LOAD serves as the best available mathematical proxy for consumption capacity."),
                    ("Anomaly Detection Utility", "Any area deviating significantly from the expected r = 0.95 regression line flags possible unmetered power theft, faulty meters, or misclassified commercial activity.")
                ], NAVY)

    footer(sl, 12)

    # ==============================================================================
    # SLIDE 13: PART 5 — VARIABLE CONNECTIONS: LOAD VS CONSUMPTION REGRESSION
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Part 5: Variable Connections", "Connected Load vs Consumed Units: Regression Dynamics",
               "Empirical validation of the 0.946 correlation through Seaborn scatter & linear regression")

    # Left: Image Card with load_vs_units.png
    add_image_card(sl, Inches(0.5), Inches(1.6), Inches(6.8), Inches(5.2),
                   "load_vs_units.png", "Figure: Connected Load (kW) vs Consumed Units (MWh) with Seaborn Regplot")

    # Right: Interpretive bullet cards
    bullet_card(sl, Inches(7.5), Inches(1.6), Inches(5.3), Inches(2.5),
                "Regression Slope & Load Factor", [
                    ("Linear Regression Trend", "The red regression line displays a consistent upward slope confirming linear scalability."),
                    ("Tight Baseload Band", "The dense cluster between 0-2,000 kW shows tightly bounded baseload domestic behavior."),
                    ("Load Factor Consistency", "Average load factor across domestic sectors remains within 18% to 26% during the monsoon month.")
                ], NAVY)

    bullet_card(sl, Inches(7.5), Inches(4.3), Inches(5.3), Inches(2.5),
                "Outlier & Anomaly Clustering", [
                    ("High Vertical Scatter", "Sectors plotted well above the regression line indicate high consumption per kW (heavy AC/luxury loads)."),
                    ("Sub-Regression Trajectory", "Points well below the line represent areas with high sanctioned load but low actual consumption (vacant apartments/plots)."),
                    ("Operational Value", "Enables TG-SPDCL to set dynamic power supply schedules and detect meter bypass.")
                ], TEAL_DARK)

    footer(sl, 13)

    # ==============================================================================
    # SLIDE 14: PART 6 — EXCEL VISUALIZATIONS: PIVOT TABLES & SLICERS
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Part 6: Excel Visualizations", "Microsoft Excel Analytics: Pivot Tables & Multi-Tier Slicers",
               "Rapid ad-hoc aggregation workflows across Circles, Divisions, and Section hierarchies")

    # 3 Pivot Table Architecture Cards
    bullet_card(sl, Inches(0.5), Inches(1.6), Inches(3.9), Inches(5.2),
                "1. Circle-Level Pivot Summary", [
                    ("Rows", "CIRCLE (20 distinct values)."),
                    ("Values", "Sum of UNITS, Sum of LOAD, Sum of TOTSERVICES, Sum of BILLDSERVICES."),
                    ("Calculated Field", "'Billing Rate %' = BILLDSERVICES / TOTSERVICES."),
                    ("Sorting", "Descending by Sum of UNITS."),
                    ("Executive Insight", "Instantly separates high-consumption metro circles (Medchal, Habsiguda) from rural districts (Gadwal)."),
                    ("Formatting", "Excel Data Bars applied to Units column for rapid visual benchmarking.")
                ], RGBColor(0x10, 0x7c, 0x41))  # Excel Green

    bullet_card(sl, Inches(4.7), Inches(1.6), Inches(3.9), Inches(5.2),
                "2. Division & Section Drill-Down", [
                    ("Hierarchical Nesting", "CIRCLE > DIVISION > SECTION."),
                    ("Metrics", "Sum of UNITS and Average Units per Billed Service."),
                    ("Interactive Slicers", "Slicers added for CIRCLE to allow field engineers to isolate one zone with one click."),
                    ("Conditional Formatting", "Highlight cells with Billing Efficiency < 85% in Light Red Fill."),
                    ("Top Division Identified", "Kukatpally (62.0M units) and Sainikpuri (55.2M units) drive regional totals.")
                ], NAVY_MED)

    bullet_card(sl, Inches(8.9), Inches(1.6), Inches(3.9), Inches(5.2),
                "3. Advanced Excel Filtering", [
                    ("Negative Units Filter", "Filter UNITS < 0 isolates the 15 billing reversal rows in seconds."),
                    ("Null Billed Services Filter", "Filtering blank BILLDSERVICES exposes the 247 zero-consumption sectors."),
                    ("Top 10% Areas", "Filtering the top decile reveals that 1,595 areas account for 43% of all domestic units."),
                    ("Domestic HT Outliers", "Filtering per-connection average > 3,000 units surfaces bulk residential high-tension connections.")
                ], GOLD_DARK)

    footer(sl, 14)

    # ==============================================================================
    # SLIDE 15: PART 6 — EXCEL VISUALIZATIONS: CHARTS & TOOL COMPARISON
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Part 6: Excel Visualizations", "Excel Visual Dashboards vs Tool Boundaries",
               "Clustered Column and Donut charts in Excel compared against Python data capabilities")

    # Left: Image of Excel-style charts
    add_image_card(sl, Inches(0.5), Inches(1.6), Inches(6.8), Inches(5.2),
                   "excel_visualization.png", "Figure: Microsoft Excel Visualizations (Clustered Column & 2D Donut Chart)")

    # Right: Comparative Matrix
    bullet_card(sl, Inches(7.5), Inches(1.6), Inches(5.3), Inches(2.5),
                "Microsoft Excel Strengths in DAV", [
                    ("Zero-Code Accessibility", "Immediate interface for field staff, managers, and non-technical stakeholders."),
                    ("Interactive Slicers & Timelines", "Fluid point-and-click dashboard filtering without code recompilation."),
                    ("Native Business Reporting", "Standard format for utility tariff filings, regulatory audits, and weekly reviews.")
                ], RGBColor(0x10, 0x7c, 0x41))

    bullet_card(sl, Inches(7.5), Inches(4.3), Inches(5.3), Inches(2.5),
                "Excel Limitations & Why Python is Essential", [
                    ("Scalability Constraints", "Excel performance degrades significantly on multi-million row smart meter streams."),
                    ("Lack of Advanced Statistical Plots", "Cannot natively generate KDE density plots, regplots, or correlation heatmaps."),
                    ("Manual Reproducibility Risk", "Vulnerable to accidental cell formula overwrites and manual copy-paste errors.")
                ], NAVY)

    footer(sl, 15)

    # ==============================================================================
    # SLIDE 16: PART 7 — PYTHON VISUALIZATIONS: CIRCLE DUAL-AXIS ANALYSIS
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Part 7: Python Visualizations", "Dual-Axis Infrastructure Analysis: Consumption vs Load",
               "Advanced Python Matplotlib & Seaborn visualization mapping 20 circles on twin y-axes")

    # Left: Image of circle_consumption.png
    add_image_card(sl, Inches(0.5), Inches(1.6), Inches(6.8), Inches(5.2),
                   "circle_consumption.png", "Figure: Energy Consumed (GWh) & Sanctioned Load (MW) by Circle (Matplotlib & Seaborn)")

    # Right: Analytical Takeaways
    bullet_card(sl, Inches(7.5), Inches(1.6), Inches(5.3), Inches(2.5),
                "Dual-Axis Engineering Alignment", [
                    ("Proportional Symmetry", "Consumption (cyan bars, GWh) and Connected Load (purple bars, MW) mirror each other perfectly."),
                    ("Peak Demand Readiness", "Validates that grid substations in Medchal, Habsiguda, and Cybercity have sanctioned capacity to meet high consumption."),
                    ("Load Allocation Equity", "District circles maintain proportional transformer capacity, avoiding wasteful idle infrastructure.")
                ], TEAL_DARK)

    bullet_card(sl, Inches(7.5), Inches(4.3), Inches(5.3), Inches(2.5),
                "Python Implementation Highlights", [
                    ("Subplot Twinx Architecture", "Created using `ax1.twinx()` to overlay two independent physical units cleanly."),
                    ("Theme & Styling", "Styled with custom dark-background aesthetic, Segoe UI typography, and subtle alpha grids."),
                    ("Automated Export", "High-DPI 300 vector rendering exported directly to disk for web and reporting pipelines.")
                ], NAVY)

    footer(sl, 16)

    # ==============================================================================
    # SLIDE 17: PART 7 — PYTHON VISUALIZATIONS: TOP 20 AREAS & BILLING EFFICIENCY
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Part 7: Python Visualizations", "Consumption Epicenters & Section Billing Benchmark",
               "Seaborn barplots isolating the top 20 demand hotspots and circle billing efficiency vs 95% target")

    # Dual Image layout
    add_image_card(sl, Inches(0.5), Inches(1.6), Inches(6.0), Inches(4.2),
                   "top20_areas.png", "Top 20 High-Demand Areas in TG-SPDCL (Seaborn Barplot)")

    add_image_card(sl, Inches(6.8), Inches(1.6), Inches(6.0), Inches(4.2),
                   "billing_efficiency.png", "Circle Billing Efficiency (%) vs 95% State Benchmark Target")

    # Bottom Takeaway Card
    bullet_card(sl, Inches(0.5), Inches(6.0), Inches(12.33), Inches(0.9),
                "Key Findings", [
                    ("Top Demand Epicenters", "Hydernagar (2.32M kWh), Moosapet (1.81M), and Vinayaknagar (1.47M) lead consumption; require immediate feeder audits."),
                    ("Billing Compliance Deficit", "5 circles lag below the 95% benchmark; Saroornagar and rural zones show significant unbilled gaps requiring mobile meter squads.")
                ], NAVY)

    footer(sl, 17)

    # ==============================================================================
    # SLIDE 18: PART 7 — PYTHON VISUALIZATIONS: KDE DISTRIBUTIONS & HEATMAP
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Part 7: Python Visualizations", "Statistical Distributions (KDE) & Seaborn Heatmap",
               "Evaluating density spread and multi-collinear relationships across all numeric indicators")

    add_image_card(sl, Inches(0.5), Inches(1.6), Inches(6.3), Inches(4.2),
                   "distribution_metrics.png", "Figure: KDE Density Curves for Household Units & Section Billing Efficiency")

    add_image_card(sl, Inches(7.1), Inches(1.6), Inches(5.7), Inches(4.2),
                   "seaborn_correlation_heatmap.png", "Figure: Seaborn Indicator Correlation Heatmap (Spectral Palette)")

    # Bottom Takeaway
    bullet_card(sl, Inches(0.5), Inches(6.0), Inches(12.33), Inches(0.9),
                "Distribution & Heatmap Synthesis", [
                    ("Household Consumption Distribution", "Heavily skewed with median 145 kWh/service; sharp exponential decay above 400 kWh."),
                    ("Billing Rate Concentration", "Sharp unimodal peak centered at 94.2%, with a concerning left-hand tail of unbilled sections (<85%).")
                ], NAVY)

    footer(sl, 18)

    # ==============================================================================
    # SLIDE 19: PART 8 — FINAL DECISION: STRATEGIC 4-PILLAR ACTION PLAN
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Part 8: Final Decision", "Strategic Decision Framework: 4-Pillar Utility Action Plan",
               "Data-backed operational interventions to maximize revenue, grid stability, and consumer reliability")

    pillars = [
        ("Pillar 1: Infrastructure Augmentation",
         "Prioritize Top 15% Areas",
         [
             "Focus distribution transformer (DTR) capital upgrades on the 2,452 'Very Large' areas driving 50.5% of total demand.",
             "Deploy thermal sensors and DTR capacity doubling in Hydernagar, Moosapet, and Kukatpally.",
             "Enforce mandatory load sanctions for luxury villas and high-rise apartment complexes."
         ], TEAL_DARK),
        ("Pillar 2: Revenue Protection & Audits",
         "Resolve 648k Unbilled Services",
         [
             "Mandate physical field verification for the 247 sections with missing billed services.",
             "Investigate the 15 negative-unit areas to distinguish legitimate meter corrections from billing manipulation.",
             "Deploy mobile meter-reading validation squads in circles falling below the 95% billing benchmark."
         ], RED),
        ("Pillar 3: Smart Grid Modernization",
         "Automated Metering (AMI / IoT)",
         [
             "Roll out bidirectional IoT smart meters across high-intensity sectors (> 250 kWh/kW).",
             "Eliminate manual paper billing errors and eliminate the 1,999 area-name ambiguity issues through GIS tagging.",
             "Provide consumers with hourly mobile-app telemetry to incentivize peak-hour energy conservation."
         ], NAVY_MED),
        ("Pillar 4: Tariff & Demand Management",
         "Time-of-Day (ToD) Pricing",
         [
             "Introduce dynamic Time-of-Day tariffs for domestic connections exceeding 10 kW connected load.",
             "Incentivize off-peak EV charging and water pumping to flatten the late-evening domestic peak.",
             "Establish automated loss-detection algorithms flagging feeders with high load but abnormally low metered units."
         ], GOLD_DARK)
    ]

    for idx, (p_title, p_sub, p_items, p_accent) in enumerate(pillars):
        row = idx // 2
        col = idx % 2
        x_p = Inches(0.5) + col * Inches(6.3)
        y_p = Inches(1.6) + row * Inches(2.65)
        
        card(sl, x_p, y_p, Inches(6.0), Inches(2.45))
        rect(sl, x_p, y_p, Inches(6.0), Inches(0.42), p_accent, radius=0.04)
        txt(sl, p_title, x_p + Inches(0.2), y_p + Inches(0.08), Inches(5.6), Inches(0.28),
            size=Pt(12), bold=True, color=WHITE, name="Calibri")
            
        txt(sl, f"Focus: {p_sub}", x_p + Inches(0.2), y_p + Inches(0.48), Inches(5.6), Inches(0.22),
            size=Pt(9.5), bold=True, color=p_accent, name="Calibri")
            
        box = sl.shapes.add_textbox(x_p + Inches(0.2), y_p + Inches(0.72), Inches(5.6), Inches(1.6))
        tf = box.text_frame
        tf.word_wrap = True
        for i_num, it in enumerate(p_items):
            p = tf.paragraphs[0] if i_num == 0 else tf.add_paragraph()
            p.space_before = Pt(3)
            p.space_after = Pt(2)
            run = p.add_run()
            run.text = f"•  {it}"
            run.font.size = Pt(9.5)
            run.font.color.rgb = DARK_TEXT
            run.font.name = "Calibri"

    footer(sl, 19)

    # ==============================================================================
    # SLIDE 20: PART 8 — FINAL DECISION: ROADMAP & BUSINESS ROI
    # ==============================================================================
    sl = prs.slides.add_slide(BLANK)
    set_bg(sl)
    header_bar(sl, "Part 8: Final Decision", "Implementation Roadmap & Measurable Business Impact",
               "Phased execution timeline and financial ROI projections for TG-SPDCL executive leadership")

    # 3-Phase Timeline Flow
    phases = [
        ("Phase 1: Immediate Sprint", "Month 1 (Days 1 - 30)", [
            "Audit the 247 zero-billed sections and 15 negative-unit records.",
            "Establish GIS standardized primary keys for all 15,947 areas.",
            "Deploy mobile billing taskforces to circles < 90% billing rate."
        ], RED),
        ("Phase 2: Targeted Deployment", "Months 2 - 6 (Days 31 - 180)", [
            "Upgrade transformers across the 2,452 'Very Large' demand areas.",
            "Install 250,000 smart meters across Top 20 urban epicenters.",
            "Integrate real-time operational dashboard for daily feeder tracking."
        ], GOLD_DARK),
        ("Phase 3: Utility Transformation", "Months 6 - 12 (Year 1)", [
            "Achieve full 97%+ system-wide billing efficiency across all circles.",
            "Launch Time-of-Day (ToD) tariff structures for high-demand users.",
            "Implement predictive ML models for proactive failure prevention."
        ], GREEN_DARK)
    ]

    for idx, (p_name, p_time, p_tasks, p_col) in enumerate(phases):
        x_ph = Inches(0.5) + idx * Inches(4.2)
        card(sl, x_ph, Inches(1.6), Inches(3.95), Inches(3.2))
        rect(sl, x_ph, Inches(1.6), Inches(3.95), Inches(0.42), p_col, radius=0.04)
        txt(sl, p_name, x_ph + Inches(0.15), Inches(1.68), Inches(3.65), Inches(0.28),
            size=Pt(12), bold=True, color=WHITE, align=PP_ALIGN.CENTER, name="Calibri")
            
        rect(sl, x_ph + Inches(0.2), Inches(2.1), Inches(3.55), Inches(0.3), ROW_ALT, radius=0.04)
        txt(sl, p_time, x_ph + Inches(0.25), Inches(2.15), Inches(3.45), Inches(0.2),
            size=Pt(9), bold=True, color=p_col, align=PP_ALIGN.CENTER, name="Calibri")
            
        box = sl.shapes.add_textbox(x_ph + Inches(0.2), Inches(2.5), Inches(3.55), Inches(2.1))
        tf = box.text_frame
        tf.word_wrap = True
        for t_idx, task in enumerate(p_tasks):
            p = tf.paragraphs[0] if t_idx == 0 else tf.add_paragraph()
            p.space_before = Pt(4)
            p.space_after = Pt(2)
            run = p.add_run()
            run.text = f"✔  {task}"
            run.font.size = Pt(9.5)
            run.font.color.rgb = DARK_TEXT
            run.font.name = "Calibri"

    # Measurable ROI Box below
    card(sl, Inches(0.5), Inches(5.0), Inches(12.33), Inches(1.8), bg=NAVY, border=GOLD)
    rect(sl, Inches(0.5), Inches(5.0), Inches(12.33), Inches(0.06), GOLD)
    txt(sl, "PROJECTED BUSINESS IMPACT & RETURN ON INVESTMENT (ROI)", Inches(0.8), Inches(5.15), Inches(11.5), Inches(0.25),
        size=Pt(11), bold=True, color=GOLD, name="Calibri")

    roi_stats = [
        ("+4.5% Billing Efficiency", "Unlocks ~₹85-110 Cr annual revenue by converting unbilled connections to active billing.", TEAL),
        ("-35% Transformer Failures", "Targeted DTR load doubling prevents summer overloads in the top 15% consumption corridors.", GREEN),
        ("100% Traceability", "Complete elimination of repeated area name ambiguity via GIS composite key tagging.", GOLD),
        ("Real-Time Grid Telemetry", "Interactive dashboard enables daily monitoring rather than 30-day delayed manual reporting.", WHITE)
    ]
    for idx, (roi_title, roi_sub, roi_color) in enumerate(roi_stats):
        x_r = Inches(0.8) + idx * Inches(2.95)
        txt(sl, roi_title, x_r, Inches(5.45), Inches(2.8), Inches(0.3),
            size=Pt(11), bold=True, color=roi_color, name="Calibri")
        txt(sl, roi_sub, x_r, Inches(5.8), Inches(2.75), Inches(0.85),
            size=Pt(8.5), color=RGBColor(0xcf, 0xd8, 0xdc), name="Calibri")

    footer(sl, 20)

    # ─── SAVE PRESENTATION ────────────────────────────────────────────────────────
    out_file = "TG_SPDCL_Electricity_Analysis_v3.pptx"
    prs.save(out_file)
    print(f"Successfully generated Version 3 presentation: '{out_file}' (20 Slides)")

if __name__ == "__main__":
    build_presentation()
