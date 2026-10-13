# -*- coding: utf-8 -*-
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
"""
PPT Generator for TG-SPDCL Electricity Consumption & Monitoring Project
Authors: D.Chander, K.Arjun, P.Sachin
Roll No: 2503A52214, 2503A52235, 2503A52240
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import os

# ─── Color Palette ─────────────────────────────────────────────────────────────
BG_DARK       = RGBColor(0x07, 0x0a, 0x13)   # near-black navy
PANEL         = RGBColor(0x0f, 0x17, 0x2a)   # dark blue panel
CYAN          = RGBColor(0x00, 0xf2, 0xfe)   # electric cyan
PURPLE        = RGBColor(0xa8, 0x55, 0xf7)   # vivid purple
EMERALD       = RGBColor(0x10, 0xb9, 0x81)   # emerald green
AMBER         = RGBColor(0xf5, 0x9e, 0x0b)   # amber
ROSE          = RGBColor(0xf4, 0x3f, 0x5e)   # rose red
WHITE         = RGBColor(0xf8, 0xfa, 0xfc)   # near white
MUTED         = RGBColor(0x94, 0xa3, 0xb8)   # slate muted
SLIDE_W       = Inches(13.33)
SLIDE_H       = Inches(7.5)

# ─── Helpers ───────────────────────────────────────────────────────────────────
def add_bg(slide, color=BG_DARK):
    """Fill slide background."""
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_rect(slide, left, top, width, height, fill_color, line_color=None, line_width=Pt(0)):
    shape = slide.shapes.add_shape(1, left, top, width, height)  # MSO_SHAPE_TYPE.RECTANGLE
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape

def add_text(slide, text, left, top, width, height, font_size=Pt(14),
             bold=False, color=WHITE, align=PP_ALIGN.LEFT, italic=False, font_name="Segoe UI"):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = font_size
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = font_name
    return txBox

def add_accent_bar(slide, color=CYAN, top=Inches(0.08)):
    """Thin colored top bar."""
    add_rect(slide, 0, top, SLIDE_W, Inches(0.06), color)

def add_slide_number(slide, num):
    add_text(slide, f"{num:02d}", Inches(12.5), Inches(7.05), Inches(0.7), Inches(0.4),
             font_size=Pt(11), color=MUTED, align=PP_ALIGN.RIGHT)

def bullet_slide(prs, title, bullets, icon="⚡", accent=CYAN, slide_num=1, subtitle=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    add_bg(slide)
    add_accent_bar(slide, accent)

    # Left panel (dark stripe)
    add_rect(slide, 0, 0, Inches(0.18), SLIDE_H, accent)

    # Title
    add_text(slide, f"{icon}  {title}", Inches(0.4), Inches(0.18),
             Inches(12.5), Inches(0.7), font_size=Pt(28), bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    if subtitle:
        add_text(slide, subtitle, Inches(0.4), Inches(0.9), Inches(12.5), Inches(0.4),
                 font_size=Pt(15), color=MUTED, align=PP_ALIGN.LEFT, italic=True)

    # Divider line
    add_rect(slide, Inches(0.4), Inches(1.28), Inches(12.5), Inches(0.02), accent)

    # Bullets
    y = Inches(1.45)
    for bullet in bullets:
        if isinstance(bullet, tuple):
            label, val = bullet
            # Label box
            add_rect(slide, Inches(0.4), y, Inches(3.5), Inches(0.42), PANEL, accent, Pt(1))
            add_text(slide, label, Inches(0.5), y + Pt(4), Inches(3.3), Inches(0.42),
                     font_size=Pt(13), bold=True, color=accent)
            add_text(slide, val, Inches(4.0), y + Pt(4), Inches(9.0), Inches(0.42),
                     font_size=Pt(13), color=WHITE)
            y += Inches(0.54)
        else:
            add_text(slide, f"  ▸  {bullet}", Inches(0.4), y, Inches(12.5), Inches(0.5),
                     font_size=Pt(14), color=WHITE)
            y += Inches(0.5)

    add_slide_number(slide, slide_num)
    return slide

def kpi_slide(prs, title, kpis, accent=CYAN, slide_num=1):
    """KPI card grid slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    add_accent_bar(slide, accent)
    add_rect(slide, 0, 0, Inches(0.18), SLIDE_H, accent)

    add_text(slide, f"📊  {title}", Inches(0.4), Inches(0.18),
             Inches(12.5), Inches(0.7), font_size=Pt(28), bold=True, color=WHITE)
    add_rect(slide, Inches(0.4), Inches(0.9), Inches(12.5), Inches(0.02), accent)

    # 4-column KPI cards
    cols = 4
    card_w = Inches(3.0)
    card_h = Inches(1.7)
    col_gap = Inches(0.22)
    start_x = Inches(0.35)
    start_y = Inches(1.1)

    kpi_colors = [CYAN, EMERALD, PURPLE, AMBER, ROSE, CYAN, EMERALD, PURPLE]

    for i, (label, val, sub) in enumerate(kpis):
        col = i % cols
        row = i // cols
        x = start_x + col * (card_w + col_gap)
        y = start_y + row * (card_h + Inches(0.18))
        c = kpi_colors[i % len(kpi_colors)]

        add_rect(slide, x, y, card_w, card_h, PANEL, c, Pt(1.5))
        # Color top strip
        add_rect(slide, x, y, card_w, Inches(0.06), c)
        add_text(slide, label, x + Inches(0.12), y + Inches(0.12), card_w - Inches(0.24), Inches(0.35),
                 font_size=Pt(11), color=MUTED, bold=False)
        add_text(slide, val, x + Inches(0.12), y + Inches(0.45), card_w - Inches(0.24), Inches(0.75),
                 font_size=Pt(22), bold=True, color=c)
        add_text(slide, sub, x + Inches(0.12), y + Inches(1.25), card_w - Inches(0.24), Inches(0.38),
                 font_size=Pt(10), color=MUTED)

    add_slide_number(slide, slide_num)
    return slide

def chart_slide(prs, title, img_path, caption, accent=CYAN, slide_num=1):
    """Slide with chart image."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    add_accent_bar(slide, accent)
    add_rect(slide, 0, 0, Inches(0.18), SLIDE_H, accent)

    add_text(slide, f"📈  {title}", Inches(0.4), Inches(0.15),
             Inches(12.5), Inches(0.65), font_size=Pt(24), bold=True, color=WHITE)
    add_rect(slide, Inches(0.4), Inches(0.82), Inches(12.5), Inches(0.02), accent)

    if img_path and os.path.exists(img_path):
        slide.shapes.add_picture(img_path, Inches(0.4), Inches(0.92), Inches(12.4), Inches(5.85))

    add_text(slide, caption, Inches(0.5), Inches(6.88), Inches(12.3), Inches(0.5),
             font_size=Pt(11), color=MUTED, italic=True, align=PP_ALIGN.CENTER)
    add_slide_number(slide, slide_num)
    return slide

def two_col_slide(prs, title, left_title, left_items, right_title, right_items,
                  accent=CYAN, slide_num=1):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    add_accent_bar(slide, accent)
    add_rect(slide, 0, 0, Inches(0.18), SLIDE_H, accent)

    add_text(slide, title, Inches(0.4), Inches(0.18),
             Inches(12.5), Inches(0.65), font_size=Pt(26), bold=True, color=WHITE)
    add_rect(slide, Inches(0.4), Inches(0.85), Inches(12.5), Inches(0.02), accent)

    # Left col
    add_rect(slide, Inches(0.35), Inches(1.0), Inches(6.0), Inches(5.9), PANEL, accent, Pt(1))
    add_text(slide, left_title, Inches(0.55), Inches(1.1), Inches(5.8), Inches(0.45),
             font_size=Pt(15), bold=True, color=accent)
    y = Inches(1.62)
    for item in left_items:
        add_text(slide, f"▸  {item}", Inches(0.55), y, Inches(5.7), Inches(0.48),
                 font_size=Pt(13), color=WHITE)
        y += Inches(0.5)

    # Right col
    add_rect(slide, Inches(6.6), Inches(1.0), Inches(6.35), Inches(5.9), PANEL, PURPLE, Pt(1))
    add_text(slide, right_title, Inches(6.75), Inches(1.1), Inches(6.1), Inches(0.45),
             font_size=Pt(15), bold=True, color=PURPLE)
    y = Inches(1.62)
    for item in right_items:
        add_text(slide, f"▸  {item}", Inches(6.75), y, Inches(6.0), Inches(0.48),
                 font_size=Pt(13), color=WHITE)
        y += Inches(0.5)

    add_slide_number(slide, slide_num)
    return slide

# ─── Build Presentation ────────────────────────────────────────────────────────
def build_ppt():
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    # ── SLIDE 1: TITLE ─────────────────────────────────────────────────────────
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)

    # Gradient accent bars
    add_rect(slide, 0, 0, SLIDE_W, Inches(0.12), CYAN)
    add_rect(slide, 0, SLIDE_H - Inches(0.12), SLIDE_W, Inches(0.12), PURPLE)
    add_rect(slide, 0, 0, Inches(0.18), SLIDE_H, CYAN)
    add_rect(slide, SLIDE_W - Inches(0.18), 0, Inches(0.18), SLIDE_H, PURPLE)

    # Background glow rectangles (decorative)
    add_rect(slide, Inches(8.5), Inches(1.0), Inches(4.5), Inches(4.5), RGBColor(0x0f, 0x17, 0x2a))

    # Main Title
    add_text(slide, "TG-SPDCL", Inches(0.55), Inches(1.4), Inches(8.0), Inches(1.1),
             font_size=Pt(52), bold=True, color=CYAN, font_name="Segoe UI Black")
    add_text(slide, "ELECTRICITY CONSUMPTION", Inches(0.55), Inches(2.45), Inches(8.5), Inches(0.9),
             font_size=Pt(34), bold=True, color=WHITE, font_name="Segoe UI")
    add_text(slide, "& MONITORING ANALYSIS", Inches(0.55), Inches(3.25), Inches(8.5), Inches(0.9),
             font_size=Pt(34), bold=True, color=PURPLE, font_name="Segoe UI")

    # Subtitle line
    add_rect(slide, Inches(0.55), Inches(4.18), Inches(5.0), Inches(0.04), AMBER)
    add_text(slide, "Data Analytics & Visualization Project  •  July 2026",
             Inches(0.55), Inches(4.3), Inches(8.0), Inches(0.45),
             font_size=Pt(16), color=AMBER, italic=True)

    # Team info box
    add_rect(slide, Inches(0.55), Inches(5.0), Inches(6.5), Inches(2.1), PANEL, CYAN, Pt(1.5))
    add_rect(slide, Inches(0.55), Inches(5.0), Inches(6.5), Inches(0.08), CYAN)
    add_text(slide, "👥  Team Members", Inches(0.75), Inches(5.1), Inches(6.2), Inches(0.4),
             font_size=Pt(13), bold=True, color=CYAN)
    add_text(slide, "D. Chander   •   K. Arjun   •   P. Sachin",
             Inches(0.75), Inches(5.52), Inches(6.2), Inches(0.42),
             font_size=Pt(16), bold=True, color=WHITE)
    add_text(slide, "Roll No: 2503A52214  •  2503A52235  •  2503A52240",
             Inches(0.75), Inches(5.95), Inches(6.2), Inches(0.4),
             font_size=Pt(13), color=MUTED)
    add_text(slide, "Source: TG-SPDCL Domestic Consumption Dataset",
             Inches(0.75), Inches(6.38), Inches(6.2), Inches(0.4),
             font_size=Pt(12), color=MUTED, italic=True)

    # Stats preview on right
    stats = [("15,947", "Total Records"), ("20", "Circles"), ("1,245 GWh", "Total Energy"),
             ("92.68%", "Billing Rate")]
    for i, (val, lbl) in enumerate(stats):
        x = Inches(9.0) + (i % 2) * Inches(2.1)
        y = Inches(2.2) + (i // 2) * Inches(1.8)
        c = [CYAN, EMERALD, PURPLE, AMBER][i]
        add_rect(slide, x, y, Inches(1.9), Inches(1.55), PANEL, c, Pt(1.5))
        add_text(slide, val, x + Inches(0.1), y + Inches(0.15), Inches(1.7), Inches(0.85),
                 font_size=Pt(26), bold=True, color=c, align=PP_ALIGN.CENTER)
        add_text(slide, lbl, x + Inches(0.1), y + Inches(1.0), Inches(1.7), Inches(0.45),
                 font_size=Pt(11), color=MUTED, align=PP_ALIGN.CENTER)

    # ── SLIDE 2: AGENDA ────────────────────────────────────────────────────────
    bullet_slide(prs, "Project Agenda",
        ["Dataset Overview & Source",
         "Dataset Structure — Rows, Columns, Data Types",
         "Key Performance Indicators (KPI Summary)",
         "Circle-Wise Consumption Rankings",
         "Top 20 High-Demand Areas",
         "Billing Efficiency Analysis",
         "Load vs Consumption Correlation",
         "Statistical Distributions & KDE",
         "Seaborn Correlation Heatmap",
         "Data Quality Issues & Outliers",
         "Limitations & Conclusions"],
        icon="📋", accent=CYAN, slide_num=2)

    # ── SLIDE 3: DATASET OVERVIEW ──────────────────────────────────────────────
    bullet_slide(prs, "Dataset Overview & Source",
        [("Source", "TG-SPDCL (Telangana Southern Power Distribution Company Ltd)"),
         ("Period", "July 2026  •  Domestic Consumer Category Only"),
         ("Rows", "15,947 sector-level records"),
         ("Columns", "11 columns (6 categorical, 5 numerical)"),
         ("Coverage", "20 Circles  •  49 Divisions  •  131 Subdivisions  •  476 Sections"),
         ("Category", "DOMESTIC only — No commercial or industrial segments"),
         ("Missing Values", "247 in BILLDSERVICES, 47 in AREA"),
         ("Negative Values", "15 rows with negative UNITS (billing corrections)")],
        icon="🗃️", accent=EMERALD, slide_num=3,
        subtitle="Each row = one electricity consumption record for a specific area & consumer category")

    # ── SLIDE 4: COLUMN DEFINITIONS ────────────────────────────────────────────
    bullet_slide(prs, "Dataset Column Descriptions",
        [("CIRCLE", "Top-level geographic region (20 circles)"),
         ("DIVISION", "Sub-region within a circle (49 divisions)"),
         ("SUBDIVISION", "Subdivision within division (131 total)"),
         ("SECTION / AREA", "Smallest geographic unit — neighborhood/locality (476 sections)"),
         ("CATDESC / CATCODE", "Consumer category — all DOMESTIC (code=1) in this dataset"),
         ("TOTSERVICES", "Total registered electricity connections (integer)"),
         ("BILLDSERVICES", "Number of actually billed connections (float, has 247 NULLs)"),
         ("UNITS", "Total electricity consumed in kWh (float, can be negative)"),
         ("LOAD", "Sanctioned connected load in kW (float, decimal values)")],
        icon="📋", accent=AMBER, slide_num=4,
        subtitle="Categorical: CIRCLE, DIVISION, SUBDIVISION, SECTION, AREA, CATDESC  |  Numerical: CATCODE, TOTSERVICES, BILLDSERVICES, UNITS, LOAD")

    # ── SLIDE 5: KPI SUMMARY ───────────────────────────────────────────────────
    kpi_slide(prs, "Key Performance Indicators — July 2026",
        [("Total Records",        "15,947",     "Sector-level area records"),
         ("Electricity Circles",  "20",          "20 circles, 49 divisions"),
         ("Total Connections",    "8,867,827",   "Registered domestic services"),
         ("Billed Connections",   "8,218,867",   "Actually billed services"),
         ("Billing Efficiency",   "92.68%",      "Target: 95%+  |  Gap: ~649K unbilled"),
         ("Total Energy (GWh)",   "1,245.52",    "1.24 billion kWh total monthly"),
         ("Total Load (MW)",      "13,778.82",   "Sanctioned connected capacity"),
         ("Avg/Household",        "151.5 kWh",   "Per billed connection monthly")],
        accent=CYAN, slide_num=5)

    # ── SLIDE 6: CIRCLE RANKINGS ────────────────────────────────────────────────
    two_col_slide(prs, "Circle-Wise Energy Consumption Rankings",
        "🏆 Top 5 Circles (Highest)",
        ["Medchal — 141.7M units (11.4%)",
         "Habsiguda — 138.1M units (11.1%)",
         "Cybercity — 125.4M units (10.1%)",
         "Hyderabad Central — 112.5M (9.0%)",
         "Saroornagar — 112.3M units (9.0%)",
         "Top 10 circles ≈ 90% of all units"],
        "📉 Bottom 5 Circles (Lowest)",
        ["Gadwal — 10.0M units",
         "Wanaparthy — 10.1M units",
         "Nagarkurnool — 13.1M units",
         "Medak — 15.9M units",
         "Vikarabad — 17.9M units",
         "District circles use ~⅓ per service vs Banjara Hills"],
        accent=CYAN, slide_num=6)

    # ── SLIDE 7: AVG PER SERVICE ────────────────────────────────────────────────
    bullet_slide(prs, "Average Consumption Per Billed Connection",
        ["Banjara Hills — ~240 units/service  (luxury residential, high-value areas)",
         "Cybercity — ~219 units/service  (tech/IT zone proximity)",
         "Hyderabad Central — ~217 units/service",
         "Secunderabad — ~194 units/service",
         "District circles (Wanaparthy, Nagarkurnool, Vikarabad) — ~75 to 83 units",
         "Hyderabad-region circles use ~2x more per service than district circles",
         "Top divisions: Kukatpally (62M), Sainikpuri (55.2M), Saroornagar (52.8M)",
         "Largest single areas: Hydernagar (2.32M), Moosapet (1.81M), Vinayaknagar (1.47M)"],
        icon="🔌", accent=PURPLE, slide_num=7,
        subtitle="Urban-rural consumption gap is stark — per-service average reveals intensity better than totals")

    # ── SLIDE 8: CHART — CIRCLE CONSUMPTION ─────────────────────────────────────
    chart_slide(prs, "Circle-Wise Consumption (GWh) & Sanctioned Load (MW)",
        "circle_consumption.png",
        "Cyan bars = Energy consumed (GWh) | Purple line = Connected load (MW) — sorted highest to lowest",
        accent=CYAN, slide_num=8)

    # ── SLIDE 9: CHART — TOP 20 AREAS ────────────────────────────────────────────
    chart_slide(prs, "Top 20 Highest Electricity Consuming Areas",
        "top20_areas.png",
        "Seaborn barplot with hue by circle — Top areas are Hydernagar, Moosapet & Vinayaknagar",
        accent=AMBER, slide_num=9)

    # ── SLIDE 10: CHART — BILLING EFFICIENCY ──────────────────────────────────────
    chart_slide(prs, "Circle Billing Efficiency (%) vs 95% Benchmark",
        "billing_efficiency.png",
        "Seaborn palette — Amber dashed line = 95% target. Red/orange bars indicate circles needing attention.",
        accent=ROSE, slide_num=10)

    # ── SLIDE 11: CHART — LOAD VS UNITS ──────────────────────────────────────────
    chart_slide(prs, "Connected Load vs Units Consumed — Seaborn Regplot",
        "load_vs_units.png",
        "Strong positive linear correlation (r ≈ 0.95) — Areas with higher load reliably show higher consumption",
        accent=EMERALD, slide_num=11)

    # ── SLIDE 12: CHART — DISTRIBUTIONS ──────────────────────────────────────────
    chart_slide(prs, "Statistical Distributions — Seaborn Histplot + KDE",
        "distribution_metrics.png",
        "Left: Household consumption is right-skewed (median ~90 kWh) | Right: Billing efficiency peaks near 97–100%",
        accent=PURPLE, slide_num=12)

    # ── SLIDE 13: DATA QUALITY & OUTLIERS ──────────────────────────────────────────
    two_col_slide(prs, "Data Quality Issues & Outliers",
        "⚠️ Data Issues Found",
        ["247 missing BILLDSERVICES — all have 0 UNITS (inactive areas)",
         "47 missing AREA names — filled as 'UNKNOWN AREA'",
         "15 rows with negative UNITS (billing corrections/adjustments)",
         "1,999 duplicate AREA names across different sections",
         "Largest record: 2.32M units vs median 37,000 (extreme outlier)",
         "Load range: 0 kW to 33,617 kW (household averages vary widely)"],
        "🔧 Cleaning Strategy",
        ["NaN BILLDSERVICES → imputed from TOTSERVICES",
         "Negative UNITS → clipped to 0 (treated as corrections)",
         "Outliers retained with flagging — valid high-consumption areas",
         "Duplicates are geographic labels, NOT duplicate records",
         "1.5 × IQR outlier method used for identification",
         "Zero-consumption rows kept — may be seasonal / inactive"],
        accent=ROSE, slide_num=13)

    # ── SLIDE 14: KEY PATTERNS & INSIGHTS ──────────────────────────────────────────
    bullet_slide(prs, "Key Patterns & Analytical Insights",
        ["Consumption is strongly urban — top 10 Hyderabad circles = ~90% of all units",
         "Top 10% of areas by size = ~43% of total electricity consumed",
         "Areas labelled 'DOMESTIC HT', 'CT', 'HV' show 3,000–11,500 units/service (bulk connections)",
         "Several THANDA (tribal hamlet) areas in Nalgonda/Suryapet show near-zero consumption",
         "Units & Load are highly correlated (r ≈ 0.95) — load is a reliable predictor",
         "Billing efficiency gap: 92.68% vs 95% target → ~649,000 connections unbilled",
         "The 2,452 largest areas (~15% of all) account for 50.5% of total units",
         "Sort descending by per-service avg reveals 'HT/HV' bulk consumers billed as domestic"],
        icon="🔍", accent=CYAN, slide_num=14,
        subtitle="Pattern: Consumption is concentrated — a small fraction of areas drives the majority of demand")

    # ── SLIDE 15: TECHNOLOGY STACK ──────────────────────────────────────────────
    two_col_slide(prs, "Technology Stack & Tools Used",
        "🐍 Python Libraries",
        ["Pandas — data loading, cleaning, aggregation",
         "NumPy — vectorized feature engineering",
         "Matplotlib — publication-quality static plots",
         "Seaborn — statistical visualizations & KDE",
         "python-pptx — automated PPT generation",
         "json — dashboard data export"],
        "🌐 Web Dashboard",
        ["HTML5 + CSS3 — dark-mode glassmorphism UI",
         "Chart.js — interactive bar, line, scatter charts",
         "JavaScript — dynamic filtering & pagination",
         "KPI Cards — real-time metric calculations",
         "Anomaly Alerts — billing efficiency flags",
         "CSV Export — one-click filtered data download"],
        accent=EMERALD, slide_num=15)

    # ── SLIDE 16: LIMITATIONS ────────────────────────────────────────────────────
    bullet_slide(prs, "Dataset Limitations",
        ["No time-series data — only July 2026 snapshot; cannot track trends over time",
         "Domestic only — commercial & industrial sectors excluded; total demand picture incomplete",
         "No weather/temperature data to explain seasonal consumption peaks",
         "No household-level info — income, family size, or appliance usage unknown",
         "No socioeconomic data — cannot correlate consumption with poverty/development level",
         "1,999 repeated AREA names — geographic deduplication needed for spatial analysis",
         "Negative UNITS (~15 rows) — billing adjustments, not actual negative consumption",
         "HT/HV 'domestic' rows — bulk consumers may inflate domestic category averages"],
        icon="⚠️", accent=ROSE, slide_num=16,
        subtitle="Limitations define the boundary of conclusions — always qualify findings accordingly")

    # ── SLIDE 17: CONCLUSION & GITHUB ────────────────────────────────────────────
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    add_rect(slide, 0, 0, SLIDE_W, Inches(0.12), CYAN)
    add_rect(slide, 0, SLIDE_H - Inches(0.12), SLIDE_W, Inches(0.12), PURPLE)
    add_rect(slide, 0, 0, Inches(0.18), SLIDE_H, EMERALD)

    add_text(slide, "✅  Conclusion & Summary", Inches(0.45), Inches(0.22),
             Inches(12.5), Inches(0.7), font_size=Pt(30), bold=True, color=WHITE)
    add_rect(slide, Inches(0.4), Inches(0.95), Inches(12.5), Inches(0.04), EMERALD)

    conclusions = [
        "Urban concentration: Top 10 Hyderabad circles = 90% of total consumption",
        "Per-service average reveals Banjara Hills (240 kWh) is ~3× higher than district circles (~80 kWh)",
        "Billing efficiency at 92.68% — 649,000 unbilled connections represent revenue loss opportunity",
        "Load and consumption are strongly correlated (r ≈ 0.95) — load data is a reliable proxy",
        "Top 15% of areas (by size) consume over 50% of all electricity",
    ]
    y = Inches(1.1)
    for c in conclusions:
        add_rect(slide, Inches(0.4), y, Inches(12.4), Inches(0.54), PANEL, EMERALD, Pt(1))
        add_text(slide, f"✔  {c}", Inches(0.6), y + Pt(5), Inches(12.0), Inches(0.5),
                 font_size=Pt(13.5), color=WHITE)
        y += Inches(0.62)

    # GitHub link
    add_rect(slide, Inches(0.4), Inches(5.05), Inches(12.4), Inches(0.85), RGBColor(0x1e, 0x29, 0x3b), CYAN, Pt(1.5))
    add_text(slide, "🔗  GitHub Repository:",
             Inches(0.6), Inches(5.12), Inches(3.5), Inches(0.55), font_size=Pt(13), bold=True, color=CYAN)
    add_text(slide, "https://github.com/arjunkanukuntla/dav-project",
             Inches(3.8), Inches(5.12), Inches(9.0), Inches(0.55), font_size=Pt(13), color=WHITE)

    add_text(slide, "Thank You  •  D.Chander  •  K.Arjun  •  P.Sachin",
             Inches(0.4), Inches(6.15), Inches(12.5), Inches(0.5),
             font_size=Pt(18), bold=True, color=AMBER, align=PP_ALIGN.CENTER)
    add_text(slide, "Roll No: 2503A52214  •  2503A52235  •  2503A52240",
             Inches(0.4), Inches(6.65), Inches(12.5), Inches(0.4),
             font_size=Pt(13), color=MUTED, align=PP_ALIGN.CENTER)

    add_slide_number(slide, 17)

    # ── SAVE ──────────────────────────────────────────────────────────────────
    out_path = "TG_SPDCL_Electricity_Analysis.pptx"
    prs.save(out_path)
    print(f"\n✅  Presentation saved: {out_path}")
    print(f"   Slides: 17  |  Size: Widescreen 13.33 x 7.5 inches")
    return out_path

if __name__ == "__main__":
    build_ppt()
