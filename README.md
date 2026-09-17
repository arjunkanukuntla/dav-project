# TG-SPDCL Electricity Consumption Analytics & Dashboard

An end-to-end data analysis pipeline and interactive web monitoring dashboard for the **TG-SPDCL Domestic Electricity Consumption** dataset (`TG-SPDCL_consumption_detail_domestic_07_2026.csv`).

---

## 📊 Project Overview

- **Dataset Scope**: 15,947 domestic sector records across 20 Electricity Circles and 49 Divisions.
- **Key Metrics Analyzed**:
  - Total Active Domestic Connections: **8,867,827**
  - Billed Domestic Connections: **8,218,867** (Overall Billing Efficiency: **92.68%**)
  - Total Monthly Energy Consumed: **1,245.52 GWh** (1,245,523,023 kWh)
  - Total Sanctioned Connected Load: **13,778.82 MW** (13,778,816 kW)
  - Avg Monthly Household Usage: **151.5 kWh** / billed connection

---

## 📁 Repository Structure

```text
├── TG-SPDCL_consumption_detail_domestic_07_2026.csv   # Domestic consumption dataset
├── main.py                                            # Pandas analysis & Matplotlib chart generator
├── analysis_summary.json                              # Exported analytics payload for dashboard
├── index.html                                         # Dashboard HTML structure
├── styles.css                                         # Dark-mode glassmorphism design system
├── app.js                                             # Interactive Chart.js & data controller
├── circle_consumption.png                             # Matplotlib Plot: Energy vs Load by Circle
├── top20_areas.png                                    # Matplotlib Plot: Top 20 High Demand Sectors
├── billing_efficiency.png                             # Matplotlib Plot: Circle Billing Rates vs 95% Benchmark
├── load_vs_units.png                                  # Matplotlib Plot: Connected Load vs Units Scatter
└── distribution_metrics.png                           # Matplotlib Plot: Consumption & Billing Histograms
```

---

## 🚀 Quick Start Guide

### 1. Run Data Analysis & Plot Generator (Python)

Ensure Python 3.x is installed, then run:

```bash
python main.py
```

This script cleans data, calculates derived indicators, generates 5 high-resolution Matplotlib PNG charts, and outputs `analysis_summary.json`.

### 2. Launch Web Dashboard

To launch the live interactive web dashboard locally:

```bash
python -m http.server 8000
```

Open your browser and navigate to: **[http://localhost:8000](http://localhost:8000)**

---

## 🌟 Dashboard Features

- **KPI Cards**: Live stats for Total Units (GWh), Active Services, Billing Efficiency %, and Sanctioned Load (MW).
- **Interactive Slicers**: Slicer dropdowns for Circle (20), Division (49), and real-time Search.
- **Chart.js Visualizations**: Bar/Line combo charts, billing rate benchmarks, scatter plots, and top demand areas.
- **Anomaly Detection Feed**: Automated alerts for low billing efficiency sections (< 85%) and high load intensity hotspots (> 250 kWh/kW).
- **Sector Data Explorer**: Paginated, searchable table with one-click **CSV Export**.
- **Matplotlib Report Gallery**: Modal previewer for python-generated publication plots.
