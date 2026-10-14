import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns

# Set style
plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
plt.rcParams['font.sans-serif'] = 'Arial'
plt.rcParams['font.family'] = 'sans-serif'

NAVY = '#0a193b'
TEAL = '#00b4d8'
AMBER = '#f59e0b'
GREEN = '#10b981'
RED = '#ef4444'
SLATE = '#64748b'
LIGHT_BG = '#f8fafc'

excel_path = 'climate_weather_trend_dataset.xlsx'
df_global = pd.read_excel(excel_path, sheet_name='Global_Annual').iloc[:, :5]
df_city = pd.read_excel(excel_path, sheet_name='City_Monthly')
df_events = pd.read_excel(excel_path, sheet_name='Extreme_Events')

# 1. Slide 4: Regression Plot (CO2 vs Sea Level Anomaly & Temp Anomaly)
fig, ax1 = plt.subplots(figsize=(7, 4.2), dpi=200)
fig.patch.set_facecolor('#ffffff')
ax1.set_facecolor('#ffffff')

x = df_global['CO2_ppm']
y = df_global['Sea_Level_Anomaly_mm']
m, b = np.polyfit(x, y, 1)
r = np.corrcoef(x, y)[0, 1]

ax1.scatter(x, y, color=TEAL, alpha=0.85, edgecolors=NAVY, s=55, label='Annual Observations (1950–2024)', zorder=3)
ax1.plot(x, m*x + b, color=NAVY, linewidth=2.5, linestyle='-', label=f'Linear Fit: y = {m:.2f}x - {-b:.1f}\n$R^2$ = {r**2:.4f} (p < 0.001)', zorder=4)

ax1.set_title('Empirical Regression: Atmospheric CO₂ vs Sea-Level Anomaly', fontsize=12, fontweight='bold', color=NAVY, pad=12)
ax1.set_xlabel('Atmospheric CO₂ Concentration (ppm)', fontsize=10, fontweight='bold', color=NAVY)
ax1.set_ylabel('Global Sea-Level Anomaly (mm)', fontsize=10, fontweight='bold', color=NAVY)
ax1.grid(True, linestyle='--', alpha=0.5, color='#cbd5e1')
ax1.legend(loc='upper left', frameon=True, facecolor='#f8fafc', edgecolor='#cbd5e1', fontsize=8.5)

# Callout annotation
peak_x = x.max()
peak_y = y.max()
ax1.annotate(f'2024 Peak: {peak_x:.1f} ppm\nAnomaly: +{peak_y:.1f} mm',
             xy=(peak_x, peak_y), xytext=(peak_x - 35, peak_y - 25),
             arrowprops=dict(facecolor=AMBER, shrink=0.08, width=1.5, headwidth=6),
             bbox=dict(boxstyle="round,pad=0.3", fc="#fef3c7", ec=AMBER, lw=1.2),
             fontsize=8, fontweight='bold', color=NAVY)

plt.tight_layout()
plt.savefig('temp_co2_regression.png', dpi=200, bbox_inches='tight')
plt.close()
print('Created temp_co2_regression.png')

# 2. Slide 5: Excel-Style Dashboard Graphic (Multi-panel: City Climates + Extreme Event Donut)
fig, (ax_col, ax_donut) = plt.subplots(1, 2, figsize=(7.2, 4.3), dpi=200, gridspec_kw={'width_ratios': [1.3, 1]})
fig.patch.set_facecolor('#ffffff')
ax_col.set_facecolor('#ffffff')
ax_donut.set_facecolor('#ffffff')

# City climate clustered bar
cities = ['Chennai', 'London', 'Nairobi', 'NYC', 'Sydney', 'Tokyo']
city_full = ['Chennai, India', 'London, UK', 'Nairobi, Kenya', 'New York, USA', 'Sydney, Australia', 'Tokyo, Japan']
avg_temps = [df_city[df_city['City'] == c]['Avg_Temp_C'].mean() for c in city_full]
avg_precip = [df_city[df_city['City'] == c]['Precipitation_mm'].mean() for c in city_full]

x_indices = np.arange(len(cities))
width = 0.38

rects1 = ax_col.bar(x_indices - width/2, avg_temps, width, label='Avg Temp (°C)', color=AMBER, edgecolor=NAVY, linewidth=0.8)
ax_col_sec = ax_col.twinx()
rects2 = ax_col_sec.bar(x_indices + width/2, avg_precip, width, label='Precipitation (mm)', color=TEAL, edgecolor=NAVY, linewidth=0.8)

ax_col.set_title('Excel BI: City Climate Profiles', fontsize=10.5, fontweight='bold', color=NAVY, pad=8)
ax_col.set_xticks(x_indices)
ax_col.set_xticklabels(cities, rotation=25, fontsize=8.5, fontweight='bold', color=NAVY)
ax_col.set_ylabel('Temp (°C)', fontsize=9, fontweight='bold', color=AMBER)
ax_col_sec.set_ylabel('Precip (mm)', fontsize=9, fontweight='bold', color=TEAL)
ax_col.grid(True, linestyle=':', alpha=0.4, color='#cbd5e1')
ax_col_sec.grid(False)

# Donut chart for extreme events
event_totals = df_events.groupby('Event_Type')['Global_Event_Count'].sum().sort_values(ascending=False)
colors_donut = [AMBER, TEAL, GREEN, '#8b5cf6', RED]
labels_clean = ['Wildfire', 'Heatwave', 'Flood', 'Storm', 'Drought']
wedges, texts, autotexts = ax_donut.pie(
    event_totals, 
    labels=labels_clean, 
    autopct='%1.1f%%',
    startangle=140, 
    colors=colors_donut,
    wedgeprops=dict(width=0.45, edgecolor='#ffffff', linewidth=1.5),
    textprops=dict(fontsize=8, color=NAVY)
)
for at in autotexts:
    at.set_fontsize(7.5)
    at.set_weight('bold')
    at.set_color('#ffffff')

ax_donut.set_title('Excel Pivot: Event Share', fontsize=10.5, fontweight='bold', color=NAVY, pad=8)

plt.tight_layout()
plt.savefig('excel_visualization.png', dpi=200, bbox_inches='tight')
plt.close()
print('Created excel_visualization.png')

# 3. Slide 6 - Figure 1: Dual-Axis Planetary Telemetry
fig, ax1 = plt.subplots(figsize=(6, 3.2), dpi=200)
fig.patch.set_facecolor('#ffffff')
ax1.set_facecolor('#ffffff')

years = df_global['Year']
co2 = df_global['CO2_ppm']
sea = df_global['Sea_Level_Anomaly_mm']
ice = df_global['Arctic_Ice_Extent_Anomaly_M_km2']

line1 = ax1.plot(years, co2, color=NAVY, linewidth=2.2, label='CO₂ Level (ppm)')
ax1.set_ylabel('Atmospheric CO₂ (ppm)', color=NAVY, fontsize=9, fontweight='bold')
ax1.tick_params(axis='y', labelcolor=NAVY, labelsize=8)
ax1.set_xlabel('Observation Year (1950–2024)', fontsize=9, fontweight='bold', color=NAVY)
ax1.tick_params(axis='x', labelsize=8)

ax2 = ax1.twinx()
line2 = ax2.plot(years, sea, color=TEAL, linewidth=2.0, linestyle='--', label='Sea Level Anomaly (mm)')
line3 = ax2.plot(years, ice * 40, color=RED, linewidth=1.8, linestyle=':', label='Arctic Ice Loss (scaled)')
ax2.set_ylabel('Ocean & Polar Anomaly Index', color=TEAL, fontsize=9, fontweight='bold')
ax2.tick_params(axis='y', labelcolor=TEAL, labelsize=8)

lines = line1 + line2 + line3
labels = [l.get_label() for l in lines]
ax1.legend(lines, labels, loc='upper left', frameon=True, facecolor='#f8fafc', edgecolor='#cbd5e1', fontsize=7.5)
ax1.set_title('Fig 1: Dual-Axis Decadal Telemetry (1950–2024)', fontsize=10, fontweight='bold', color=NAVY, pad=8)
ax1.grid(True, linestyle='--', alpha=0.4, color='#cbd5e1')

plt.tight_layout()
plt.savefig('fig1_dual_axis.png', dpi=200, bbox_inches='tight')
plt.close()
print('Created fig1_dual_axis.png')

# 4. Slide 6 - Figure 2: Top Hazard Hotspots (Pareto Ranking)
fig, ax = plt.subplots(figsize=(6, 3.2), dpi=200)
fig.patch.set_facecolor('#ffffff')
ax.set_facecolor('#ffffff')

event_ranks = df_events.groupby('Event_Type')['Global_Event_Count'].sum().sort_values(ascending=True)
colors_bar = [SLATE, '#8b5cf6', TEAL, GREEN, AMBER]

bars = ax.barh(event_ranks.index, event_ranks.values, color=colors_bar, edgecolor=NAVY, height=0.6, linewidth=0.8)
for bar in bars:
    w = bar.get_width()
    pct = (w / event_ranks.values.sum()) * 100
    ax.text(w + 12, bar.get_y() + bar.get_height()/2, f'{int(w)} ({pct:.1f}%)',
            va='center', ha='left', fontsize=8, fontweight='bold', color=NAVY)

ax.set_xlim(0, 700)
ax.set_title('Fig 2: Global Extreme Hazard Severity Hotspots', fontsize=10, fontweight='bold', color=NAVY, pad=8)
ax.set_xlabel('Cumulative Event Count (2000–2024)', fontsize=8.5, fontweight='bold', color=NAVY)
ax.tick_params(axis='both', labelsize=8)
ax.grid(True, axis='x', linestyle='--', alpha=0.4, color='#cbd5e1')

plt.tight_layout()
plt.savefig('fig2_hotspots.png', dpi=200, bbox_inches='tight')
plt.close()
print('Created fig2_hotspots.png')

# 5. Slide 6 - Figure 3: Benchmark vs Baseline Thresholds
fig, ax = plt.subplots(figsize=(6, 3.2), dpi=200)
fig.patch.set_facecolor('#ffffff')
ax.set_facecolor('#ffffff')

# Rolling 5-year mean of temp anomaly
df_global['Temp_Roll5'] = df_global['Global_Temp_Anomaly_C'].rolling(window=5, min_periods=1).mean()
ax.plot(df_global['Year'], df_global['Global_Temp_Anomaly_C'], color='#94a3b8', alpha=0.5, label='Annual Anomaly', linewidth=1)
ax.plot(df_global['Year'], df_global['Temp_Roll5'], color=RED, linewidth=2.2, label='5-Yr Smoothed Trend')

# Threshold bands
ax.axhline(0.15, color=AMBER, linestyle='--', linewidth=1.5, label='High Risk Watch (>+0.15°C)')
ax.axhline(0.00, color=NAVY, linestyle='-', linewidth=1.0, alpha=0.6)
ax.fill_between(df_global['Year'], 0.15, 0.25, color='#fee2e2', alpha=0.4)

ax.set_title('Fig 3: Thermal Anomaly vs Climate Risk Watchline', fontsize=10, fontweight='bold', color=NAVY, pad=8)
ax.set_xlabel('Observation Year', fontsize=8.5, fontweight='bold', color=NAVY)
ax.set_ylabel('Global Temp Anomaly (°C)', fontsize=8.5, fontweight='bold', color=NAVY)
ax.set_ylim(-0.3, 0.25)
ax.tick_params(axis='both', labelsize=8)
ax.legend(loc='lower right', frameon=True, facecolor='#f8fafc', edgecolor='#cbd5e1', fontsize=7.5)
ax.grid(True, linestyle='--', alpha=0.4, color='#cbd5e1')

plt.tight_layout()
plt.savefig('fig3_benchmark.png', dpi=200, bbox_inches='tight')
plt.close()
print('Created fig3_benchmark.png')

# 6. Slide 6 - Figure 4: Seaborn Covariance Heatmap
fig, ax = plt.subplots(figsize=(6, 3.2), dpi=200)
fig.patch.set_facecolor('#ffffff')
ax.set_facecolor('#ffffff')

corr_data = df_global[['Year', 'Global_Temp_Anomaly_C', 'CO2_ppm', 'Sea_Level_Anomaly_mm', 'Arctic_Ice_Extent_Anomaly_M_km2']].copy()
corr_data.columns = ['Year', 'Temp Anom', 'CO₂ ppm', 'Sea Level', 'Arctic Ice']
corr_matrix = corr_data.corr()

sns.heatmap(corr_matrix, annot=True, fmt='.2f', cmap='coolwarm', vmin=-1, vmax=1,
            cbar_kws={'shrink': 0.8}, ax=ax, annot_kws={'fontsize': 8, 'weight': 'bold'},
            linewidths=1, linecolor='#ffffff')

ax.set_title('Fig 4: Seaborn Empirical Correlation Matrix', fontsize=10, fontweight='bold', color=NAVY, pad=8)
ax.tick_params(axis='both', labelsize=7.5, labelcolor=NAVY)

plt.tight_layout()
plt.savefig('fig4_heatmap.png', dpi=200, bbox_inches='tight')
plt.close()
print('Created fig4_heatmap.png')
