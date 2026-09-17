import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import json
import os

# Set Matplotlib style for sleek dark-themed charts
plt.style.use('dark_background')
plt.rcParams['font.sans-serif'] = 'Segoe UI'
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['axes.edgecolor'] = '#334155'
plt.rcParams['axes.linewidth'] = 1.2
plt.rcParams['grid.color'] = '#1e293b'
plt.rcParams['grid.linestyle'] = '--'
plt.rcParams['grid.alpha'] = 0.6

def run_analysis():
    csv_file = 'TG-SPDCL_consumption_detail_domestic_07_2026.csv'
    print(f"Loading dataset: {csv_file}...")
    df = pd.read_csv(csv_file)
    
    # 1. Data Cleaning & Cleaning Nulls
    df['area'] = df['area'].fillna('UNKNOWN AREA')
    df['billdservices'] = df['billdservices'].fillna(df['totservices'])
    
    # Handle negative units if any (data entry anomalies)
    df['units'] = df['units'].apply(lambda x: max(0, x))
    df['load'] = df['load'].apply(lambda x: max(0.01, x))
    df['totservices'] = df['totservices'].apply(lambda x: max(1, x))
    df['billdservices'] = df['billdservices'].apply(lambda x: max(0, x))
    
    # 2. Feature Engineering
    df['unbilled_services'] = df['totservices'] - df['billdservices']
    df['billing_rate_pct'] = np.where(df['totservices'] > 0, (df['billdservices'] / df['totservices']) * 100, 0)
    df['billing_rate_pct'] = df['billing_rate_pct'].clip(0, 100)
    
    df['units_per_billed_service'] = np.where(df['billdservices'] > 0, df['units'] / df['billdservices'], 0)
    df['units_per_kw_load'] = np.where(df['load'] > 0, df['units'] / df['load'], 0)
    df['load_per_service'] = df['load'] / df['totservices']
    
    print("\n--- Overall Dataset Summary ---")
    total_records = len(df)
    total_circles = df['circle'].nunique()
    total_divisions = df['division'].nunique()
    total_services = int(df['totservices'].sum())
    total_billed_services = int(df['billdservices'].sum())
    total_units_gwh = float(df['units'].sum() / 1e6)
    total_load_mw = float(df['load'].sum() / 1000)
    overall_billing_rate = float((total_billed_services / total_services) * 100)
    avg_units_per_billed = float(df['units'].sum() / total_billed_services) if total_billed_services > 0 else 0
    
    print(f"Total Sector Records: {total_records:,}")
    print(f"Total Circles: {total_circles}")
    print(f"Total Divisions: {total_divisions}")
    print(f"Total Services: {total_services:,}")
    print(f"Total Billed Services: {total_billed_services:,}")
    print(f"Overall Billing Rate: {overall_billing_rate:.2f}%")
    print(f"Total Energy Consumed: {total_units_gwh:.2f} GWh ({df['units'].sum():,.0f} kWh)")
    print(f"Total Connected Load: {total_load_mw:.2f} MW ({df['load'].sum():,.2f} kW)")
    print(f"Avg Monthly Consumption per Billed Connection: {avg_units_per_billed:.1f} kWh")
    
    # 3. Circle Aggregations
    circle_summary = df.groupby('circle').agg(
        records=('area', 'count'),
        totservices=('totservices', 'sum'),
        billdservices=('billdservices', 'sum'),
        units=('units', 'sum'),
        load=('load', 'sum')
    ).reset_index()
    
    circle_summary['billing_rate_pct'] = (circle_summary['billdservices'] / circle_summary['totservices']) * 100
    circle_summary['units_gwh'] = circle_summary['units'] / 1e6
    circle_summary['avg_units_per_service'] = circle_summary['units'] / circle_summary['billdservices']
    circle_summary['load_mw'] = circle_summary['load'] / 1000
    circle_summary = circle_summary.sort_values(by='units', ascending=False)
    
    # Top 20 Areas
    top_20_areas = df.sort_values(by='units', ascending=False).head(20)[[
        'area', 'circle', 'division', 'totservices', 'billdservices', 'units', 'load', 'billing_rate_pct'
    ]]
    
    # Anomaly Detection
    low_billing_sections = df[df['billing_rate_pct'] < 80].sort_values(by='unbilled_services', ascending=False).head(15)[[
        'circle', 'division', 'section', 'area', 'totservices', 'billdservices', 'unbilled_services', 'billing_rate_pct'
    ]]
    
    high_intensity_areas = df[df['units_per_kw_load'] > 250].sort_values(by='units_per_kw_load', ascending=False).head(15)[[
        'circle', 'area', 'units', 'load', 'units_per_kw_load', 'units_per_billed_service'
    ]]
    
    # 4. Generate Matplotlib Charts
    print("\nGenerating Matplotlib Visualizations...")
    
    # Figure 1: Circle Consumption & Load
    fig, ax1 = plt.subplots(figsize=(14, 7), facecolor='#0b0f19')
    ax1.set_facecolor('#0f172a')
    
    circles = circle_summary['circle']
    units_gwh = circle_summary['units_gwh']
    load_mw = circle_summary['load_mw']
    
    x = np.arange(len(circles))
    width = 0.4
    
    rects1 = ax1.bar(x - width/2, units_gwh, width, label='Consumption (GWh)', color='#00f2fe', alpha=0.85, edgecolor='#38bdf8')
    
    ax2 = ax1.twinx()
    rects2 = ax2.bar(x + width/2, load_mw, width, label='Connected Load (MW)', color='#a855f7', alpha=0.85, edgecolor='#c084fc')
    
    ax1.set_xlabel('Electricity Circle', fontsize=12, fontweight='bold', color='#f8fafc', labelpad=10)
    ax1.set_ylabel('Energy Consumed (GWh)', fontsize=12, fontweight='bold', color='#00f2fe', labelpad=10)
    ax2.set_ylabel('Sanctioned Load (MW)', fontsize=12, fontweight='bold', color='#c084fc', labelpad=10)
    ax1.set_title('TG-SPDCL Domestic Electricity Consumption & Load by Circle (July 2026)', fontsize=14, fontweight='bold', color='#f8fafc', pad=15)
    
    ax1.set_xticks(x)
    ax1.set_xticklabels(circles, rotation=45, ha='right', color='#cbd5e1', fontsize=10)
    ax1.grid(True, axis='y', color='#334155', alpha=0.5)
    
    # Combined Legend
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper right', facecolor='#1e293b', edgecolor='#475569')
    
    plt.tight_layout()
    chart1_path = 'circle_consumption.png'
    plt.savefig(chart1_path, dpi=300, bbox_inches='tight', facecolor='#0b0f19')
    plt.close()
    print(f"Saved: {chart1_path}")
    
    # Figure 2: Top 20 High Consumption Areas
    fig, ax = plt.subplots(figsize=(12, 8), facecolor='#0b0f19')
    ax.set_facecolor('#0f172a')
    
    top20_sorted = top_20_areas.sort_values(by='units', ascending=True)
    areas_clean = [a[:28] + '...' if len(a) > 28 else a for a in top20_sorted['area']]
    units_mwh = top20_sorted['units'] / 1000
    
    bars = ax.barh(areas_clean, units_mwh, color='#38bdf8', edgecolor='#0284c7', alpha=0.9)
    
    # Gradient coloring top 3
    for i, bar in enumerate(bars):
        if i >= len(bars) - 3:
            bar.set_color('#f43f5e')
            bar.set_edgecolor('#fb7185')
        elif i >= len(bars) - 7:
            bar.set_color('#fbbf24')
            bar.set_edgecolor('#fde047')
            
    ax.set_xlabel('Energy Consumed (MWh / thousand units)', fontsize=11, fontweight='bold', color='#f8fafc')
    ax.set_title('Top 20 Electricity Consuming Areas in TG-SPDCL (Domestic)', fontsize=14, fontweight='bold', color='#f8fafc', pad=15)
    ax.grid(True, axis='x', color='#334155', alpha=0.5)
    
    for bar in bars:
        width = bar.get_width()
        ax.text(width + 5, bar.get_y() + bar.get_height()/2, f'{width:,.0f} MWh', 
                va='center', ha='left', color='#cbd5e1', fontsize=8, fontweight='bold')
        
    plt.tight_layout()
    chart2_path = 'top20_areas.png'
    plt.savefig(chart2_path, dpi=300, bbox_inches='tight', facecolor='#0b0f19')
    plt.close()
    print(f"Saved: {chart2_path}")
    
    # Figure 3: Billing Efficiency Distribution across Circles
    fig, ax = plt.subplots(figsize=(12, 6), facecolor='#0b0f19')
    ax.set_facecolor('#0f172a')
    
    circle_eff = circle_summary.sort_values(by='billing_rate_pct', ascending=True)
    y_pos = np.arange(len(circle_eff))
    
    colors = ['#ef4444' if rate < 93 else '#3b82f6' if rate < 97 else '#10b981' for rate in circle_eff['billing_rate_pct']]
    bars = ax.barh(y_pos, circle_eff['billing_rate_pct'], color=colors, alpha=0.85, height=0.6)
    
    ax.axvline(x=95, color='#f59e0b', linestyle='--', linewidth=1.5, label='Target Benchmark (95%)')
    ax.set_yticks(y_pos)
    ax.set_yticklabels(circle_eff['circle'], color='#cbd5e1', fontsize=10)
    ax.set_xlim(80, 101)
    ax.set_xlabel('Billing Efficiency (% Billed vs Total Connections)', fontsize=11, fontweight='bold', color='#f8fafc')
    ax.set_title('Billing Efficiency (%) by Electricity Circle', fontsize=14, fontweight='bold', color='#f8fafc', pad=15)
    ax.grid(True, axis='x', color='#334155', alpha=0.5)
    ax.legend(loc='lower right', facecolor='#1e293b', edgecolor='#475569')
    
    for bar in bars:
        w = bar.get_width()
        ax.text(w + 0.3, bar.get_y() + bar.get_height()/2, f'{w:.1f}%', va='center', color='#f8fafc', fontsize=9, fontweight='bold')
        
    plt.tight_layout()
    chart3_path = 'billing_efficiency.png'
    plt.savefig(chart3_path, dpi=300, bbox_inches='tight', facecolor='#0b0f19')
    plt.close()
    print(f"Saved: {chart3_path}")

    # Figure 4: Connected Load vs Consumed Units Scatter Plot
    fig, ax = plt.subplots(figsize=(10, 6), facecolor='#0b0f19')
    ax.set_facecolor('#0f172a')
    
    # Filter reasonable range for plot visualization
    scatter_df = df[(df['load'] < 5000) & (df['units'] < 500000)].sample(min(3000, len(df)), random_state=42)
    
    scatter = ax.scatter(scatter_df['load'], scatter_df['units'] / 1000, 
                         c=scatter_df['billing_rate_pct'], cmap='viridis', 
                         alpha=0.6, edgecolors='none', s=25)
    
    cbar = plt.colorbar(scatter, ax=ax)
    cbar.set_label('Billing Rate (%)', color='#f8fafc', fontweight='bold')
    cbar.ax.yaxis.set_tick_params(color='#cbd5e1')
    plt.setp(plt.getp(cbar.ax.axes, 'yticklabels'), color='#cbd5e1')
    
    # Trendline
    z = np.polyfit(scatter_df['load'], scatter_df['units'] / 1000, 1)
    p = np.poly1d(z)
    ax.plot(scatter_df['load'], p(scatter_df['load']), color='#ef4444', linestyle='-', linewidth=2, label=f'Linear Trend (Slope: {z[0]:.2f} MWh/kW)')
    
    ax.set_xlabel('Connected Load (kW)', fontsize=11, fontweight='bold', color='#f8fafc')
    ax.set_ylabel('Units Consumed (MWh)', fontsize=11, fontweight='bold', color='#f8fafc')
    ax.set_title('Correlation: Connected Load (kW) vs Energy Consumed (MWh)', fontsize=13, fontweight='bold', color='#f8fafc', pad=15)
    ax.grid(True, color='#334155', alpha=0.5)
    ax.legend(loc='upper left', facecolor='#1e293b', edgecolor='#475569')
    
    plt.tight_layout()
    chart4_path = 'load_vs_units.png'
    plt.savefig(chart4_path, dpi=300, bbox_inches='tight', facecolor='#0b0f19')
    plt.close()
    print(f"Saved: {chart4_path}")

    # Figure 5: Distribution Metrics
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5), facecolor='#0b0f19')
    ax1.set_facecolor('#0f172a')
    ax2.set_facecolor('#0f172a')
    
    # Histogram 1: Units per Billed Connection
    valid_units_per_conn = df[(df['units_per_billed_service'] > 0) & (df['units_per_billed_service'] < 1000)]['units_per_billed_service']
    ax1.hist(valid_units_per_conn, bins=40, color='#00f2fe', edgecolor='#0284c7', alpha=0.75)
    ax1.axvline(valid_units_per_conn.median(), color='#f43f5e', linestyle='--', linewidth=2, label=f'Median: {valid_units_per_conn.median():.0f} kWh')
    ax1.set_title('Distribution of Monthly Consumption per Household (kWh)', color='#f8fafc', fontweight='bold')
    ax1.set_xlabel('Monthly Consumption per Billed Connection (kWh)', color='#cbd5e1')
    ax1.set_ylabel('Sector Frequency', color='#cbd5e1')
    ax1.grid(True, color='#334155', alpha=0.5)
    ax1.legend(facecolor='#1e293b', edgecolor='#475569')
    
    # Histogram 2: Billing Rate %
    ax2.hist(df['billing_rate_pct'], bins=30, color='#10b981', edgecolor='#059669', alpha=0.75)
    ax2.axvline(df['billing_rate_pct'].mean(), color='#fbbf24', linestyle='--', linewidth=2, label=f'Mean: {df["billing_rate_pct"].mean():.1f}%')
    ax2.set_title('Distribution of Section Billing Efficiency (%)', color='#f8fafc', fontweight='bold')
    ax2.set_xlabel('Billing Efficiency Rate (%)', color='#cbd5e1')
    ax2.set_ylabel('Sector Frequency', color='#cbd5e1')
    ax2.grid(True, color='#334155', alpha=0.5)
    ax2.legend(facecolor='#1e293b', edgecolor='#475569')
    
    plt.tight_layout()
    chart5_path = 'distribution_metrics.png'
    plt.savefig(chart5_path, dpi=300, bbox_inches='tight', facecolor='#0b0f19')
    plt.close()
    print(f"Saved: {chart5_path}")

    # 5. Export JSON Summaries for Dashboard
    summary_export = {
        "kpis": {
            "total_records": total_records,
            "total_circles": total_circles,
            "total_divisions": total_divisions,
            "total_services": total_services,
            "total_billed_services": total_billed_services,
            "unbilled_services": total_services - total_billed_services,
            "total_units_gwh": round(total_units_gwh, 2),
            "total_units_kwh": float(df['units'].sum()),
            "total_load_mw": round(total_load_mw, 2),
            "overall_billing_rate": round(overall_billing_rate, 2),
            "avg_units_per_billed": round(avg_units_per_billed, 1)
        },
        "circles": circle_summary.to_dict(orient='records'),
        "top_20_areas": top_20_areas.to_dict(orient='records'),
        "low_billing_sections": low_billing_sections.to_dict(orient='records'),
        "high_intensity_areas": high_intensity_areas.to_dict(orient='records'),
        "sample_records": df.head(1000).to_dict(orient='records')  # Fast payload sample
    }
    
    with open('analysis_summary.json', 'w') as f:
        json.dump(summary_export, f, indent=2)
    print("Exported analysis summary to 'analysis_summary.json'")
    print("Analysis pipeline executed successfully!")

if __name__ == '__main__':
    run_analysis()
