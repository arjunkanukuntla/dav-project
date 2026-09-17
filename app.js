/* TG-SPDCL Electricity Consumption Dashboard Controller */

let globalData = {
    kpis: {},
    circles: [],
    topAreas: [],
    lowBilling: [],
    highIntensity: [],
    records: []
};

let filteredRecords = [];
let currentPage = 1;
const itemsPerPage = 25;

// Chart Instances
let chartInstances = {};

document.addEventListener('DOMContentLoaded', async () => {
    console.log("Initializing TG-SPDCL Power Dashboard...");
    await loadData();
    setupEventListeners();
    setupTabs();
});

async function loadData() {
    try {
        const response = await fetch('analysis_summary.json');
        if (!response.ok) throw new Error("Failed to load analysis_summary.json");
        const json = await response.json();
        
        globalData.kpis = json.kpis;
        globalData.circles = json.circles;
        globalData.topAreas = json.top_20_areas;
        globalData.lowBilling = json.low_billing_sections;
        globalData.highIntensity = json.high_intensity_areas;
        globalData.records = json.sample_records || [];
        
        filteredRecords = [...globalData.records];
        
        renderKPIs(globalData.kpis);
        populateFilters();
        renderCharts();
        renderAnomalies();
        renderTable();
        
    } catch (err) {
        console.error("Data load error:", err);
        // Fallback banner if missing file
        document.querySelector('.dashboard-container').innerHTML = `
            <div style="text-align: center; padding: 50px; background: rgba(244,63,94,0.1); border: 1px solid #f43f5e; border-radius: 16px;">
                <h2 style="color: #f43f5e;"><i class="fa-solid fa-triangle-exclamation"></i> Analysis Summary Missing</h2>
                <p style="color: #cbd5e1; margin-top: 10px;">Please ensure <code>python main.py</code> has been executed to generate <code>analysis_summary.json</code>.</p>
            </div>
        `;
    }
}

function renderKPIs(kpis) {
    document.getElementById('kpiTotalUnits').innerHTML = `${(kpis.total_units_gwh).toLocaleString()} <span>GWh</span>`;
    document.getElementById('kpiTotalServices').innerHTML = `${(kpis.total_services).toLocaleString()} <span>Conn</span>`;
    document.getElementById('kpiBilledServices').innerText = `${(kpis.total_billed_services).toLocaleString()} Billed Connections`;
    document.getElementById('kpiBillingRate').innerHTML = `${kpis.overall_billing_rate}%`;
    document.getElementById('kpiUnbilled').innerText = `${(kpis.unbilled_services).toLocaleString()} Unbilled Services`;
    document.getElementById('kpiTotalLoad').innerHTML = `${(kpis.total_load_mw).toLocaleString()} <span>MW</span>`;
    document.getElementById('kpiAvgCons').innerText = `Avg ${kpis.avg_units_per_billed} kWh / Billed Conn`;
}

function populateFilters() {
    const circleSelect = document.getElementById('circleSelect');
    circleSelect.innerHTML = '<option value="ALL">All Circles (20)</option>';
    
    const circleNames = globalData.circles.map(c => c.circle).sort();
    circleNames.forEach(c => {
        const opt = document.createElement('option');
        opt.value = c;
        opt.textContent = c;
        circleSelect.appendChild(opt);
    });
}

function updateDivisionFilter(selectedCircle) {
    const divisionSelect = document.getElementById('divisionSelect');
    divisionSelect.innerHTML = '<option value="ALL">All Divisions</option>';
    
    if (selectedCircle === 'ALL') {
        divisionSelect.disabled = true;
        return;
    }
    
    divisionSelect.disabled = false;
    const divisions = [...new Set(globalData.records
        .filter(r => r.circle === selectedCircle)
        .map(r => r.division)
    )].sort();
    
    divisions.forEach(d => {
        const opt = document.createElement('option');
        opt.value = d;
        opt.textContent = d;
        divisionSelect.appendChild(opt);
    });
}

function setupEventListeners() {
    const circleSelect = document.getElementById('circleSelect');
    const divisionSelect = document.getElementById('divisionSelect');
    const searchInput = document.getElementById('searchInput');
    const btnReset = document.getElementById('btnResetFilters');
    const btnExport = document.getElementById('btnExportCSV');
    const btnRefresh = document.getElementById('btnRefresh');

    circleSelect.addEventListener('change', (e) => {
        updateDivisionFilter(e.target.value);
        applyFilters();
    });

    divisionSelect.addEventListener('change', applyFilters);
    searchInput.addEventListener('input', applyFilters);

    btnReset.addEventListener('click', () => {
        circleSelect.value = 'ALL';
        updateDivisionFilter('ALL');
        searchInput.value = '';
        applyFilters();
    });

    btnExport.addEventListener('click', exportCSV);
    btnRefresh.addEventListener('click', loadData);

    // Pagination
    document.getElementById('btnPrevPage').addEventListener('click', () => {
        if (currentPage > 1) {
            currentPage--;
            renderTable();
        }
    });

    document.getElementById('btnNextPage').addEventListener('click', () => {
        const maxPage = Math.ceil(filteredRecords.length / itemsPerPage);
        if (currentPage < maxPage) {
            currentPage++;
            renderTable();
        }
    });
}

function applyFilters() {
    const selectedCircle = document.getElementById('circleSelect').value;
    const selectedDivision = document.getElementById('divisionSelect').value;
    const searchTerm = document.getElementById('searchInput').value.trim().toLowerCase();

    filteredRecords = globalData.records.filter(r => {
        const matchCircle = selectedCircle === 'ALL' || r.circle === selectedCircle;
        const matchDivision = selectedDivision === 'ALL' || r.division === selectedDivision;
        const matchSearch = !searchTerm || 
            (r.area && r.area.toLowerCase().includes(searchTerm)) || 
            (r.section && r.section.toLowerCase().includes(searchTerm));
        
        return matchCircle && matchDivision && matchSearch;
    });

    currentPage = 1;
    
    // Update Badge
    const badge = document.getElementById('activeFilterBadge');
    if (selectedCircle === 'ALL' && !searchTerm) {
        badge.textContent = 'Showing All Circles';
    } else {
        badge.textContent = `Filtered: ${filteredRecords.length} sectors`;
    }

    // Recalculate KPIs based on filtered subset
    if (filteredRecords.length > 0) {
        const totalUnits = filteredRecords.reduce((sum, r) => sum + r.units, 0);
        const totalServices = filteredRecords.reduce((sum, r) => sum + r.totservices, 0);
        const totalBilled = filteredRecords.reduce((sum, r) => sum + r.billdservices, 0);
        const totalLoad = filteredRecords.reduce((sum, r) => sum + r.load, 0);
        const billingRate = totalServices > 0 ? (totalBilled / totalServices) * 100 : 0;
        const avgCons = totalBilled > 0 ? totalUnits / totalBilled : 0;

        renderKPIs({
            total_units_gwh: round(totalUnits / 1e6, 2),
            total_services: totalServices,
            total_billed_services: totalBilled,
            unbilled_services: totalServices - totalBilled,
            overall_billing_rate: round(billingRate, 2),
            total_load_mw: round(totalLoad / 1000, 2),
            avg_units_per_billed: round(avgCons, 1)
        });
    }

    renderTable();
}

function setupTabs() {
    const tabBtns = document.querySelectorAll('.tab-btn');
    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            tabBtns.forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            
            btn.classList.add('active');
            const tabId = btn.getAttribute('data-tab');
            document.getElementById(tabId).classList.add('active');
        });
    });
}

/* Render Chart.js Visualizations */
function renderCharts() {
    // Destroy existing chart instances if re-rendering
    Object.keys(chartInstances).forEach(key => {
        if (chartInstances[key]) chartInstances[key].destroy();
    });

    const circlesData = globalData.circles;
    const topCircles = circlesData.slice(0, 10); // Top 10 for clean visual

    // Chart 1: Circle Consumption Bar/Line Combo Chart
    const ctx1 = document.getElementById('chartCircleConsumption').getContext('2d');
    chartInstances['circleConsumption'] = new Chart(ctx1, {
        type: 'bar',
        data: {
            labels: topCircles.map(c => c.circle),
            datasets: [
                {
                    label: 'Consumption (GWh)',
                    data: topCircles.map(c => c.units_gwh),
                    backgroundColor: 'rgba(0, 242, 254, 0.75)',
                    borderColor: '#00f2fe',
                    borderWidth: 1,
                    yAxisID: 'y'
                },
                {
                    label: 'Sanctioned Load (MW)',
                    data: topCircles.map(c => c.load_mw),
                    type: 'line',
                    borderColor: '#a855f7',
                    backgroundColor: 'rgba(168, 85, 247, 0.2)',
                    borderWidth: 3,
                    pointBackgroundColor: '#c084fc',
                    tension: 0.3,
                    yAxisID: 'y1'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { labels: { color: '#cbd5e1' } },
                tooltip: { backgroundColor: '#0f172a', titleColor: '#00f2fe' }
            },
            scales: {
                x: { ticks: { color: '#94a3b8' }, grid: { color: '#1e293b' } },
                y: { 
                    type: 'linear', position: 'left',
                    ticks: { color: '#00f2fe' }, grid: { color: '#1e293b' },
                    title: { display: true, text: 'Energy (GWh)', color: '#00f2fe' }
                },
                y1: { 
                    type: 'linear', position: 'right',
                    ticks: { color: '#c084fc' }, grid: { drawOnChartArea: false },
                    title: { display: true, text: 'Load (MW)', color: '#c084fc' }
                }
            }
        }
    });

    // Chart 2: Billing Efficiency by Circle Horizontal Bar
    const ctx2 = document.getElementById('chartBillingEfficiency').getContext('2d');
    const sortedByEff = [...circlesData].sort((a, b) => b.billing_rate_pct - a.billing_rate_pct);
    
    chartInstances['billingEfficiency'] = new Chart(ctx2, {
        type: 'bar',
        data: {
            labels: sortedByEff.map(c => c.circle),
            datasets: [{
                label: 'Billing Efficiency (%)',
                data: sortedByEff.map(c => c.billing_rate_pct),
                backgroundColor: sortedByEff.map(c => c.billing_rate_pct < 93 ? 'rgba(244, 63, 94, 0.8)' : 'rgba(16, 185, 129, 0.8)'),
                borderColor: sortedByEff.map(c => c.billing_rate_pct < 93 ? '#f43f5e' : '#10b981'),
                borderWidth: 1
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: { backgroundColor: '#0f172a' }
            },
            scales: {
                x: { min: 80, max: 100, ticks: { color: '#cbd5e1' }, grid: { color: '#1e293b' } },
                y: { ticks: { color: '#cbd5e1', font: { size: 10 } }, grid: { display: false } }
            }
        }
    });

    // Chart 3: Connected Load vs Consumed Energy Scatter
    const ctx3 = document.getElementById('chartLoadCorrelation').getContext('2d');
    const scatterPoints = globalData.records.slice(0, 500).map(r => ({
        x: r.load,
        y: r.units / 1000
    }));

    chartInstances['loadCorrelation'] = new Chart(ctx3, {
        type: 'scatter',
        data: {
            datasets: [{
                label: 'Sectors (kW vs MWh)',
                data: scatterPoints,
                backgroundColor: 'rgba(56, 189, 248, 0.6)',
                borderColor: '#0284c7',
                pointRadius: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { labels: { color: '#cbd5e1' } }
            },
            scales: {
                x: { title: { display: true, text: 'Load (kW)', color: '#94a3b8' }, ticks: { color: '#cbd5e1' }, grid: { color: '#1e293b' } },
                y: { title: { display: true, text: 'Energy (MWh)', color: '#94a3b8' }, ticks: { color: '#cbd5e1' }, grid: { color: '#1e293b' } }
            }
        }
    });

    // Chart 4: Top 10 High Consumption Areas Horizontal Bar
    const ctx4 = document.getElementById('chartTopAreas').getContext('2d');
    const top10Areas = globalData.topAreas.slice(0, 10);

    chartInstances['topAreas'] = new Chart(ctx4, {
        type: 'bar',
        data: {
            labels: top10Areas.map(a => a.area.length > 20 ? a.area.slice(0, 20) + '...' : a.area),
            datasets: [{
                label: 'Units Consumed (kWh)',
                data: top10Areas.map(a => a.units),
                backgroundColor: 'rgba(245, 158, 11, 0.75)',
                borderColor: '#f59e0b',
                borderWidth: 1
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            },
            scales: {
                x: { ticks: { color: '#cbd5e1' }, grid: { color: '#1e293b' } },
                y: { ticks: { color: '#cbd5e1', font: { size: 10 } }, grid: { display: false } }
            }
        }
    });
}

function renderAnomalies() {
    const lowBillingList = document.getElementById('lowBillingList');
    const highIntensityList = document.getElementById('highIntensityList');

    document.getElementById('lowBillingCount').innerText = `${globalData.lowBilling.length} Identified`;
    document.getElementById('highIntensityCount').innerText = `${globalData.highIntensity.length} Identified`;

    lowBillingList.innerHTML = globalData.lowBilling.map(item => `
        <div class="alert-item">
            <div class="alert-info">
                <h5>${item.area}</h5>
                <p><i class="fa-solid fa-location-dot"></i> ${item.circle} &bull; ${item.section}</p>
                <p>Services: ${item.billdservices} / ${item.totservices} Billed</p>
            </div>
            <div class="alert-metric">
                <div class="val danger-text">${item.billing_rate_pct.toFixed(1)}%</div>
                <span class="kpi-sub">${item.unbilled_services} Unbilled</span>
            </div>
        </div>
    `).join('');

    highIntensityList.innerHTML = globalData.highIntensity.map(item => `
        <div class="alert-item">
            <div class="alert-info">
                <h5>${item.area}</h5>
                <p><i class="fa-solid fa-location-dot"></i> ${item.circle}</p>
                <p>Total Units: ${item.units.toLocaleString()} kWh | Load: ${item.load} kW</p>
            </div>
            <div class="alert-metric">
                <div class="val warning-text">${item.units_per_kw_load.toFixed(1)}</div>
                <span class="kpi-sub">kWh / kW Load</span>
            </div>
        </div>
    `).join('');
}

function renderTable() {
    const tbody = document.getElementById('tableBody');
    const totalRecords = filteredRecords.length;
    const maxPage = Math.ceil(totalRecords / itemsPerPage) || 1;

    if (currentPage > maxPage) currentPage = maxPage;

    const startIdx = (currentPage - 1) * itemsPerPage;
    const endIdx = Math.min(startIdx + itemsPerPage, totalRecords);

    document.getElementById('tableRecordsSummary').innerText = totalRecords > 0 
        ? `Showing ${startIdx + 1} - ${endIdx} of ${totalRecords.toLocaleString()} records`
        : `No matching records found`;

    document.getElementById('pageIndicator').innerText = `Page ${currentPage} of ${maxPage}`;
    document.getElementById('btnPrevPage').disabled = currentPage === 1;
    document.getElementById('btnNextPage').disabled = currentPage === maxPage;

    const pageRecords = filteredRecords.slice(startIdx, endIdx);

    tbody.innerHTML = pageRecords.map(r => `
        <tr>
            <td><strong>${r.circle}</strong></td>
            <td>${r.division}</td>
            <td>${r.section}</td>
            <td>${r.area}</td>
            <td class="text-right">${r.totservices.toLocaleString()}</td>
            <td class="text-right">${r.billdservices.toLocaleString()}</td>
            <td class="text-right">
                <span class="${r.billing_rate_pct < 85 ? 'danger-text' : 'emerald-text'}" style="font-weight:700;">
                    ${r.billing_rate_pct.toFixed(1)}%
                </span>
            </td>
            <td class="text-right">${r.units.toLocaleString()}</td>
            <td class="text-right">${r.load.toLocaleString()}</td>
        </tr>
    `).join('');
}

function exportCSV() {
    if (!filteredRecords.length) return alert("No data available to export.");

    const headers = ['circle', 'division', 'subdivision', 'section', 'area', 'totservices', 'billdservices', 'billing_rate_pct', 'units', 'load'];
    const csvRows = [headers.join(',')];

    filteredRecords.forEach(r => {
        const row = [
            `"${r.circle}"`,
            `"${r.division}"`,
            `"${r.subdivision}"`,
            `"${r.section}"`,
            `"${r.area}"`,
            r.totservices,
            r.billdservices,
            r.billing_rate_pct.toFixed(2),
            r.units,
            r.load
        ];
        csvRows.push(row.join(','));
    });

    const blob = new Blob([csvRows.join('\n')], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `TG_SPDCL_Consumption_Filtered_${Date.now()}.csv`;
    a.click();
    window.URL.revokeObjectURL(url);
}

function openModal(imgSrc, caption) {
    const modal = document.getElementById('imageModal');
    const modalImg = document.getElementById('modalImg');
    const modalCaption = document.getElementById('modalCaption');

    modal.style.display = "flex";
    modalImg.src = imgSrc;
    modalCaption.innerText = caption;
}

function closeModal() {
    document.getElementById('imageModal').style.display = "none";
}

function round(val, decimals = 2) {
    return Number(Math.round(val + 'e' + decimals) + 'e-' + decimals);
}
