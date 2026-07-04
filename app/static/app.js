/**
 * Token Usage Inspector — Frontend Logic
 */
(function () {
    'use strict';

    // --- State ---
    let state = {
        data: null,
        filteredRecords: [],
        sortKey: null,
        sortDir: 'desc',
        charts: {},
        costByResourceView: 'total',
    };

    const $ = (sel) => document.querySelector(sel);
    const $$ = (sel) => document.querySelectorAll(sel);

    // --- DOM Refs ---
    const uploadSection = $('#upload-section');
    const dashboard = $('#dashboard');
    const loading = $('#loading');
    const errorSec = $('#error');
    const errorMsg = $('#error-msg');
    const summaryCards = $('#summary-cards');
    const tableBody = $('#table-body');
    const fileInput = $('#file-input');
    const tableFilter = $('#table-filter');
    const recordCount = $('#record-count');
    const dataFilesList = $('#data-files-list');
    const loadSelectedBtn = $('#load-selected');

    // --- Formatting ---
    const fmt = new Intl.NumberFormat('en-US');
    function fmtNum(n) { return fmt.format(Math.round(n)); }
    function fmtCost(n) { return n.toFixed(2); }
    function fmtPct(part, total) { return total > 0 ? ((part / total) * 100).toFixed(1) + '%' : '0%'; }

    // --- Chart Colors ---
    // Shared Chart.js defaults for a cohesive, refined look
    Chart.defaults.font.family = "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif";
    Chart.defaults.font.size = 11.5;
    Chart.defaults.color = '#64748b';
    Chart.defaults.borderColor = '#eef0f5';
    Chart.defaults.animation.duration = 800;
    Chart.defaults.plugins.legend.labels.usePointStyle = true;
    Chart.defaults.plugins.legend.labels.pointStyle = 'circle';
    Chart.defaults.plugins.legend.labels.boxWidth = 8;
    Chart.defaults.plugins.legend.labels.boxHeight = 8;
    Chart.defaults.plugins.legend.labels.padding = 14;
    Chart.defaults.plugins.tooltip.backgroundColor = 'rgba(15, 23, 42, 0.95)';
    Chart.defaults.plugins.tooltip.titleColor = '#ffffff';
    Chart.defaults.plugins.tooltip.bodyColor = '#e2e8f0';
    Chart.defaults.plugins.tooltip.padding = 10;
    Chart.defaults.plugins.tooltip.cornerRadius = 8;
    Chart.defaults.plugins.tooltip.boxPadding = 6;
    Chart.defaults.plugins.tooltip.titleFont = { weight: '600', size: 12 };
    Chart.defaults.plugins.tooltip.bodyFont = { size: 11.5 };
    Chart.defaults.plugins.tooltip.borderColor = 'rgba(255, 255, 255, 0.08)';
    Chart.defaults.plugins.tooltip.borderWidth = 1;

    const COLORS = {
        input: '#6366f1',
        output: '#10b981',
        cache: '#f59e0b',
        // Distinct, accessible palette for model segments (varied hues, not all purple)
        palette: [
            '#6366f1', '#10b981', '#f59e0b', '#0ea5e9',
            '#ec4899', '#8b5cf6', '#14b8a6', '#f97316',
            '#a855f7', '#84cc16',
        ],
    };
    // Distinguishable palette for per-resource coloring (10 colors, all distinct hues)
    const RESOURCE_PALETTE = [
        '#6366f1', '#10b981', '#f59e0b', '#ef4444',
        '#8b5cf6', '#06b6d4', '#f97316', '#84cc16',
        '#ec4899', '#14b8a6',
    ];
    function getResourceColor(name, resourceNames) {
        // Sort resource names to get a stable index, then assign from palette
        const sorted = [...resourceNames].sort();
        const idx = sorted.indexOf(name);
        return RESOURCE_PALETTE[idx % RESOURCE_PALETTE.length];
    }

    // --- Show/Hide Sections ---
    function show(el) { el.classList.remove('hidden'); }
    function hide(el) { el.classList.add('hidden'); }

    function setLoading(isLoading) {
        if (isLoading) {
            show(loading);
            hide(dashboard);
            hide(uploadSection);
            hide(errorSec);
        } else {
            hide(loading);
        }
    }

    function setError(msg) {
        errorMsg.textContent = msg;
        show(errorSec);
        hide(loading);
        hide(dashboard);
        hide(uploadSection);
    }

    // --- Load Data ---
    async function loadSample() {
        setLoading(true);
        try {
            const res = await fetch('/api/sample');
            if (!res.ok) throw new Error(`Server error: ${res.status}`);
            const data = await res.json();
            if (data.error) throw new Error(data.error);
            state.data = data;
            state.filteredRecords = [...data.records];
            renderAll();
        } catch (e) {
            setError(e.message);
        }
    }

    async function uploadFile(file) {
        setLoading(true);
        try {
            const form = new FormData();
            form.append('file', file);
            const res = await fetch('/api/upload', { method: 'POST', body: form });
            if (!res.ok) throw new Error(`Server error: ${res.status}`);
            const data = await res.json();
            if (data.error) throw new Error(data.error);
            state.data = data;
            state.filteredRecords = [...data.records];
            renderAll();
        } catch (e) {
            setError(e.message);
        }
    }

    async function uploadMultipleFiles(fileList) {
        setLoading(true);
        try {
            const form = new FormData();
            for (const file of fileList) {
                form.append('files', file);
            }
            const res = await fetch('/api/upload-multiple', { method: 'POST', body: form });
            if (!res.ok) throw new Error(`Server error: ${res.status}`);
            const data = await res.json();
            if (data.error) throw new Error(data.error);
            state.data = data;
            state.filteredRecords = [...data.records];
            renderAll();
        } catch (e) {
            setError(e.message);
        }
    }

    async function fetchDataFiles() {
        try {
            const res = await fetch('/api/data-files');
            if (!res.ok) return;
            const data = await res.json();
            renderDataFiles(data.files || []);
        } catch (e) {
            // Silently ignore — data-files is an enhancement
        }
    }

    function renderDataFiles(files) {
        if (!dataFilesList) return;
        if (files.length === 0) {
            dataFilesList.innerHTML = '<p class="data-files-empty">No .xlsx files in data/ folder</p>';
            return;
        }
        dataFilesList.innerHTML = files.map((f) =>
            `<label class="file-item">
                <input type="checkbox" class="data-file-cb" value="${esc(f.name)}">
                <span>${esc(f.name)}</span>
                <span class="file-size">${fmtFileSize(f.size)}</span>
            </label>`
        ).join('');

        // Update load-selected button state on checkbox change
        dataFilesList.querySelectorAll('.data-file-cb').forEach((cb) => {
            cb.addEventListener('change', updateLoadSelectedBtn);
        });
        updateLoadSelectedBtn();
    }

    function updateLoadSelectedBtn() {
        if (!loadSelectedBtn) return;
        const checked = dataFilesList.querySelectorAll('.data-file-cb:checked');
        loadSelectedBtn.disabled = checked.length === 0;
    }

    function fmtFileSize(bytes) {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
        return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
    }

    // --- Load Selected Files ---
    if (loadSelectedBtn) {
        loadSelectedBtn.addEventListener('click', async () => {
            const checked = dataFilesList.querySelectorAll('.data-file-cb:checked');
            if (checked.length === 0) return;
            setLoading(true);
            try {
                const names = [];
                checked.forEach((cb) => names.push(cb.value));
                const res = await fetch('/api/merge-data-files', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ files: names }),
                });
                if (!res.ok) throw new Error(`Server error: ${res.status}`);
                const data = await res.json();
                if (data.error) throw new Error(data.error);
                state.data = data;
                state.filteredRecords = [...data.records];
                renderAll();
            } catch (e) {
                setError(e.message);
            }
        });
    }

    // --- Render All ---
    function renderAll() {
        if (!state.data) return;
        hide(uploadSection);
        hide(loading);
        hide(errorSec);
        show(dashboard);
        renderSummary();
        renderCharts();
        renderPrices();
        renderTable();
    }

    // --- Summary Cards ---
    function renderSummary() {
        const s = state.data.summary;
        const cards = [
            { label: 'Total Tokens', value: fmtNum(s.tokens_total) },
            { label: 'Input Tokens', value: fmtNum(s.tokens_input) },
            { label: 'Output Tokens', value: fmtNum(s.tokens_output) },
            { label: 'Cache Hit Tokens', value: fmtNum(s.tokens_cache_hit) },
            { label: 'Total Cost', value: '¥' + fmtCost(s.cost) },
            { label: 'API Keys', value: s.api_key_count },
            { label: 'Models', value: s.model_count },
            { label: 'Records', value: s.total_records },
        ];
        summaryCards.innerHTML = cards
            .map((c) => `<div class="summary-card"><div class="card-label">${c.label}</div><div class="card-value">${c.value}</div></div>`)
            .join('');
    }

    // --- Charts ---
    function destroyCharts() {
        Object.values(state.charts).forEach((c) => c.destroy());
        state.charts = {};
    }

    // Shared axis styling helpers
    function axisTitle(text) {
        return { display: true, text, color: '#94a3b8', font: { size: 11, weight: '600' }, padding: { top: 8, bottom: 4 } };
    }
    function gridLine() {
        return { color: '#eef0f5', drawBorder: false, drawTicks: false };
    }
    function tickOpts() {
        return { color: '#94a3b8', font: { size: 10.5 }, padding: 6 };
    }

    function renderCharts() {
        destroyCharts();
        const d = state.data;

        // Build resource_id -> resource_name mapping from records
        const idToName = {};
        d.records.forEach((r) => { idToName[r.resource_id] = r.resource_name; });

        // Sorted list of all resource names for consistent coloring
        const allResourceNames = Object.keys(d.by_resource_name);

        // ─── Cost by Resource (toggleable bar chart) ─────────────────────
        populateCostByResourceToggle();
        renderCostByResourceChart();

        // ─── Tokens by Resource (stacked horizontal bar) ──────────────────
        const byResource = Object.entries(d.by_resource_name).sort((a, b) => b[1].cost - a[1].cost);
        state.charts.tokensByResource = new Chart($('#chart-tokens-by-resource'), {
            type: 'bar',
            data: {
                labels: byResource.map(([name]) => name),
                datasets: [
                    {
                        label: 'Input',
                        data: byResource.map(([, v]) => v.tokens_input),
                        backgroundColor: COLORS.input,
                        borderRadius: { topLeft: 0, topRight: 0, bottomLeft: 4, bottomRight: 4 },
                        borderSkipped: false,
                        barPercentage: 0.72,
                        categoryPercentage: 0.8,
                    },
                    {
                        label: 'Output',
                        data: byResource.map(([, v]) => v.tokens_output),
                        backgroundColor: COLORS.output,
                        borderSkipped: false,
                        barPercentage: 0.72,
                        categoryPercentage: 0.8,
                    },
                    {
                        label: 'Cache Hit',
                        data: byResource.map(([, v]) => v.tokens_cache_hit),
                        backgroundColor: COLORS.cache,
                        borderRadius: { topLeft: 4, topRight: 4, bottomLeft: 0, bottomRight: 0 },
                        borderSkipped: false,
                        barPercentage: 0.72,
                        categoryPercentage: 0.8,
                    },
                ],
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                layout: { padding: { top: 4, bottom: 4, left: 0, right: 8 } },
                plugins: {
                    legend: {
                        position: 'bottom',
                        align: 'start',
                        labels: { boxWidth: 8, boxHeight: 8, padding: 14 },
                    },
                    tooltip: {
                        callbacks: {
                            title: (ctx) => byResource[ctx[0].dataIndex][0],
                            label: (ctx) => '  ' + ctx.dataset.label + ': ' + fmtNum(ctx.raw),
                        },
                    },
                },
                scales: {
                    x: {
                        stacked: true,
                        title: axisTitle('Tokens'),
                        ticks: { ...tickOpts(), callback: (v) => fmtNum(v) },
                        grid: gridLine(),
                        border: { display: false },
                    },
                    y: {
                        stacked: true,
                        ticks: { ...tickOpts(), color: '#475569', font: { size: 11.5, weight: '500' } },
                        grid: { display: false },
                        border: { display: false },
                    },
                },
            },
        });

        // ─── Cost Timeline (line chart, per-resource datasets) ────────────
        const dates = Object.keys(d.timeline).sort();
        // Collect per-resource per-date data
        const resourceTimelineCost = {};
        const resourceTimelineTokens = {};
        allResourceNames.forEach((name) => {
            resourceTimelineCost[name] = [];
            resourceTimelineTokens[name] = [];
        });
        dates.forEach((date) => {
            const dayEntry = d.timeline[date];
            // Initialize all resources to 0 for this date
            const dayCosts = {};
            const dayTokens = {};
            allResourceNames.forEach((name) => { dayCosts[name] = 0; dayTokens[name] = 0; });
            // Add data from by_key entries, mapped to resource_name
            Object.entries(dayEntry.by_key).forEach(([keyId, keyData]) => {
                const rName = idToName[keyId];
                if (rName) {
                    dayCosts[rName] = (dayCosts[rName] || 0) + keyData.cost;
                    dayTokens[rName] = (dayTokens[rName] || 0) + keyData.tokens_total;
                }
            });
            allResourceNames.forEach((name) => {
                resourceTimelineCost[name].push(dayCosts[name]);
                resourceTimelineTokens[name].push(dayTokens[name]);
            });
        });

        const timelineDatasets = allResourceNames.map((name) => ({
            label: name,
            data: resourceTimelineCost[name],
            borderColor: getResourceColor(name, allResourceNames),
            backgroundColor: getResourceColor(name, allResourceNames) + '20',
            tension: 0.35,
            fill: false,
            borderWidth: 2,
            pointRadius: 0,
            pointHoverRadius: 5,
            pointHoverBorderWidth: 2,
            pointHoverBackgroundColor: '#ffffff',
        }));
        state.charts.timelineCost = new Chart($('#chart-timeline-cost'), {
            type: 'line',
            data: {
                labels: dates,
                datasets: timelineDatasets,
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                layout: { padding: { top: 4, bottom: 4, left: 0, right: 8 } },
                plugins: {
                    legend: {
                        display: true,
                        position: 'bottom',
                        align: 'start',
                        labels: { boxWidth: 8, boxHeight: 8, padding: 14 },
                    },
                    tooltip: {
                        callbacks: {
                            label: (ctx) => '  ' + ctx.dataset.label + ': ¥' + fmtCost(ctx.raw),
                        },
                    },
                },
                scales: {
                    x: {
                        title: axisTitle('Date'),
                        ticks: tickOpts(),
                        grid: { display: false },
                        border: { display: false },
                    },
                    y: {
                        title: axisTitle('Cost (¥)'),
                        ticks: { ...tickOpts(), callback: (v) => '¥' + fmtCost(v) },
                        grid: gridLine(),
                        border: { display: false },
                    },
                },
            },
        });

        // ─── Tokens Timeline (line chart, per-resource datasets) ──────────
        const tokenDatasets = allResourceNames.map((name) => ({
            label: name,
            data: resourceTimelineTokens[name],
            borderColor: getResourceColor(name, allResourceNames),
            backgroundColor: getResourceColor(name, allResourceNames) + '20',
            tension: 0.35,
            fill: false,
            borderWidth: 2,
            pointRadius: 0,
            pointHoverRadius: 5,
            pointHoverBorderWidth: 2,
            pointHoverBackgroundColor: '#ffffff',
        }));
        state.charts.timelineTokens = new Chart($('#chart-timeline-tokens'), {
            type: 'line',
            data: {
                labels: dates,
                datasets: tokenDatasets,
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                layout: { padding: { top: 4, bottom: 4, left: 0, right: 8 } },
                plugins: {
                    legend: {
                        display: true,
                        position: 'bottom',
                        align: 'start',
                        labels: { boxWidth: 8, boxHeight: 8, padding: 14 },
                    },
                    tooltip: {
                        callbacks: {
                            label: (ctx) => '  ' + ctx.dataset.label + ': ' + fmtNum(ctx.raw),
                        },
                    },
                },
                scales: {
                    x: {
                        title: axisTitle('Date'),
                        ticks: tickOpts(),
                        grid: { display: false },
                        border: { display: false },
                    },
                    y: {
                        title: axisTitle('Tokens'),
                        ticks: { ...tickOpts(), callback: (v) => fmtNum(v) },
                        grid: gridLine(),
                        border: { display: false },
                    },
                },
            },
        });

        // ─── Cost by Model (doughnut) ─────────────────────────────────────
        const byModel = Object.entries(d.by_model).sort((a, b) => b[1].cost - a[1].cost);
        state.charts.costByModel = new Chart($('#chart-cost-by-model'), {
            type: 'doughnut',
            data: {
                labels: byModel.map(([m]) => m),
                datasets: [{
                    data: byModel.map(([, v]) => v.cost),
                    backgroundColor: COLORS.palette.slice(0, byModel.length),
                    borderColor: '#ffffff',
                    borderWidth: 3,
                    hoverOffset: 8,
                    spacing: 2,
                }],
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                cutout: '62%',
                layout: { padding: 8 },
                plugins: {
                    legend: {
                        position: 'right',
                        align: 'center',
                        labels: { boxWidth: 8, boxHeight: 8, padding: 10 },
                    },
                    tooltip: {
                        callbacks: {
                            label: (ctx) => '  ' + ctx.label + ': ¥' + fmtCost(ctx.raw),
                        },
                    },
                },
            },
        });

        // ─── Tokens by Model (stacked horizontal bar) ─────────────────────
        state.charts.tokensByModel = new Chart($('#chart-tokens-by-model'), {
            type: 'bar',
            data: {
                labels: byModel.map(([m]) => m),
                datasets: [
                    {
                        label: 'Input',
                        data: byModel.map(([, v]) => v.tokens_input),
                        backgroundColor: COLORS.input,
                        borderRadius: { topLeft: 0, topRight: 0, bottomLeft: 4, bottomRight: 4 },
                        borderSkipped: false,
                        barPercentage: 0.72,
                        categoryPercentage: 0.8,
                    },
                    {
                        label: 'Output',
                        data: byModel.map(([, v]) => v.tokens_output),
                        backgroundColor: COLORS.output,
                        borderSkipped: false,
                        barPercentage: 0.72,
                        categoryPercentage: 0.8,
                    },
                    {
                        label: 'Cache Hit',
                        data: byModel.map(([, v]) => v.tokens_cache_hit),
                        backgroundColor: COLORS.cache,
                        borderRadius: { topLeft: 4, topRight: 4, bottomLeft: 0, bottomRight: 0 },
                        borderSkipped: false,
                        barPercentage: 0.72,
                        categoryPercentage: 0.8,
                    },
                ],
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                layout: { padding: { top: 4, bottom: 4, left: 0, right: 8 } },
                plugins: {
                    legend: {
                        position: 'bottom',
                        align: 'start',
                        labels: { boxWidth: 8, boxHeight: 8, padding: 14 },
                    },
                    tooltip: {
                        callbacks: {
                            label: (ctx) => '  ' + ctx.dataset.label + ': ' + fmtNum(ctx.raw),
                        },
                    },
                },
                scales: {
                    x: {
                        stacked: true,
                        title: axisTitle('Tokens'),
                        ticks: { ...tickOpts(), callback: (v) => fmtNum(v) },
                        grid: gridLine(),
                        border: { display: false },
                    },
                    y: {
                        stacked: true,
                        ticks: { ...tickOpts(), color: '#475569', font: { size: 11.5, weight: '500' } },
                        grid: { display: false },
                        border: { display: false },
                    },
                },
            },
        });

        // ─── Model Breakdown by Resource ──────────────────────────────────
        renderModelBreakdown();
    }

    // ─── Cost by Resource Chart (toggleable views) ────────────────────────
    function renderCostByResourceChart() {
        if (state.charts.costByResource) {
            state.charts.costByResource.destroy();
            delete state.charts.costByResource;
        }

        const d = state.data;
        const view = state.costByResourceView;
        const byResource = Object.entries(d.by_resource_name).sort((a, b) => b[1].cost - a[1].cost);
        const allResourceNames = Object.keys(d.by_resource_name);

        let datasets;
        if (view === 'total') {
            datasets = [{
                label: 'Cost (¥)',
                data: byResource.map(([, v]) => v.cost),
                backgroundColor: byResource.map(([name]) => getResourceColor(name, allResourceNames)),
                borderRadius: 4,
                borderSkipped: false,
                barPercentage: 0.72,
                categoryPercentage: 0.8,
            }];
        } else if (view === 'by_model') {
            // Compute per-resource per-model cost from records
            const resModelCost = {};
            d.records.forEach((r) => {
                const name = r.resource_name;
                if (!resModelCost[name]) resModelCost[name] = {};
                resModelCost[name][r.model] = (resModelCost[name][r.model] || 0) + r.cost;
            });
            const allModels = [...new Set(d.records.map((r) => r.model))].sort();
            datasets = allModels.map((model, i) => ({
                label: model,
                data: byResource.map(([name]) => (resModelCost[name] && resModelCost[name][model]) || 0),
                backgroundColor: COLORS.palette[i % COLORS.palette.length],
                borderSkipped: false,
                barPercentage: 0.72,
                categoryPercentage: 0.8,
            }));
            if (datasets.length > 0) {
                datasets[0].borderRadius = { topLeft: 0, topRight: 0, bottomLeft: 4, bottomRight: 4 };
                datasets[datasets.length - 1].borderRadius = { topLeft: 4, topRight: 4, bottomLeft: 0, bottomRight: 0 };
            }
        } else if (view === 'by_token_type') {
            // Compute estimated cost per token type per resource from records using prices
            const prices = d.prices || {};
            const resEstCost = { input: {}, output: {}, cache: {} };
            d.records.forEach((r) => {
                const name = r.resource_name;
                const p = prices[r.model] || {};
                resEstCost.input[name] = (resEstCost.input[name] || 0) + r.tokens_input * (p.input || 0);
                resEstCost.output[name] = (resEstCost.output[name] || 0) + r.tokens_output * (p.output || 0);
                resEstCost.cache[name] = (resEstCost.cache[name] || 0) + r.tokens_cache_hit * (p.cache_hit || 0);
            });
            datasets = [
                {
                    label: 'Input Cost',
                    data: byResource.map(([name]) => resEstCost.input[name] || 0),
                    backgroundColor: COLORS.input,
                    borderRadius: { topLeft: 0, topRight: 0, bottomLeft: 4, bottomRight: 4 },
                    borderSkipped: false,
                    barPercentage: 0.72,
                    categoryPercentage: 0.8,
                },
                {
                    label: 'Output Cost',
                    data: byResource.map(([name]) => resEstCost.output[name] || 0),
                    backgroundColor: COLORS.output,
                    borderSkipped: false,
                    barPercentage: 0.72,
                    categoryPercentage: 0.8,
                },
                {
                    label: 'Cache Hit Cost',
                    data: byResource.map(([name]) => resEstCost.cache[name] || 0),
                    backgroundColor: COLORS.cache,
                    borderRadius: { topLeft: 4, topRight: 4, bottomLeft: 0, bottomRight: 0 },
                    borderSkipped: false,
                    barPercentage: 0.72,
                    categoryPercentage: 0.8,
                },
            ];
        }

        state.charts.costByResource = new Chart($('#chart-cost-by-resource'), {
            type: 'bar',
            data: { labels: byResource.map(([name]) => name), datasets },
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                layout: { padding: { top: 4, bottom: 4, left: 0, right: 8 } },
                plugins: {
                    legend: view === 'total' ? { display: false } : { position: 'bottom', align: 'start', labels: { boxWidth: 8, boxHeight: 8, padding: 14 } },
                    tooltip: {
                        callbacks: {
                            title: (ctx) => byResource[ctx[0].dataIndex][0],
                            label: (ctx) => '  ' + (view === 'total' ? '¥' + fmtCost(ctx.raw) : ctx.dataset.label + ': ¥' + fmtCost(ctx.raw)),
                        },
                    },
                },
                scales: {
                    x: {
                        stacked: view !== 'total',
                        title: axisTitle('Cost (¥)'),
                        ticks: { ...tickOpts(), callback: (v) => '¥' + fmtCost(v) },
                        grid: gridLine(),
                        border: { display: false },
                    },
                    y: {
                        stacked: view !== 'total',
                        ticks: { ...tickOpts(), color: '#475569', font: { size: 11.5, weight: '500' } },
                        grid: { display: false },
                        border: { display: false },
                    },
                },
            },
        });
    }

    function populateCostByResourceToggle() {
        const container = $('#toggle-cost-by-resource');
        if (!container) return;
        const views = [
            { key: 'total', label: 'Total' },
            { key: 'by_model', label: 'By Model' },
            { key: 'by_token_type', label: 'By Token Type' },
        ];
        container.innerHTML = views.map((v) =>
            `<button class="btn-sm${state.costByResourceView === v.key ? ' active' : ''}" data-view="${v.key}">${v.label}</button>`
        ).join('');
        container.querySelectorAll('.btn-sm').forEach((btn) => {
            btn.addEventListener('click', () => {
                state.costByResourceView = btn.dataset.view;
                container.querySelectorAll('.btn-sm').forEach((b) => b.classList.remove('active'));
                btn.classList.add('active');
                renderCostByResourceChart();
            });
        });
    }

    // ─── Model Breakdown by Resource ───────────────────────────────────────
    function renderModelBreakdown() {
        const d = state.data;
        const container = $('#model-per-resource');
        const allResourceNames = Object.keys(d.by_resource_name).sort();

        // Compute per-resource per-model token type data from records
        const resModelData = {};  // { resource_name: { model: { tokens_input, tokens_output, tokens_cache_hit, cost } } }
        allResourceNames.forEach((name) => { resModelData[name] = {}; });
        d.records.forEach((r) => {
            const name = r.resource_name;
            if (!resModelData[name]) resModelData[name] = {};
            if (!resModelData[name][r.model]) {
                resModelData[name][r.model] = { tokens_input: 0, tokens_output: 0, tokens_cache_hit: 0, cost: 0 };
            }
            resModelData[name][r.model].tokens_input += r.tokens_input;
            resModelData[name][r.model].tokens_output += r.tokens_output;
            resModelData[name][r.model].tokens_cache_hit += r.tokens_cache_hit;
            resModelData[name][r.model].cost += r.cost;
        });

        const prices = d.prices || {};
        let html = '';
        const breakdownCharts = [];

        allResourceNames.forEach((name) => {
            const models = Object.entries(resModelData[name]).sort((a, b) => b[1].cost - a[1].cost);
            const safeId = 'model-breakdown-' + name.replace(/[^a-zA-Z0-9_-]/g, '_');
            const tokenCanvasId = 'chart-tokens-' + safeId;
            const costCanvasId = 'chart-cost-' + safeId;
            html += `<div class="resource-block">
                <details open>
                    <summary><h3>${esc(name)}</h3></summary>
                    <details open>
                        <summary>Tokens by Model</summary>
                        <canvas id="${tokenCanvasId}"></canvas>
                    </details>
                    <details open>
                        <summary>Estimated Cost by Model</summary>
                        <canvas id="${costCanvasId}"></canvas>
                    </details>
                </details>
            </div>`;
            breakdownCharts.push({ name, models, tokenCanvasId, costCanvasId });
        });

        container.innerHTML = html;

        // Render mini bar charts — defer to next frame so DOM is ready
        requestAnimationFrame(() => {
            breakdownCharts.forEach(({ name, models, tokenCanvasId, costCanvasId }) => {
                const tokenEl = document.getElementById(tokenCanvasId);
                const costEl = document.getElementById(costCanvasId);

                // Token chart (stacked by type)
                if (tokenEl) {
                    const chartKey = 'modelTokens_' + tokenCanvasId;
                    if (state.charts[chartKey]) {
                        state.charts[chartKey].destroy();
                        delete state.charts[chartKey];
                    }
                    state.charts[chartKey] = new Chart(tokenEl, {
                        type: 'bar',
                        data: {
                            labels: models.map(([m]) => m),
                            datasets: [
                                {
                                    label: 'Input',
                                    data: models.map(([, v]) => v.tokens_input),
                                    backgroundColor: COLORS.input,
                                    borderRadius: { topLeft: 0, topRight: 0, bottomLeft: 4, bottomRight: 4 },
                                    borderSkipped: false,
                                    barPercentage: 0.72,
                                    categoryPercentage: 0.85,
                                },
                                {
                                    label: 'Output',
                                    data: models.map(([, v]) => v.tokens_output),
                                    backgroundColor: COLORS.output,
                                    borderSkipped: false,
                                    barPercentage: 0.72,
                                    categoryPercentage: 0.85,
                                },
                                {
                                    label: 'Cache Hit',
                                    data: models.map(([, v]) => v.tokens_cache_hit),
                                    backgroundColor: COLORS.cache,
                                    borderRadius: { topLeft: 4, topRight: 4, bottomLeft: 0, bottomRight: 0 },
                                    borderSkipped: false,
                                    barPercentage: 0.72,
                                    categoryPercentage: 0.85,
                                },
                            ],
                        },
                        options: {
                            indexAxis: 'y',
                            responsive: true,
                            maintainAspectRatio: false,
                            layout: { padding: { top: 4, bottom: 4, left: 0, right: 8 } },
                            plugins: {
                                legend: { display: false },
                                tooltip: {
                                    callbacks: {
                                        label: (ctx) => '  ' + ctx.dataset.label + ': ' + fmtNum(ctx.raw),
                                    },
                                },
                            },
                            scales: {
                                x: {
                                    stacked: true,
                                    title: axisTitle('Tokens'),
                                    ticks: { ...tickOpts(), callback: (v) => fmtNum(v) },
                                    grid: gridLine(),
                                    border: { display: false },
                                },
                                y: {
                                    stacked: true,
                                    ticks: { ...tickOpts(), color: '#475569', font: { size: 11, weight: '500' } },
                                    grid: { display: false },
                                    border: { display: false },
                                },
                            },
                        },
                    });
                }

                // Cost chart (stacked by token type)
                if (costEl) {
                    const chartKey = 'modelCost_' + costCanvasId;
                    if (state.charts[chartKey]) {
                        state.charts[chartKey].destroy();
                        delete state.charts[chartKey];
                    }
                    state.charts[chartKey] = new Chart(costEl, {
                        type: 'bar',
                        data: {
                            labels: models.map(([m]) => m),
                            datasets: [
                                {
                                    label: 'Input',
                                    data: models.map(([modelName, v]) => {
                                        const p = prices[modelName] || {};
                                        return v.tokens_input * (p.input || 0);
                                    }),
                                    backgroundColor: COLORS.input,
                                    borderRadius: { topLeft: 0, topRight: 0, bottomLeft: 4, bottomRight: 4 },
                                    borderSkipped: false,
                                    barPercentage: 0.72,
                                    categoryPercentage: 0.85,
                                },
                                {
                                    label: 'Output',
                                    data: models.map(([modelName, v]) => {
                                        const p = prices[modelName] || {};
                                        return v.tokens_output * (p.output || 0);
                                    }),
                                    backgroundColor: COLORS.output,
                                    borderSkipped: false,
                                    barPercentage: 0.72,
                                    categoryPercentage: 0.85,
                                },
                                {
                                    label: 'Cache Hit',
                                    data: models.map(([modelName, v]) => {
                                        const p = prices[modelName] || {};
                                        return v.tokens_cache_hit * (p.cache_hit || 0);
                                    }),
                                    backgroundColor: COLORS.cache,
                                    borderRadius: { topLeft: 4, topRight: 4, bottomLeft: 0, bottomRight: 0 },
                                    borderSkipped: false,
                                    barPercentage: 0.72,
                                    categoryPercentage: 0.85,
                                },
                            ],
                        },
                        options: {
                            indexAxis: 'y',
                            responsive: true,
                            maintainAspectRatio: false,
                            layout: { padding: { top: 4, bottom: 4, left: 0, right: 8 } },
                            plugins: {
                                legend: { display: false },
                                tooltip: {
                                    callbacks: {
                                        label: (ctx) => '  ' + ctx.dataset.label + ': ¥' + fmtCost(ctx.raw),
                                    },
                                },
                            },
                            scales: {
                                x: {
                                    stacked: true,
                                    title: axisTitle('Cost (¥)'),
                                    ticks: { ...tickOpts(), callback: (v) => '¥' + fmtCost(v) },
                                    grid: gridLine(),
                                    border: { display: false },
                                },
                                y: {
                                    stacked: true,
                                    ticks: { ...tickOpts(), color: '#475569', font: { size: 11, weight: '500' } },
                                    grid: { display: false },
                                    border: { display: false },
                                },
                            },
                        },
                    });
                }
            });
            // Resize charts when a details panel opens/closes
            container.querySelectorAll('details').forEach((d) => {
                d.addEventListener('toggle', () => {
                    Object.values(state.charts).forEach((chart) => {
                        if (chart.canvas && d.contains(chart.canvas)) {
                            chart.resize();
                        }
                    });
                });
            });
        });
    }

    // ─── Estimated Prices by Model ──────────────────────────────────────
    function renderPrices() {
        const prices = state.data.prices;
        if (!prices) return;
        const container = $('#prices-container');
        const models = Object.keys(prices).sort();
        const hasCache = models.some((m) => prices[m].cache_hit !== null);

        let html = '<table class="prices-table"><thead><tr><th>Model</th><th>Input (¥/1M tokens)</th><th>Output (¥/1M tokens)</th>';
        if (hasCache) html += '<th>Cache Hit (¥/1M tokens)</th>';
        html += '</tr></thead><tbody>';

        models.forEach((model) => {
            const p = prices[model];
            const inPrice = p.input !== null ? '¥' + (p.input * 1000000).toFixed(4) : '<span class="price-na">—</span>';
            const outPrice = p.output !== null ? '¥' + (p.output * 1000000).toFixed(4) : '<span class="price-na">—</span>';
            const cachePrice = p.cache_hit !== null ? '¥' + (p.cache_hit * 1000000).toFixed(4) : '<span class="price-na">—</span>';
            html += '<tr><td>' + esc(model) + '</td><td>' + inPrice + '</td><td>' + outPrice + '</td>';
            if (hasCache) html += '<td>' + cachePrice + '</td>';
            html += '</tr>';
        });

        html += '</tbody></table>';
        container.innerHTML = html;
    }

    // --- Table ---
    function renderTable() {
        const records = state.filteredRecords;
        recordCount.textContent = records.length + ' record' + (records.length !== 1 ? 's' : '');

        tableBody.innerHTML = records
            .map((r) => {
                const kid = r.resource_id.slice(-12);
                return `<tr>
                    <td>${esc(r.date)}</td>
                    <td>${esc(r.resource_name)}</td>
                    <td><span class="api-key-cell" title="${esc(r.resource_id)}">${esc(kid)}</span></td>
                    <td>${esc(r.model)}</td>
                    <td>${fmtNum(r.tokens_input)}</td>
                    <td>${fmtNum(r.tokens_output)}</td>
                    <td>${fmtNum(r.tokens_cache_hit)}</td>
                    <td>${fmtNum(r.tokens_total)}</td>
                    <td class="cost-cell">¥${fmtCost(r.cost)}</td>
                </tr>`;
            })
            .join('');

        if (records.length === 0 && state.data) {
            tableBody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:24px;color:var(--text-secondary)">No matching records</td></tr>';
        }
    }

    function esc(s) {
        if (!s) return '';
        const d = document.createElement('div');
        d.textContent = s;
        return d.innerHTML;
    }

    // --- Filter ---
    tableFilter.addEventListener('input', () => {
        const q = tableFilter.value.toLowerCase();
        state.filteredRecords = state.data.records.filter((r) => {
            return (
                r.date.toLowerCase().includes(q) ||
                r.resource_name.toLowerCase().includes(q) ||
                (r.resource_id || '').toLowerCase().includes(q) ||
                r.model.toLowerCase().includes(q) ||
                r.usage_desc.toLowerCase().includes(q)
            );
        });
        renderTable();
    });

    // --- Sort ---
    document.querySelector('#data-table thead').addEventListener('click', (e) => {
        const th = e.target.closest('th');
        if (!th) return;
        const key = th.dataset.sort;
        if (!key) return;

        if (state.sortKey === key) {
            state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
        } else {
            state.sortKey = key;
            state.sortDir = 'desc';
        }

        state.filteredRecords.sort((a, b) => {
            let va = a[key], vb = b[key];
            if (typeof va === 'string') va = va.toLowerCase();
            if (typeof vb === 'string') vb = vb.toLowerCase();
            if (va < vb) return state.sortDir === 'asc' ? -1 : 1;
            if (va > vb) return state.sortDir === 'asc' ? 1 : -1;
            return 0;
        });

        // Update header classes
        $$('#data-table th').forEach((h) => h.classList.remove('sorted', 'asc', 'desc'));
        th.classList.add('sorted', state.sortDir);

        renderTable();
    });

    // --- Upload Events ---
    const uploadZone = $('#upload-zone');

    uploadZone.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', () => {
        if (fileInput.files.length) {
            if (fileInput.files.length === 1) {
                uploadFile(fileInput.files[0]);
            } else {
                uploadMultipleFiles(fileInput.files);
            }
        }
    });

    uploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadZone.classList.add('drag-over');
    });
    uploadZone.addEventListener('dragleave', () => {
        uploadZone.classList.remove('drag-over');
    });
    uploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadZone.classList.remove('drag-over');
        if (e.dataTransfer.files.length) {
            if (e.dataTransfer.files.length === 1) {
                uploadFile(e.dataTransfer.files[0]);
            } else {
                uploadMultipleFiles(e.dataTransfer.files);
            }
        }
    });

    // --- Load Sample ---
    $('#load-sample').addEventListener('click', (e) => {
        e.stopPropagation();
        loadSample();
    });

    // --- Init: fetch data file list, do NOT auto-load sample ---
    fetchDataFiles();
})();
