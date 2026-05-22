/* ============================================
   MESSSTELLEN MANAGER PRO - v6.0
   Professional Field Measurement Tool
   Bug-free, Performance-optimized
   ============================================ */

'use strict';

// ─── DATA STRUCTURE ───
const DEPTHS = ['0.8', '1.6', '3.2'];

const TABLE_STRUCTURE = [
    { group: 'Basis', class: 'basis', columns: ['Kennzeichen', 'Alt-Kz.', 'Typ', 'Örtlichkeit', 'Meter [m]', 'Datum'] },
    { group: '0.8m', class: '08', columns: ['R1 [Ω]_0.8', 'R2 [Ω]_0.8', 'R3 [Ω]_0.8', 'ρ1 [Ωm]_0.8', 'ρ2 [Ωm]_0.8', 'ρ3 [Ωm]_0.8', 'MW [Ωm]_0.8', 'SD [Ωm]_0.8', 'Anhang_0.8'] },
    { group: '1.6m', class: '16', columns: ['R1 [Ω]_1.6', 'R2 [Ω]_1.6', 'R3 [Ω]_1.6', 'ρ1 [Ωm]_1.6', 'ρ2 [Ωm]_1.6', 'ρ3 [Ωm]_1.6', 'MW [Ωm]_1.6', 'SD [Ωm]_1.6', 'Anhang_1.6'] },
    { group: '3.2m', class: '32', columns: ['R1 [Ω]_3.2', 'R2 [Ω]_3.2', 'R3 [Ω]_3.2', 'ρ1 [Ωm]_3.2', 'ρ2 [Ωm]_3.2', 'ρ3 [Ωm]_3.2', 'MW [Ωm]_3.2', 'SD [Ωm]_3.2', 'Anhang_3.2'] },
    { group: 'Anhang (Gesamt)', class: 'anhang_global', columns: ['Anhang_Global'] },
    { group: 'Potential', class: 'potential', columns: ['Pot. Ein [V]', 'Pot. Aus [V]', 'AC [V]'] },
    { group: 'Spannung', class: 'spannung', columns: ['Spannung Ein', 'Spannung Aus'] },
    { group: 'Strom', class: 'strom', columns: ['Strom Ein', 'Strom Aus', 'Strom Diff.', 'Mikro Diff.'] },
    { group: 'Widerstand', class: 'widerstand', columns: ['R [Ω]', 'Bodenwid. [Ωm]', '+/-', 'Ra', 'Kommentar'] },
    { group: 'Audit', class: 'audit', columns: ['Erderinformationen', 'Existiert die Messstelle', 'Typ korrekt', 'Fernüberwacht'] },
    { group: 'GPS-Daten', class: 'special', columns: ['GPS-Lat', 'GPS-Lng'] }
];

const DEPTH_COLORS = {
    'basis': { bg: 'transparent', border: '#334155', text: '#94a3b8' },
    '08': { bg: 'rgba(232,121,249,0.06)', border: '#e879f9', text: '#e879f9' },
    '16': { bg: 'rgba(251,191,36,0.06)', border: '#fbbf24', text: '#fbbf24' },
    '32': { bg: 'rgba(34,211,238,0.06)', border: '#22d3ee', text: '#22d3ee' },
    'zusatz': { bg: 'rgba(57,255,20,0.06)', border: '#39ff14', text: '#39ff14' },
    'potential': { bg: 'rgba(255,140,0,0.06)', border: '#ff8c00', text: '#ff8c00' },
    'spannung': { bg: 'rgba(255,0,60,0.06)', border: '#ff003c', text: '#ff003c' },
    'strom': { bg: 'rgba(59,130,246,0.06)', border: '#3b82f6', text: '#3b82f6' },
    'widerstand': { bg: 'rgba(168,85,247,0.06)', border: '#a855f7', text: '#a855f7' },
    'audit': { bg: 'transparent', border: '#64748b', text: '#64748b' },
    'special': { bg: 'rgba(59,130,246,0.04)', border: '#3b82f6', text: '#3b82f6' },
    'anhang_global': { bg: 'transparent', border: '#64748b', text: '#94a3b8' }
};

// ─── APPLICATION STATE ───
const AppState = {
    data: [],
    hiddenColumns: new Set(['32', 'potential', 'spannung', 'strom', 'widerstand', 'audit', 'special']),
    zoomLevel: 100,
    selectedCell: null,
    columnWidths: {},
    activeColor: '#e879f9',
    liveFollow: false,
    newCols: new Set(),
    filters: { 'Kennzeichen': '' },
    userMarker: null,
    drawItems: null,
    activeDrawTool: null,
    hiddenMapColors: new Set(['#22d3ee']),
    undoStack: [],
    redoStack: [],
    maxUndo: 30,
    chartMode: 'bar',
    depthMarkers: { '0.8': [], '1.6': [], '3.2': [] }
};

const STORAGE_KEY = 'messstellen_v6_pro';
let map = null;
let layers = {};
let captureMode = false;
let snipTexts = [];
let draggingTextIdx = -1;

// ─── UTILITY FUNCTIONS ───
const $ = (id) => document.getElementById(id);
const on = (id, evt, cb) => { const el = $(id); if (el) el.addEventListener(evt, cb); };

function showToast(message, duration = 3000) {
    const existing = document.querySelectorAll('.modern-toast');
    existing.forEach(t => { t.style.opacity = '0'; });

    const toast = document.createElement('div');
    toast.className = 'modern-toast';
    toast.innerHTML = `<i class="fas fa-info-circle"></i> <span>${message}</span>`;
    document.body.appendChild(toast);

    requestAnimationFrame(() => toast.classList.add('show'));

    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => { if (toast.parentNode) toast.parentNode.removeChild(toast); }, 400);
    }, duration);
}

function pushUndo() {
    AppState.undoStack.push(JSON.parse(JSON.stringify(AppState.data)));
    if (AppState.undoStack.length > AppState.maxUndo) AppState.undoStack.shift();
    AppState.redoStack = [];
}

function undo() {
    if (AppState.undoStack.length === 0) { showToast('Nichts zum Rückgängig machen'); return; }
    AppState.redoStack.push(JSON.parse(JSON.stringify(AppState.data)));
    AppState.data = AppState.undoStack.pop();
    renderTable();
    showToast('Rückgängig');
}

function redo() {
    if (AppState.redoStack.length === 0) { showToast('Nichts zum Wiederholen'); return; }
    AppState.undoStack.push(JSON.parse(JSON.stringify(AppState.data)));
    AppState.data = AppState.redoStack.pop();
    renderTable();
    showToast('Wiederholt');
}

// ─── ZOOM ───
window.handleZoom = function(delta) {
    AppState.zoomLevel = Math.max(50, Math.min(200, AppState.zoomLevel + delta));
    const val = AppState.zoomLevel;
    const display = $('zoomLevelDisplay');
    if (display) display.textContent = val + '%';

    const wrapper = document.querySelector('.table-wrapper');
    if (wrapper) {
        wrapper.style.zoom = val / 100;
    }
};

// ─── TAB SWITCHING ───
function switchTab(tabName) {
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));

    if (tabName === 'map') {
        openMap();
        return;
    }

    const panel = $('tab' + tabName.charAt(0).toUpperCase() + tabName.slice(1));
    if (panel) panel.classList.add('active');

    const btn = $('btnTab' + tabName.charAt(0).toUpperCase() + tabName.slice(1));
    if (btn) btn.classList.add('active');

    if (tabName === 'plot') renderAppPlot();
}

// ─── INITIALIZATION ───
document.addEventListener('DOMContentLoaded', () => {
    // Unregister old service workers to prevent caching issues
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.getRegistrations().then(function(regs) {
            regs.forEach(function(r) { r.unregister(); });
        });
    }

    initUI();
    initConnectionCheck();
});

function initConnectionCheck() {
    window.addEventListener('online', () => showToast('Online'));
    window.addEventListener('offline', () => showToast('Offline-Modus aktiv'));
}

function initUI() {
    // Welcome screen
    on('btnWelcomeBrowse', 'click', () => $('fileInput').click());
    on('btnWelcomeBlank', 'click', () => {
        AppState.data = [];
        createEmptyRows(50);
        $('fileStatus').textContent = 'Neues Projekt erstellt';
        $('btnFinalStart').disabled = false;
    });
    on('btnFinalStart', 'click', () => {
        $('welcomeOverlay').style.display = 'none';
        $('mainApp').style.display = 'flex';
        renderTable();
        showToast('Projekt gestartet');
    });

    // Header buttons
    on('btnHome', 'click', () => {
        $('welcomeOverlay').style.display = 'flex';
        $('mainApp').style.display = 'none';
        $('fileStatus').textContent = 'Keine Datei ausgewählt';
        $('btnFinalStart').disabled = true;
    });
    on('btnBrowse', 'click', () => $('fileInput').click());
    on('btnStartBlank', 'click', () => {
        closeMap();
        if (AppState.data.length === 0) createEmptyRows(50);
        renderTable();
    });
    on('btnUndo', 'click', undo);
    on('btnRedo', 'click', redo);
    on('btnSaveAll', 'click', () => { saveToStorage(); showToast('Gespeichert'); });
    on('btnExportExcel', 'click', openExportModal);
    on('btnOpenProjects', 'click', () => { renderProjectList(); $('standortModal').style.display = 'flex'; });
    on('btnCloseStandortModal', 'click', () => { $('standortModal').style.display = 'none'; });

    // File input
    on('fileInput', 'change', (e) => { if (e.target.files[0]) handleFileUpload(e.target.files[0]); });

    // Tabs
    on('btnTabTable', 'click', () => switchTab('table'));
    on('btnTabPlot', 'click', () => switchTab('plot'));
    on('btnToggleMap', 'click', () => switchTab('map'));

    // Toolbar
    on('btnColManager', 'click', openColManager);
    on('btnCloseColManager', 'click', () => { $('colManagerModal').style.display = 'none'; });
    on('btnRefreshTable', 'click', () => { renderTable(); showToast('Aktualisiert'); });

    // Row/Column management
    on('btnAddRow', 'click', () => {
        pushUndo();
        if (AppState.selectedCell) {
            const rowIndex = parseInt(AppState.selectedCell.split('-')[0]);
            AppState.data.splice(rowIndex + 1, 0, { _isNew: true });
        } else {
            AppState.data.push({ _isNew: true });
        }
        renderTable();
        saveToStorage();
        showToast('Zeile hinzugefügt');
    });

    on('btnDeleteRow', 'click', () => {
        if (!AppState.selectedCell) { showToast('Bitte zuerst eine Zeile auswählen'); return; }
        pushUndo();
        AppState.data.splice(parseInt(AppState.selectedCell.split('-')[0]), 1);
        renderTable();
        saveToStorage();
        showToast('Zeile gelöscht');
    });

    on('btnAddCol', 'click', () => {
        const name = prompt('Name der neuen Spalte:');
        if (!name || !name.trim()) return;
        const colName = name.trim();

        // Check if column already exists anywhere
        for (let g of TABLE_STRUCTURE) {
            if (g.columns.includes(colName)) { showToast('Spalte existiert bereits'); return; }
        }

        let targetGroup = null;
        let insertIdx = -1;

        // If a cell is selected, insert next to that column
        if (AppState.selectedCell && typeof AppState.selectedCell === 'string') {
            const selectedCol = AppState.selectedCell.split('-')[1];
            for (let g of TABLE_STRUCTURE) {
                const idx = g.columns.indexOf(selectedCol);
                if (idx !== -1) {
                    targetGroup = g;
                    insertIdx = idx + 1;
                    break;
                }
            }
        }

        // Fallback: add to Zusatz group
        if (!targetGroup) {
            targetGroup = TABLE_STRUCTURE.find(g => g.group === 'Zusatz');
            if (!targetGroup) {
                targetGroup = { group: 'Zusatz', class: 'zusatz', columns: [] };
                TABLE_STRUCTURE.push(targetGroup);
            }
            insertIdx = targetGroup.columns.length;
        }

        pushUndo();
        targetGroup.columns.splice(insertIdx, 0, colName);
        AppState.newCols.add(colName);
        AppState.data.forEach(row => { if (row[colName] === undefined) row[colName] = ''; });
        renderTable();
        saveToStorage();
        showToast(`Spalte "${colName}" eingefügt`);
    });

    on('btnDeleteCol', 'click', () => {
        if (!AppState.selectedCell) { showToast('Bitte zuerst eine Spalte auswählen'); return; }
        const colToDelete = AppState.selectedCell.split('-')[1];
        if (!colToDelete) return;

        if (!confirm(`Spalte "${colToDelete}" wirklich löschen?`)) return;

        pushUndo();
        for (let g of TABLE_STRUCTURE) {
            const idx = g.columns.indexOf(colToDelete);
            if (idx !== -1) { g.columns.splice(idx, 1); break; }
        }
        AppState.newCols.delete(colToDelete);
        renderTable();
        saveToStorage();
        showToast(`Spalte "${colToDelete}" gelöscht`);
    });

    // Map buttons
    on('btnCloseMap', 'click', closeMap);
    on('btnMapStandard', 'click', () => switchLayer('standard'));
    on('btnMapSatellite', 'click', () => switchLayer('satellite'));
    on('btnUploadPlan', 'click', () => $('inputPlanImage').click());
    on('inputPlanImage', 'change', handlePlanUpload);
    on('btnGoToLocation', 'click', handleMapSearch);
    on('inpMapSearch', 'keypress', (e) => { if (e.key === 'Enter') handleMapSearch(); });
    on('btnLocateMe', 'click', startGPS);
    on('btnSetStartPoint', 'click', setStartPoint);
    on('btnShowMeasureLine', 'click', showMeasureLine);
    on('btnStopAudit', 'click', () => stopDrawing());

    // Depth markers & lines
    on('btnMarker08', 'click', () => setMarkerByDepth('#e879f9'));
    on('btnMarker16', 'click', () => setMarkerByDepth('#fbbf24'));
    on('btnMarker32', 'click', () => setMarkerByDepth('#22d3ee'));
    on('btnLine08', 'click', () => setLineByDepth('#e879f9'));
    on('btnLine16', 'click', () => setLineByDepth('#fbbf24'));
    on('btnLine32', 'click', () => setLineByDepth('#22d3ee'));

    // Visibility toggles
    on('btnToggle08', 'click', () => toggleLayerColor('#e879f9', 'btnToggle08'));
    on('btnToggle16', 'click', () => toggleLayerColor('#fbbf24', 'btnToggle16'));
    on('btnToggle32', 'click', () => toggleLayerColor('#22d3ee', 'btnToggle32'));

    // Numbered markers
    document.querySelectorAll('.num-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            activateNumberedPlacement(btn.dataset.num);
        });
    });

    // Snipping
    on('btnSnipRect', 'click', () => startSnip('rectangle'));
    on('btnSnipCircle', 'click', () => startSnip('circle'));
    on('btnSnipPoly', 'click', () => startSnip('polygon'));
    on('btnClearShapes', 'click', () => {
        if (layers.drawItems) layers.drawItems.clearLayers();
        // Reset all marker counters
        AppState.depthMarkers = { '0.8': [], '1.6': [], '3.2': [] };
        // Also clear saved map data so it doesn't restore old markers
        saveMapData();
        showToast('Alle Objekte gelöscht — Nummerierung zurückgesetzt');
    });

    // Snip editor
    on('btnSnipText', 'click', addTextToSnip);
    on('btnSnipRetry', 'click', () => { $('snipConfirmModal').style.display = 'none'; });
    on('btnDiscardSnip', 'click', () => { $('snipConfirmModal').style.display = 'none'; });
    on('btnSaveSnip', 'click', saveFinalSnip);
    on('btnSnipDownload', 'click', downloadSnip);
    on('snipTargetType', 'change', updateSnipColor);

    // Chart mode toggles
    on('btnChartBalken', 'click', () => {
        AppState.chartMode = 'bar';
        $('btnChartBalken').classList.add('active');
        $('btnChartScatter').classList.remove('active');
        renderAppPlot();
    });
    on('btnChartScatter', 'click', () => {
        AppState.chartMode = 'scatter';
        $('btnChartScatter').classList.add('active');
        $('btnChartBalken').classList.remove('active');
        renderAppPlot();
    });

    initEditorEvents();
}


// ─── FILE HANDLING ───
function handleFileUpload(file) {
    AppState.data = [];
    AppState.newCols.clear();

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

            let startIdx = -1;
            for (let i = 0; i < Math.min(15, allRows.length); i++) {
                if (JSON.stringify(allRows[i]).toLowerCase().includes('kennzeichen')) {
                    startIdx = i + 1;
                    break;
                }
            }

            if (startIdx === -1) {
                showToast('Ungültiges Datei-Format');
                return;
            }

            const mapped = [];
            for (let i = startIdx; i < allRows.length; i++) {
                const r = allRows[i];
                if (!r || r.length < 2) continue;
                mapped.push({
                    'Kennzeichen': String(r[1] || ''),
                    'Alt-Kz.': String(r[2] || ''),
                    'Typ': String(r[3] || ''),
                    'Örtlichkeit': String(r[5] || ''),
                    'Meter [m]': String(r[4] || '')
                });
            }

            AppState.data = mapped;

            const fileStatus = $('fileStatus');
            if (fileStatus) fileStatus.textContent = `Datei: ${file.name}`;

            const btnStart = $('btnFinalStart');
            if (btnStart) btnStart.disabled = false;

            const pInput = $('projectName');
            if (pInput) pInput.value = file.name.split(/[._\s-]/)[0];

            showToast(`${mapped.length} Messstellen geladen`);
        } catch (err) {
            console.error('Import error:', err);
            showToast('Import-Fehler: ' + err.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

// ─── TABLE RENDERING ───
function renderTable() {
    const thead = $('tableHead');
    const tbody = $('tableBody');
    if (!thead || !tbody) return;

    // Build header
    thead.innerHTML = '';
    const tr1 = document.createElement('tr');
    const tr2 = document.createElement('tr');
    tr1.innerHTML = '<th rowspan="2" style="width:40px;text-align:center;">#</th>';

    TABLE_STRUCTURE.forEach(g => {
        if (AppState.hiddenColumns.has(g.class)) return;
        const groupCols = g.columns;
        if (groupCols.length === 0) return;
        const cs = DEPTH_COLORS[g.class] || DEPTH_COLORS.basis;

        const th = document.createElement('th');
        th.textContent = g.group;
        th.colSpan = groupCols.length;
        th.style.borderBottom = `3px solid ${cs.border}`;
        th.style.borderLeft = `3px solid ${cs.border}`;
        th.style.color = cs.text;
        th.style.background = cs.bg;
        tr1.appendChild(th);

        groupCols.forEach((c, cIdx) => {
            const sh = document.createElement('th');
            sh.style.minWidth = AppState.columnWidths[c] || '80px';
            sh.style.color = cs.text;
            // First column of each group gets colored left border
            if (cIdx === 0) {
                sh.style.borderLeft = `3px solid ${cs.border}`;
            }
            // ALL sub-header cells get bottom border in group color
            sh.style.borderBottom = `2px solid ${cs.border}`;
            const label = c.includes('_') ? c.split('_')[0] : c;
            sh.textContent = label;

            if (c === 'Kennzeichen') {
                const fi = document.createElement('input');
                fi.type = 'text';
                fi.placeholder = 'Filter...';
                fi.className = 'header-filter-input';
                fi.value = AppState.filters['Kennzeichen'] || '';
                fi.onclick = (e) => e.stopPropagation();
                fi.oninput = (e) => {
                    AppState.filters['Kennzeichen'] = e.target.value;
                    renderTableBody();
                };
                sh.appendChild(fi);
            }
            tr2.appendChild(sh);
        });
    });
    thead.append(tr1, tr2);

    renderTableBody();
    renderAppPlot();
}

function renderTableBody() {
    const tbody = $('tableBody');
    if (!tbody) return;
    tbody.innerHTML = '';

    let filteredData = AppState.data.map((row, index) => ({ ...row, _originalIndex: index }));
    const kzFilter = (AppState.filters['Kennzeichen'] || '').trim().toLowerCase();

    if (kzFilter) {
        if (kzFilter.includes('-')) {
            const parts = kzFilter.split('-').map(p => p.trim());
            if (parts.length === 2) {
                const [s, e] = parts;
                filteredData = filteredData.filter(row => {
                    const v = (row['Kennzeichen'] || '').toString().toLowerCase();
                    if (!isNaN(s) && !isNaN(e) && !isNaN(v)) return Number(v) >= Number(s) && Number(v) <= Number(e);
                    return v >= s && v <= e;
                });
            }
        } else {
            filteredData = filteredData.filter(row => (row['Kennzeichen'] || '').toString().toLowerCase().includes(kzFilter));
        }
    }

    filteredData.forEach((row) => {
        const tr = document.createElement('tr');
        const idx = row._originalIndex;
        const originalRow = AppState.data[idx];

        // Highlight new rows with visible background (tan/brown)
        if (originalRow._isNew) {
            tr.style.background = 'rgba(210, 180, 140, 0.15)';
        }

        tr.innerHTML = `<td style="text-align:center;font-weight:700;color:var(--text-muted);width:40px;">${idx + 1}</td>`;

        TABLE_STRUCTURE.forEach(g => {
            if (AppState.hiddenColumns.has(g.class)) return;
            const cs = DEPTH_COLORS[g.class] || DEPTH_COLORS.basis;

            g.columns.forEach((col, colIdx) => {
                const td = document.createElement('td');

                // First cell of each group gets a colored left border
                if (colIdx === 0) {
                    td.style.borderLeft = `3px solid ${cs.border}`;
                }

                if (col.toLowerCase().includes('anhang')) {
                    td.style.textAlign = 'center';
                    if (originalRow[col]) {
                        const viewBtn = document.createElement('button');
                        viewBtn.className = 'toolbar-btn';
                        viewBtn.innerHTML = '<i class="fas fa-eye"></i>';
                        viewBtn.onclick = () => openImagePreview(originalRow[col]);
                        td.appendChild(viewBtn);

                        const delBtn = document.createElement('button');
                        delBtn.className = 'toolbar-btn';
                        delBtn.style.color = 'var(--danger)';
                        delBtn.innerHTML = '<i class="fas fa-trash"></i>';
                        delBtn.onclick = () => {
                            if (confirm('Foto löschen?')) {
                                delete originalRow[col];
                                renderTable();
                                saveToStorage();
                            }
                        };
                        td.appendChild(delBtn);
                    } else {
                        td.textContent = '-';
                        td.style.color = 'var(--text-muted)';
                    }
                } else {
                    const inp = document.createElement('input');
                    inp.type = 'text';
                    inp.value = originalRow[col] || '';
                    inp.className = 'cell-input';
                    inp.dataset.col = col;
                    inp.dataset.row = idx;

                    if (col.startsWith('\u03C1') || col.startsWith('MW') || col.startsWith('SD')) {
                        inp.readOnly = true;
                        inp.style.color = 'var(--text-muted)';
                        inp.style.fontWeight = '600';
                    }

                    // Numeric input mode for measurement fields
                    if (col.startsWith('R') && col.includes('[')) {
                        inp.inputMode = 'decimal';
                    }

                    // New row or new column highlight (tan/brown)
                    if (originalRow._isNew || AppState.newCols.has(col)) {
                        inp.style.background = 'rgba(210, 180, 140, 0.35)';
                    }

                    inp.onfocus = () => { AppState.selectedCell = `${idx}-${col}`; };
                    inp.oninput = (e) => {
                        originalRow[col] = e.target.value;
                        updateLiveCalculations(originalRow, col, tr);
                        saveToStorage();
                    };

                    // Apply formatting if exists
                    if (typeof cellFormatting !== 'undefined' && cellFormatting[`${idx}-${col}`]) {
                        if (typeof applyFormattingToCell === 'function') {
                            applyFormattingToCell(inp, cellFormatting[`${idx}-${col}`]);
                        }
                    }

                    td.appendChild(inp);
                }
                tr.appendChild(td);
            });
        });
        tbody.appendChild(tr);
    });
}

function updateLiveCalculations(row, col, tr) {
    if (!(col.startsWith('R') && col.includes('_'))) return;
    const parts = col.split('_');
    if (parts.length < 2) return;

    const dStr = parts[1];
    const d = parseFloat(dStr);
    const n = col.substring(1, 2);
    const rVal = parseFloat((row[col] || '0').replace(',', '.'));

    if (isNaN(rVal) || isNaN(d)) return;

    const rhoKey = `\u03C1${n} [\u03A9m]_${dStr}`;
    row[rhoKey] = (2 * Math.PI * d * rVal).toFixed(2);

    let sum = 0, count = 0;
    for (let k = 1; k <= 3; k++) {
        const rk = parseFloat((row[`R${k} [\u03A9]_${dStr}`] || '0').replace(',', '.'));
        if (!isNaN(rk) && rk > 0) { sum += rk; count++; }
    }

    if (count > 0) {
        const mwR = sum / count;
        row[`MW [\u03A9m]_${dStr}`] = (2 * Math.PI * d * mwR).toFixed(2);

        if (count > 1) {
            let sqSum = 0;
            for (let k = 1; k <= 3; k++) {
                const rk = parseFloat((row[`R${k} [\u03A9]_${dStr}`] || '0').replace(',', '.'));
                if (!isNaN(rk) && rk > 0) sqSum += Math.pow(rk - mwR, 2);
            }
            row[`SD [\u03A9m]_${dStr}`] = (2 * Math.PI * d * Math.sqrt(sqSum / (count - 1))).toFixed(2);
        } else {
            row[`SD [\u03A9m]_${dStr}`] = '0.00';
        }
    }

    // Update displayed values in the row
    if (tr) {
        tr.querySelectorAll('input').forEach(inp => {
            const tc = inp.dataset.col;
            if (tc && row[tc] !== undefined && tc !== col) inp.value = row[tc];
        });
    }
}

// ─── STORAGE ───
function saveToStorage() {
    try {
        const pName = ($('projectName') || {}).value || 'Unbenanntes Projekt';

        if (layers.drawItems) {
            layers.drawItems.eachLayer(l => {
                if (!l.feature) l.feature = { type: 'Feature', properties: {} };
                if (l.options && l.options.markerColor) l.feature.properties.markerColor = l.options.markerColor;
                if (!l.feature.properties.markerColor && l.options && l.options.color) l.feature.properties.markerColor = l.options.color;
                if (l.options && l.options.isNumbered) {
                    l.feature.properties.isNumbered = true;
                    l.feature.properties.markerNumber = l.options.markerNumber;
                }
            });
        }

        const project = {
            name: pName,
            data: AppState.data,
            mapData: layers.drawItems ? layers.drawItems.toGeoJSON() : null,
            timestamp: Date.now(),
            hiddenMapColors: Array.from(AppState.hiddenMapColors),
            newCols: Array.from(AppState.newCols)
        };

        let library = JSON.parse(localStorage.getItem('messstellen_library') || '{}');
        library[pName] = project;
        localStorage.setItem('messstellen_library', JSON.stringify(library));
        localStorage.setItem('current_project_id', pName);
    } catch (err) {
        console.error('Save error:', err);
        if (err.name === 'QuotaExceededError') {
            showToast('Speicher voll! Bitte alte Projekte löschen.');
        }
    }
}

function saveMapData() {
    saveToStorage();
}

function loadFromStorage() {
    const pName = localStorage.getItem('current_project_id');
    if (!pName) return;

    let library = JSON.parse(localStorage.getItem('messstellen_library') || '{}');
    const project = library[pName];

    if (project) {
        AppState.data = project.data || [];
        const nameInput = $('projectName');
        if (nameInput) nameInput.value = project.name;

        AppState.hiddenMapColors = new Set(project.hiddenMapColors || []);
        AppState.newCols = new Set(project.newCols || []);

        if (AppState.newCols.size > 0) {
            let zusatzGroup = TABLE_STRUCTURE.find(g => g.group === 'Zusatz');
            if (!zusatzGroup) {
                zusatzGroup = { group: 'Zusatz', class: 'zusatz', columns: [] };
                TABLE_STRUCTURE.push(zusatzGroup);
            }
            AppState.newCols.forEach(colName => {
                if (!zusatzGroup.columns.includes(colName)) zusatzGroup.columns.push(colName);
            });
        }
    }
}

function createEmptyRows(n) {
    for (let i = 0; i < n; i++) AppState.data.push({});
}


// ─── MAP MODULE ───
function openMap() {
    const m = $('mapLayout');
    if (m) {
        m.style.display = 'flex';
        m.classList.add('show');
        if (!map) initMap();
        setTimeout(() => {
            if (map) {
                map.invalidateSize();
                refreshMapVisibility();
            }
        }, 300);
    }
}

function closeMap() {
    const m = $('mapLayout');
    if (m) { m.style.display = 'none'; m.classList.remove('show'); }
}

function initMap() {
    if (map) return;
    const mapContainer = $('map');
    if (!mapContainer) return;

    map = L.map('map', {
        zoomControl: false,
        attributionControl: true,
        doubleClickZoom: false,
        preferCanvas: true
    }).setView([51.1657, 10.4515], 6);

    L.control.zoom({ position: 'bottomleft' }).addTo(map);

    // Suppress draw tooltips
    try {
        L.drawLocal.draw.handlers.rectangle.tooltip.start = '';
        L.drawLocal.draw.handlers.circle.tooltip.start = '';
        L.drawLocal.draw.handlers.polygon.tooltip.start = '';
        L.drawLocal.draw.handlers.polyline.tooltip.start = '';
        L.drawLocal.draw.handlers.simpleshape.tooltip.end = '';
    } catch (e) { /* ignore if not available */ }

    // Tile layers
    layers.standard = L.tileLayer('https://{s}.google.com/vt/lyrs=m&x={x}&y={y}&z={z}', {
        maxZoom: 22, maxNativeZoom: 20,
        subdomains: ['mt0', 'mt1', 'mt2', 'mt3'],
        attribution: '&copy; Google Maps'
    });

    layers.satellite = L.tileLayer('https://{s}.google.com/vt/lyrs=y&x={x}&y={y}&z={z}', {
        maxZoom: 22, maxNativeZoom: 20,
        subdomains: ['mt0', 'mt1', 'mt2', 'mt3'],
        attribution: '&copy; Google Maps'
    });

    layers.standard.addTo(map);
    layers.drawItems = new L.FeatureGroup().addTo(map);

    // Draw events
    map.on(L.Draw.Event.CREATED, (e) => {
        const layer = e.layer;
        const color = AppState.activeColor;

        if (captureMode) {
            layer.setStyle({ color: 'transparent', fillOpacity: 0 });
            map.addLayer(layer);
            triggerAusschnitt(layer);
            captureMode = false;
            return;
        }

        if (layer instanceof L.Marker) {
            const latlng = layer.getLatLng();
            const coloredMarker = createPinMarker(latlng, color);
            coloredMarker.addTo(layers.drawItems);
            setupInteractiveLayer(coloredMarker);
        } else {
            layer.setStyle({ color: color, weight: 6 });
            layer.options.markerColor = color;
            layers.drawItems.addLayer(layer);
            setupInteractiveLayer(layer);
        }

        if (AppState.activeDrawTool) {
            setTimeout(() => { if (AppState.activeDrawTool) AppState.activeDrawTool.enable(); }, 50);
        }
        saveMapData();
    });

    // Location events
    map.on('locationfound', handleLocationFound);
    map.on('locationerror', handleLocationError);
    map.on('dragstart', () => {
        if (AppState.liveFollow) {
            AppState.liveFollow = false;
            showToast('Auto-Follow deaktiviert');
        }
    });

    restoreMapDrawings();
}

function createPinMarker(latlng, color) {
    // Count from depthMarkers array (resets when "Löschen" is clicked)
    var depthKey = color === '#e879f9' ? '0.8' : (color === '#fbbf24' ? '1.6' : '3.2');
    if (!AppState.depthMarkers) AppState.depthMarkers = { '0.8': [], '1.6': [], '3.2': [] };
    var count = AppState.depthMarkers[depthKey].length;
    console.log('createPinMarker: depth=' + depthKey + ', array length=' + AppState.depthMarkers[depthKey].length + ', count=' + count);

    const pinSvg = `<svg width="22" height="28" viewBox="0 0 22 28" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M11 0C4.9 0 0 4.9 0 11C0 19.25 11 28 11 28C11 28 22 19.25 22 11C22 4.9 17.1 0 11 0Z" fill="${color}"/><circle cx="11" cy="11" r="6" fill="white"/><text x="11" y="14" text-anchor="middle" font-size="9" font-weight="bold" font-family="Inter,sans-serif" fill="${color}">${count}</text></svg>`;
    return L.marker(latlng, {
        draggable: true,
        interactive: true,
        zIndexOffset: 1000,
        markerColor: color,
        icon: L.divIcon({
            html: `<div style="width:22px;height:28px;filter:drop-shadow(0 1px 3px rgba(0,0,0,0.35));">${pinSvg}</div>`,
            iconSize: [22, 28], iconAnchor: [11, 28], className: 'custom-marker-icon'
        })
    });
}

function setupInteractiveLayer(layer) {
    layer.on('click', (e) => {
        L.DomEvent.stopPropagation(e);
        if (confirm('Objekt löschen?')) {
            if (layers.drawItems) {
                layers.drawItems.removeLayer(layer);

                // Decrement depthMarkers counter for this color
                var color = layer.options && layer.options.markerColor;
                if (color && AppState.depthMarkers && layer.getLatLng && !layer.options.isNumbered && !layer.options.isDistLabel) {
                    var depthKey = color === '#e879f9' ? '0.8' : (color === '#fbbf24' ? '1.6' : '3.2');
                    if (AppState.depthMarkers[depthKey] && AppState.depthMarkers[depthKey].length > 0) {
                        AppState.depthMarkers[depthKey].pop();
                    }
                }

                saveMapData();
                showToast('Objekt gelöscht');
            }
        }
    });
}

function switchLayer(key) {
    if (!map) return;
    if (layers.standard) map.removeLayer(layers.standard);
    if (layers.satellite) map.removeLayer(layers.satellite);

    if (key === 'satellite') {
        layers.satellite.addTo(map);
        $('btnMapSatellite').classList.add('active');
        $('btnMapStandard').classList.remove('active');
    } else {
        layers.standard.addTo(map);
        $('btnMapStandard').classList.add('active');
        $('btnMapSatellite').classList.remove('active');
    }
    setTimeout(() => map.invalidateSize(), 100);
}

function handlePlanUpload(e) {
    const file = e.target.files[0];
    if (!file || !map) return;
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
        const center = map.getCenter();
        const latOff = (img.height / 20) / 111320;
        const lngOff = (img.width / 20) / (111320 * Math.cos(center.lat * Math.PI / 180));
        const bounds = L.latLngBounds([center.lat - latOff, center.lng - lngOff], [center.lat + latOff, center.lng + lngOff]);
        L.imageOverlay(url, bounds).addTo(map);
        map.fitBounds(bounds);
        showToast('Karte geladen');
    };
    img.src = url;
}

let searchMarker = null;
async function handleMapSearch() {
    if (searchMarker && map) map.removeLayer(searchMarker);
    const query = ($('inpMapSearch') || {}).value || '';
    if (!query.trim()) return;

    const cleanQuery = query.replace(',', '.');
    const coordMatch = cleanQuery.match(/^(-?\d+\.?\d*)\s*[\s,]\s*(-?\d+\.?\d*)$/);

    if (coordMatch) {
        const lat = parseFloat(coordMatch[1]), lng = parseFloat(coordMatch[2]);
        map.setView([lat, lng], 19);
        searchMarker = L.marker([lat, lng]).addTo(map).bindPopup(`${lat}, ${lng}`).openPopup();
        showToast(`Koordinaten: ${lat.toFixed(5)}, ${lng.toFixed(5)}`);
        return;
    }

    try {
        $('loadingOverlay').style.display = 'flex';
        const resp = await fetch(`https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(query)}`);
        const results = await resp.json();
        $('loadingOverlay').style.display = 'none';

        if (results.length > 0) {
            const res = results[0];
            map.setView([res.lat, res.lon], 19);
            searchMarker = L.marker([res.lat, res.lon]).addTo(map).bindPopup(res.display_name).openPopup();
            showToast(`Gefunden: ${res.display_name.split(',')[0]}`);
        } else {
            showToast('Standort nicht gefunden');
        }
    } catch (err) {
        $('loadingOverlay').style.display = 'none';
        showToast('Suche fehlgeschlagen: ' + err.message);
    }
}

function startGPS() {
    if (!map) return;
    showToast('GPS-Suche aktiv...');
    AppState.firstLocationFound = false;
    map.locate({ setView: false, maxZoom: 18, enableHighAccuracy: true, watch: true, timeout: 30000 });
    initCompass();
}

function handleLocationFound(e) {
    const latlng = e.latlng;
    const acc = Math.round(e.accuracy);

    if (AppState.userMarker) map.removeLayer(AppState.userMarker);
    if (AppState.accuracyCircle) map.removeLayer(AppState.accuracyCircle);

    const userIcon = L.divIcon({
        html: `<div style="background:var(--accent,#3b82f6);width:28px;height:28px;border-radius:50%;border:3px solid white;box-shadow:0 2px 8px rgba(0,0,0,0.4);display:flex;align-items:center;justify-content:center;"><i class="fas fa-user" style="color:white;font-size:12px;"></i></div>`,
        iconSize: [28, 28], iconAnchor: [14, 14], className: 'user-icon'
    });

    AppState.userMarker = L.marker(latlng, { icon: userIcon, zIndexOffset: 2000 }).addTo(map);
    AppState.accuracyCircle = L.circle(latlng, Math.min(acc, 50), {
        color: '#3b82f6', fillColor: 'rgba(59,130,246,0.1)', weight: 1
    }).addTo(map);

    if (!AppState.firstLocationFound) {
        map.flyTo(latlng, 18, { animate: true, duration: 1.2 });
        AppState.firstLocationFound = true;
    } else if (AppState.liveFollow) {
        map.panTo(latlng);
    }

    // Update GPS status
    const statusText = $('gpsStatusText');
    const statusDot = $('gpsAccuracyDot');
    if (statusText) statusText.textContent = `GPS: ${acc}m`;
    if (statusDot) {
        statusDot.style.background = acc < 10 ? '#10b981' : (acc < 30 ? '#f59e0b' : '#ef4444');
        statusDot.classList.add('pulse');
    }

    // Distance calculation
    if (AppState.startPoint) {
        const dist = map.distance(latlng, AppState.startPoint);
        const badge = $('distanceBadge');
        const val = $('distanceValue');
        if (badge && val) { val.textContent = dist.toFixed(1); badge.style.display = 'flex'; }
    }
}

function handleLocationError(e) {
    const messages = {
        1: 'Standort-Berechtigung verweigert',
        2: 'Position nicht verfügbar',
        3: 'Zeitüberschreitung'
    };
    showToast('GPS: ' + (messages[e.code] || e.message));
}

function initCompass() {
    const compass = $('mapCompass');
    if (!compass) return;

    const handleOrientation = (event) => {
        let heading = event.webkitCompassHeading || (event.alpha ? 360 - event.alpha : null);
        if (heading != null) compass.style.transform = `rotate(${-heading}deg)`;
    };

    if (typeof DeviceOrientationEvent !== 'undefined' && typeof DeviceOrientationEvent.requestPermission === 'function') {
        DeviceOrientationEvent.requestPermission()
            .then(r => { if (r === 'granted') window.addEventListener('deviceorientation', handleOrientation, true); })
            .catch(() => {});
    } else {
        window.addEventListener('deviceorientationabsolute', handleOrientation, true);
        window.addEventListener('deviceorientation', handleOrientation, true);
    }
}

function setStartPoint() {
    if (!map) return;
    
    // Always use current GPS position if available
    if (!AppState.userMarker) {
        showToast('Bitte zuerst GPS aktivieren (GPS-Button klicken)');
        return;
    }
    
    const latlng = AppState.userMarker.getLatLng();
    AppState.startPoint = latlng;

    // Reset marker groups — new start = new group
    AppState.depthMarkers = { '0.8': [], '1.6': [], '3.2': [] };

    if (AppState.startMarker) map.removeLayer(AppState.startMarker);
    AppState.startMarker = L.marker(latlng, {
        icon: L.divIcon({
            html: '<div style="background:#ef4444;width:20px;height:20px;border-radius:50%;border:3px solid white;box-shadow:0 2px 8px rgba(0,0,0,0.4);"></div>',
            iconSize: [20, 20], iconAnchor: [10, 10], className: 'start-marker'
        }),
        zIndexOffset: 3000
    }).addTo(map);
    
    // Zoom to current position
    map.setView(latlng, 20, { animate: true });
    showToast('Startpunkt an GPS-Position gesetzt');
}

// ─── MEASUREMENT LINE (Distance Reference) ───
function showMeasureLine() {
    if (!map) return;
    if (!AppState.userMarker) {
        showToast('Bitte zuerst GPS aktivieren');
        return;
    }

    // Remove old
    if (AppState.measureCircles) {
        AppState.measureCircles.forEach(function(c) { map.removeLayer(c); });
    }
    AppState.measureCircles = [];

    var center = AppState.userMarker.getLatLng();
    var depths = [
        { dist: 0.8, color: '#e879f9', label: '0.8m' },
        { dist: 1.6, color: '#fbbf24', label: '1.6m' },
        { dist: 3.2, color: '#22d3ee', label: '3.2m' }
    ];

    depths.forEach(function(d) {
        var circle = L.circle(center, {
            radius: d.dist, color: d.color, weight: 3,
            fillColor: d.color, fillOpacity: 0.05, dashArray: '8, 5'
        }).addTo(map);

        var labelPos = L.latLng(center.lat + (d.dist / 111320), center.lng);
        var label = L.marker(labelPos, {
            interactive: false,
            icon: L.divIcon({
                html: '<div style="background:' + d.color + ';color:#000;padding:3px 10px;border-radius:5px;font-size:12px;font-weight:700;white-space:nowrap;">' + d.label + '</div>',
                iconSize: [60, 22], iconAnchor: [30, 11], className: 'measure-label'
            })
        }).addTo(map);

        AppState.measureCircles.push(circle);
        AppState.measureCircles.push(label);
    });

    map.setView(center, 22, { animate: true });
    showToast('Elektrodenabstände: 0.8m / 1.6m / 3.2m');
}

// ─── GEO HELPERS (snap-to-distance) ───
function calcBearing(from, to) {
    var dLng = (to.lng - from.lng) * Math.PI / 180;
    var lat1 = from.lat * Math.PI / 180;
    var lat2 = to.lat * Math.PI / 180;
    var y = Math.sin(dLng) * Math.cos(lat2);
    var x = Math.cos(lat1) * Math.sin(lat2) - Math.sin(lat1) * Math.cos(lat2) * Math.cos(dLng);
    return Math.atan2(y, x);
}

function destPoint(from, bearing, distMeters) {
    var R = 6371000; // Earth radius in meters
    var d = distMeters / R;
    var lat1 = from.lat * Math.PI / 180;
    var lng1 = from.lng * Math.PI / 180;
    var lat2 = Math.asin(Math.sin(lat1) * Math.cos(d) + Math.cos(lat1) * Math.sin(d) * Math.cos(bearing));
    var lng2 = lng1 + Math.atan2(Math.sin(bearing) * Math.sin(d) * Math.cos(lat1), Math.cos(d) - Math.sin(lat1) * Math.sin(lat2));
    return L.latLng(lat2 * 180 / Math.PI, lng2 * 180 / Math.PI);
}

function setMarkerByDepth(color) {
    stopDrawing(true);
    if (!map) return;
    var depthLabel = color === '#e879f9' ? '0.8m' : (color === '#fbbf24' ? '1.6m' : '3.2m');
    var depthDist = color === '#e879f9' ? 0.8 : (color === '#fbbf24' ? 1.6 : 3.2);
    AppState.activeDrawToolName = 'sticky-marker';
    AppState.activeColor = color;
    AppState.activeDepthLabel = depthLabel;
    $('map').style.cursor = 'crosshair';
    map.on('click', handleStickyMarkerClick);

    // Show circle for this depth only — visual indicator (larger than real for visibility)
    if (AppState.measureCircles) {
        AppState.measureCircles.forEach(function(c) { map.removeLayer(c); });
    }
    AppState.measureCircles = [];
    var center = AppState.startPoint || (AppState.userMarker ? AppState.userMarker.getLatLng() : null);
    if (center) {
        // Precise radius: 3 × electrode spacing (total Wenner array length)
        // 0.8m → 2.4m, 1.6m → 4.8m, 3.2m → 9.6m
        var circleRadius = depthDist * 3;
        var circle = L.circle(center, {
            radius: circleRadius, color: color, weight: 4,
            fillColor: color, fillOpacity: 0.08, dashArray: '10, 6'
        }).addTo(map);
        AppState.measureCircles.push(circle);

        // Zoom to see the precise circle clearly
        map.fitBounds(circle.getBounds().pad(1.5), { animate: true, maxZoom: 22 });
    }

    showToast(depthLabel + ' Marker — auf Karte tippen');
}

function showDepthConfirmation(color, depthLabel, callback) {
    var modal = $('depthConfirmModal');
    if (!modal) { callback(); return; }
    var icon = $('depthConfirmIcon');
    var title = $('depthConfirmTitle');
    var text = $('depthConfirmText');
    var check = $('depthConfirmCheck');
    var okBtn = $('depthConfirmOk');
    var cancelBtn = $('depthConfirmCancel');

    if (!icon || !title || !text || !check || !okBtn || !cancelBtn) { callback(); return; }

    icon.style.background = color;
    icon.innerHTML = '<i class="fas fa-ruler-combined" style="color:#000;"></i>';
    title.textContent = 'Messtiefe: ' + depthLabel;
    text.innerHTML = 'Elektrodenabstand (a) = <strong>' + depthLabel + '</strong><br>Bitte bestätigen Sie den korrekten Aufbau.';

    check.checked = false;
    okBtn.disabled = true;
    check.onchange = function() { okBtn.disabled = !check.checked; };
    cancelBtn.onclick = function() {
        modal.style.display = 'none';
        // Re-enable click listener since user cancelled
        map.on('click', handleStickyMarkerClick);
    };
    okBtn.onclick = function() {
        modal.style.display = 'none';
        callback();
    };
    modal.style.display = 'flex';
}

function handleStickyMarkerClick(e) {
    var color = AppState.activeColor;
    var depthLabel = AppState.activeDepthLabel || '0.8m';
    var latlng = e.latlng;
    var depthKey = depthLabel.replace('m', '');

    // Stop clicks while confirming
    map.off('click', handleStickyMarkerClick);

    // Confirm electrode spacing
    showDepthConfirmation(color, depthLabel, function() {
        if (!AppState.depthMarkers) AppState.depthMarkers = { '0.8': [], '1.6': [], '3.2': [] };
        var markers = AppState.depthMarkers[depthKey];
        var requiredDist = parseFloat(depthKey); // 0.8, 1.6, or 3.2 meters

        // ── Determine previous point BEFORE pushing ──
        var futureNum = markers.length + 1;
        var prevLatLng = null;
        var posInGroup = ((futureNum - 1) % 3); // 0 = first in group, 1 = second, 2 = third
        if (posInGroup === 0 && AppState.startPoint) {
            prevLatLng = AppState.startPoint;
        } else if (futureNum >= 2) {
            prevLatLng = markers[futureNum - 2];
        }

        // ── Auto-snap to exact electrode spacing ──
        var snapped = false;
        var originalDist = 0;
        if (prevLatLng) {
            originalDist = map.distance(prevLatLng, latlng);
            if (Math.abs(originalDist - requiredDist) > 0.05) {
                var bearing = calcBearing(prevLatLng, latlng);
                latlng = destPoint(prevLatLng, bearing, requiredDist);
                snapped = true;
            }
        }

        // ── Push snapped position & create marker ──
        markers.push(latlng);
        var markerNum = markers.length;
        var pinMarker = createPinMarker(latlng, color);

        // Create a group layer (marker + line + label together)
        var groupLayer = L.featureGroup();
        groupLayer.addLayer(pinMarker);
        groupLayer.options = { markerColor: color };

        if (prevLatLng) {
            var finalDist = map.distance(prevLatLng, latlng);
            var distText = finalDist.toFixed(2) + ' m';

            var line = L.polyline([prevLatLng, latlng], {
                color: color, weight: 2, opacity: 0.6, dashArray: '4, 4'
            });
            groupLayer.addLayer(line);

            // Distance label — right beside the marker pin
            var distLabel = L.marker(latlng, {
                interactive: false,
                isDistLabel: true,
                icon: L.divIcon({
                    html: '<div style="background:rgba(0,0,0,0.7);color:#fff;padding:1px 4px;border-radius:3px;font-size:8px;font-weight:600;white-space:nowrap;letter-spacing:0.3px;">' + distText + '</div>',
                    iconSize: [42, 14], iconAnchor: [-14, 18], className: 'dist-label'
                })
            });
            groupLayer.addLayer(distLabel);

            if (snapped) {
                showToast(depthLabel + ' #' + markerNum + ' — ' + originalDist.toFixed(1) + 'm → ' + requiredDist.toFixed(1) + 'm ✓');
            } else {
                showToast(depthLabel + ' #' + markerNum + ' — ' + distText);
            }
        } else {
            showToast(depthLabel + ' #' + markerNum);
        }

        // Add group to map and setup delete (deletes marker + line + label together)
        groupLayer.addTo(layers.drawItems);
        groupLayer.on('click', function(ev) {
            L.DomEvent.stopPropagation(ev);
            if (confirm('Marker ' + markerNum + ' mit Linie l\u00F6schen?')) {
                layers.drawItems.removeLayer(groupLayer);
                // Decrement counter
                if (AppState.depthMarkers[depthKey].length > 0) {
                    AppState.depthMarkers[depthKey].pop();
                }
                saveMapData();
                showToast('Marker gel\u00F6scht');
            }
        });

        saveMapData();

        // Re-enable for next marker
        map.on('click', handleStickyMarkerClick);
    });
}

function stopDrawing(silent) {
    if (AppState.activeDrawTool && AppState.activeDrawTool.disable) AppState.activeDrawTool.disable();
    AppState.activeDrawTool = null;
    AppState.activeDrawToolName = null;

    if (map) {
        map.off('click', handleStickyMarkerClick);
        if (AppState._currentNumHandler) {
            map.off('click', AppState._currentNumHandler);
            AppState._currentNumHandler = null;
        }
    }

    // Clear measurement circle only when explicitly stopping (not when switching modes)
    if (!silent && AppState.measureCircles) {
        AppState.measureCircles.forEach(function(c) { if (map) map.removeLayer(c); });
        AppState.measureCircles = [];
    }
    var hud = document.getElementById('measureHud');
    if (hud) hud.remove();

    var mapEl = $('map');
    if (mapEl) mapEl.style.cursor = '';
    if (!silent) showToast('Zeichenmodus beendet');
}

function setLineByDepth(color) {
    stopDrawing(true);
    if (!map) return;
    AppState.activeColor = color;
    AppState.activeDrawTool = new L.Draw.Polyline(map, {
        repeatMode: true,
        shapeOptions: { color: color, weight: 5, lineCap: 'round', lineJoin: 'round' }
    });
    AppState.activeDrawTool.enable();
    $('map').style.cursor = 'crosshair';
    showToast('Linien-Modus aktiv');
}

function activateNumberedPlacement(num) {
    stopDrawing(true);
    if (!map) return;
    AppState.activeDrawToolName = `num-${num}`;
    AppState.activeColor = AppState.activeColor || '#e879f9';
    $('map').style.cursor = 'crosshair';

    const handler = (e) => {
        const color = AppState.activeColor;
        const marker = L.marker(e.latlng, {
            draggable: true, interactive: true, zIndexOffset: 1000,
            markerColor: color, markerNumber: num, isNumbered: true,
            icon: L.divIcon({
                html: `<div style="background:${color};color:#000;width:28px;height:28px;border-radius:50%;display:flex;align-items:center;justify-content:center;border:2px solid rgba(0,0,0,0.3);font-weight:bold;font-size:14px;box-shadow:0 2px 6px rgba(0,0,0,0.3);">${num}</div>`,
                iconSize: [28, 28], iconAnchor: [14, 14], className: 'numbered-marker'
            })
        });
        setupInteractiveLayer(marker);
        marker.addTo(layers.drawItems);
        saveMapData();
    };

    map.on('click', handler);
    AppState._currentNumHandler = handler;
    showToast(`Nummerierung ${num} aktiv`);
}

window.toggleLayerColor = function(color, btnId) {
    const colorLower = color.toLowerCase();
    const isHidden = AppState.hiddenMapColors.has(colorLower);

    if (layers.drawItems) {
        layers.drawItems.eachLayer(layer => {
            const c = (layer.options && (layer.options.markerColor || layer.options.color)) || '';
            if (c.toLowerCase() !== colorLower) return;
            const el = layer.getElement ? layer.getElement() : null;
            if (isHidden) {
                if (el) el.style.display = '';
                if (layer.setStyle) layer.setStyle({ opacity: 1, fillOpacity: 0.5 });
                if (layer.setOpacity) layer.setOpacity(1);
            } else {
                if (el) el.style.display = 'none';
                if (layer.setStyle) layer.setStyle({ opacity: 0, fillOpacity: 0 });
                if (layer.setOpacity) layer.setOpacity(0);
            }
        });
    }

    const btn = $(btnId);
    if (isHidden) {
        AppState.hiddenMapColors.delete(colorLower);
        if (btn) { btn.style.opacity = '1'; const i = btn.querySelector('i'); if (i) i.className = 'fas fa-eye'; }
    } else {
        AppState.hiddenMapColors.add(colorLower);
        if (btn) { btn.style.opacity = '0.4'; const i = btn.querySelector('i'); if (i) i.className = 'fas fa-eye-slash'; }
    }
    saveToStorage();
};

function refreshMapVisibility() {
    if (!layers.drawItems) return;
    layers.drawItems.eachLayer(layer => {
        const color = (layer.options && (layer.options.markerColor || layer.options.color)) || '';
        if (!color) return;
        const isHidden = AppState.hiddenMapColors.has(color.toLowerCase());
        const el = layer.getElement ? layer.getElement() : null;
        if (isHidden) {
            if (el) el.style.display = 'none';
            if (layer.setStyle) layer.setStyle({ opacity: 0, fillOpacity: 0 });
            if (layer.setOpacity) layer.setOpacity(0);
        } else {
            if (el) el.style.display = '';
            if (layer.setStyle) layer.setStyle({ opacity: 1, fillOpacity: 0.5 });
            if (layer.setOpacity) layer.setOpacity(1);
        }
    });
}

function restoreMapDrawings() {
    const pName = ($('projectName') || {}).value;
    if (!pName) return;
    let library = JSON.parse(localStorage.getItem('messstellen_library') || '{}');
    const project = library[pName];

    if (layers.drawItems && map && project && project.mapData) {
        layers.drawItems.clearLayers();
        try {
            L.geoJSON(project.mapData, {
                pointToLayer: (feature, latlng) => {
                    const color = feature.properties.markerColor || '#3b82f6';
                    if (feature.properties.isNumbered) {
                        return L.marker(latlng, {
                            draggable: true, markerColor: color,
                            markerNumber: feature.properties.markerNumber, isNumbered: true,
                            icon: L.divIcon({
                                html: `<div style="background:${color};color:#000;width:28px;height:28px;border-radius:50%;display:flex;align-items:center;justify-content:center;border:2px solid rgba(0,0,0,0.3);font-weight:bold;font-size:14px;box-shadow:0 2px 6px rgba(0,0,0,0.3);">${feature.properties.markerNumber}</div>`,
                                iconSize: [28, 28], iconAnchor: [14, 14], className: 'numbered-marker'
                            })
                        });
                    }
                    return createPinMarker(latlng, color);
                },
                style: (feature) => ({ color: feature.properties.markerColor || '#3b82f6', weight: 6 }),
                onEachFeature: (feature, layer) => {
                    layer.options.markerColor = feature.properties.markerColor;
                    if (feature.properties.isNumbered) {
                        layer.options.isNumbered = true;
                        layer.options.markerNumber = feature.properties.markerNumber;
                    }
                    setupInteractiveLayer(layer);
                    layers.drawItems.addLayer(layer);
                }
            });
        } catch (e) { console.error('Map restore error:', e); }
    }
}


// ─── SNIPPING / SCREENSHOT ───
function startSnip(type) {
    stopDrawing(true);
    if (!map) return;
    captureMode = true;
    const opts = { repeatMode: false };
    if (type === 'rectangle') AppState.activeDrawTool = new L.Draw.Rectangle(map, opts);
    else if (type === 'circle') AppState.activeDrawTool = new L.Draw.Circle(map, opts);
    else AppState.activeDrawTool = new L.Draw.Polygon(map, opts);
    AppState.activeDrawTool.enable();
}

async function triggerAusschnitt(targetLayer) {
    if (!targetLayer) return;

    const loadingEl = $('loadingOverlay');
    if (loadingEl) loadingEl.style.display = 'flex';

    // Populate row selector
    const rowSel = $('snipTargetRow');
    if (rowSel) {
        rowSel.innerHTML = '<option value="">-- Zeile wählen --</option>';
        AppState.data.forEach((r, i) => {
            const opt = document.createElement('option');
            opt.value = i;
            opt.textContent = `Zeile ${i + 1}: ${r.Kennzeichen || '---'}`;
            rowSel.appendChild(opt);
        });
    }

    try {
        const mapEl = $('map');
        const controls = document.querySelectorAll('.leaflet-control-container, .gps-bar, .map-sidebar, .map-compass');
        controls.forEach(c => c.style.display = 'none');

        await new Promise(r => setTimeout(r, 600));

        const canvas = await html2canvas(mapEl, {
            useCORS: true, allowTaint: true,
            scale: Math.min(window.devicePixelRatio || 1, 2),
            logging: false, backgroundColor: '#1e293b'
        });

        controls.forEach(c => c.style.display = '');

        const cropped = cropToShape(canvas, targetLayer, mapEl);
        const editorCanvas = $('snipEditorCanvas');
        editorCanvas.originalImage = cropped;
        editorCanvas.width = cropped.width;
        editorCanvas.height = cropped.height;
        snipTexts = [];
        drawEditorCanvas();

        if (loadingEl) loadingEl.style.display = 'none';
        $('snipConfirmModal').style.display = 'flex';
    } catch (err) {
        const controls = document.querySelectorAll('.leaflet-control-container, .gps-bar, .map-sidebar, .map-compass');
        controls.forEach(c => c.style.display = '');
        if (loadingEl) loadingEl.style.display = 'none';
        showToast('Screenshot-Fehler: ' + err.message);
        console.error(err);
    }
}

function cropToShape(fullCanvas, layer, mapEl) {
    const temp = document.createElement('canvas');
    const ctx = temp.getContext('2d');
    const sx = fullCanvas.width / mapEl.offsetWidth;
    const sy = fullCanvas.height / mapEl.offsetHeight;
    const b = layer.getBounds ? layer.getBounds() : L.latLngBounds(layer.getLatLng(), layer.getLatLng());
    const nw = map.latLngToContainerPoint(b.getNorthWest());
    const se = map.latLngToContainerPoint(b.getSouthEast());
    const w = Math.max(10, se.x - nw.x);
    const h = Math.max(10, se.y - nw.y);
    temp.width = w * sx;
    temp.height = h * sy;
    ctx.drawImage(fullCanvas, nw.x * sx, nw.y * sy, temp.width, temp.height, 0, 0, temp.width, temp.height);
    return temp;
}

function addTextToSnip() {
    const t = prompt('Text eingeben:');
    if (!t) return;
    const colorInput = $('snipTextColor');
    snipTexts.push({ text: t, x: 100, y: 100, color: colorInput ? colorInput.value : '#ffffff', fontSize: 32 });
    drawEditorCanvas();
}

function drawEditorCanvas() {
    const canvas = $('snipEditorCanvas');
    if (!canvas || !canvas.originalImage) return;
    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(canvas.originalImage, 0, 0);
    snipTexts.forEach(t => {
        ctx.font = `bold ${t.fontSize}px Inter, sans-serif`;
        ctx.shadowColor = 'rgba(0,0,0,0.8)';
        ctx.shadowBlur = 4;
        ctx.fillStyle = t.color;
        ctx.fillText(t.text, t.x, t.y);
        ctx.shadowBlur = 0;
    });
}

function initEditorEvents() {
    const canvas = $('snipEditorCanvas');
    if (!canvas) return;

    on('btnSnipZoomIn', 'click', () => {
        const c = $('snipEditorCanvas');
        if (c) c.style.width = (c.offsetWidth + 80) + 'px';
    });
    on('btnSnipZoomOut', 'click', () => {
        const c = $('snipEditorCanvas');
        if (c) c.style.width = Math.max(150, c.offsetWidth - 80) + 'px';
    });

    canvas.onmousedown = (e) => {
        const p = getCanvasPos(e, canvas);
        draggingTextIdx = snipTexts.findIndex(t => p.x > t.x && p.x < t.x + 200 && p.y > t.y - 40 && p.y < t.y);
    };
    canvas.onmousemove = (e) => {
        if (draggingTextIdx === -1) return;
        const p = getCanvasPos(e, canvas);
        snipTexts[draggingTextIdx].x = p.x;
        snipTexts[draggingTextIdx].y = p.y;
        drawEditorCanvas();
    };
    canvas.onmouseup = () => { draggingTextIdx = -1; };
}

function getCanvasPos(e, canvas) {
    const r = canvas.getBoundingClientRect();
    return { x: (e.clientX - r.left) * (canvas.width / r.width), y: (e.clientY - r.top) * (canvas.height / r.height) };
}

function updateSnipColor() {
    const val = ($('snipTargetType') || {}).value;
    const indicator = $('snipColorIndicator');
    if (!indicator) return;
    if (val === 'Anhang_0.8') indicator.style.background = 'var(--depth-08)';
    else if (val === 'Anhang_1.6') indicator.style.background = 'var(--depth-16)';
    else if (val === 'Anhang_3.2') indicator.style.background = 'var(--depth-32)';
    else indicator.style.background = 'var(--text-primary)';
}

function saveFinalSnip() {
    const canvas = $('snipEditorCanvas');
    const targetIdx = parseInt(($('snipTargetRow') || {}).value);
    const targetCol = ($('snipTargetType') || {}).value;

    if (isNaN(targetIdx) || !targetCol) {
        showToast('Bitte zuerst eine Zeile auswählen!');
        return;
    }
    if (!AppState.data[targetIdx]) {
        showToast('Zeile existiert nicht');
        return;
    }

    AppState.data[targetIdx][targetCol] = canvas.toDataURL();
    renderTable();
    saveToStorage();
    $('snipConfirmModal').style.display = 'none';
    showToast('Foto gespeichert in Zeile ' + (targetIdx + 1));
}

function downloadSnip() {
    const canvas = $('snipEditorCanvas');
    if (!canvas) return;
    const link = document.createElement('a');
    link.download = `Messstelle_${Date.now()}.png`;
    link.href = canvas.toDataURL('image/png');
    link.click();
    showToast('Heruntergeladen');
}

function openImagePreview(url) {
    const overlay = document.createElement('div');
    overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.9);z-index:999999;display:flex;align-items:center;justify-content:center;cursor:pointer;';
    overlay.onclick = () => document.body.removeChild(overlay);
    const img = new Image();
    img.src = url;
    img.style.cssText = 'max-width:90%;max-height:90%;border-radius:8px;box-shadow:0 8px 32px rgba(0,0,0,0.5);';
    overlay.appendChild(img);
    document.body.appendChild(overlay);
}

// ─── COLUMN MANAGER ───
function openColManager() {
    const modal = $('colManagerModal');
    const grid = $('colManagerGrid');
    if (!modal || !grid) return;

    grid.innerHTML = '';
    TABLE_STRUCTURE.forEach(g => {
        const isVisible = !AppState.hiddenColumns.has(g.class);
        const cs = DEPTH_COLORS[g.class] || DEPTH_COLORS.basis;
        const item = document.createElement('div');
        item.style.cssText = `display:flex;align-items:center;justify-content:space-between;padding:12px 16px;background:var(--bg-primary);border:1px solid ${isVisible ? cs.border : 'var(--border)'};border-radius:var(--radius-sm);cursor:pointer;transition:all 0.15s;`;
        item.innerHTML = `
            <span style="font-size:12px;font-weight:600;color:${isVisible ? cs.text : 'var(--text-muted)'}">${g.group}</span>
            <i class="fas ${isVisible ? 'fa-eye' : 'fa-eye-slash'}" style="color:${isVisible ? cs.text : 'var(--text-muted)'}"></i>
        `;
        item.onclick = () => {
            if (AppState.hiddenColumns.has(g.class)) AppState.hiddenColumns.delete(g.class);
            else AppState.hiddenColumns.add(g.class);
            renderTable();
            saveToStorage();
            openColManager();
        };
        grid.appendChild(item);
    });
    modal.style.display = 'flex';
}

// ─── PROJECTS ───
function renderProjectList() {
    const grid = $('standortGrid');
    if (!grid) return;
    const library = JSON.parse(localStorage.getItem('messstellen_library') || '{}');

    if (Object.keys(library).length === 0) {
        grid.innerHTML = '<p style="color:var(--text-muted);text-align:center;padding:24px;">Keine Projekte vorhanden</p>';
        return;
    }

    grid.innerHTML = '';
    Object.values(library).sort((a, b) => b.timestamp - a.timestamp).forEach(p => {
        const item = document.createElement('div');
        item.style.cssText = 'display:flex;align-items:center;justify-content:space-between;padding:14px 16px;background:var(--bg-primary);border:1px solid var(--border);border-radius:var(--radius-sm);cursor:pointer;transition:all 0.15s;';
        item.innerHTML = `
            <div>
                <div style="font-weight:600;font-size:13px;color:var(--accent);">${p.name}</div>
                <div style="font-size:11px;color:var(--text-muted);">${(p.data || []).length} Zeilen &middot; ${new Date(p.timestamp).toLocaleString('de-DE')}</div>
            </div>
            <button class="btn-icon" style="color:var(--danger);" data-delete="${p.name}"><i class="fas fa-trash"></i></button>
        `;
        item.onclick = (e) => {
            if (e.target.closest('[data-delete]')) {
                const name = e.target.closest('[data-delete]').dataset.delete;
                if (confirm(`"${name}" löschen?`)) {
                    delete library[name];
                    localStorage.setItem('messstellen_library', JSON.stringify(library));
                    renderProjectList();
                }
                return;
            }
            loadProject(p.name);
            $('standortModal').style.display = 'none';
        };
        grid.appendChild(item);
    });
}

function loadProject(name) {
    let library = JSON.parse(localStorage.getItem('messstellen_library') || '{}');
    const p = library[name];
    if (p) {
        localStorage.setItem('current_project_id', name);
        AppState.data = p.data || [];
        AppState.hiddenMapColors = new Set(p.hiddenMapColors || []);
        AppState.newCols = new Set(p.newCols || []);
        const nameInput = $('projectName');
        if (nameInput) nameInput.value = p.name;
        renderTable();
        if (map) restoreMapDrawings();
        showToast(`Projekt geladen: ${p.name}`);
    }
}

// ─── EXPORT ───
function openExportModal() {
    if (AppState.data.length === 0) { showToast('Keine Daten zum Exportieren'); return; }
    const chk08 = $('exportChk08');
    const chk16 = $('exportChk16');
    const chk32 = $('exportChk32');
    if (chk08) chk08.checked = !AppState.hiddenColumns.has('08');
    if (chk16) chk16.checked = !AppState.hiddenColumns.has('16');
    if (chk32) chk32.checked = !AppState.hiddenColumns.has('32');
    $('exportModal').style.display = 'flex';
}

window.confirmExportDepths = function() {
    const opts = {
        export08: $('exportChk08') && $('exportChk08').checked,
        export16: $('exportChk16') && $('exportChk16').checked,
        export32: $('exportChk32') && $('exportChk32').checked
    };
    if (!opts.export08 && !opts.export16 && !opts.export32) { showToast('Mindestens eine Tiefe auswählen'); return; }
    $('exportModal').style.display = 'none';
    exportExcel(opts);
};

async function exportExcel(opts) {
    const loadingEl = $('loadingOverlay');
    if (loadingEl) loadingEl.style.display = 'flex';

    try {
        const workbook = new ExcelJS.Workbook();
        const projectName = ($('projectName') || {}).value || 'Export';

        const headerFont = { name: 'Inter', bold: true, size: 13 };
        const titleFont = { name: 'Inter', bold: true, size: 18 };
        const dataFont = { name: 'Inter', size: 12 };
        const center = { horizontal: 'center', vertical: 'middle', wrapText: true };
        const border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

        const depthsToExport = [];
        if (opts.export08) depthsToExport.push('08');
        if (opts.export16) depthsToExport.push('16');
        if (opts.export32) depthsToExport.push('32');
        const depthLabels = { '08': '0.80 m', '16': '1.60 m', '32': '3.20 m' };
        const depthArgbColors = { '08': 'FFF5D0FE', '16': 'FFFFFBEB', '32': 'FFCFFAFE' };
        const customCols = Array.from(AppState.newCols);

        // Load logo
        let logoId = null;
        try {
            const logoResp = await fetch('./idb_logo.jpg');
            const logoBlob = await logoResp.blob();
            const logoBase64 = await new Promise((resolve) => {
                const reader = new FileReader();
                reader.onloadend = () => resolve(reader.result.split(',')[1]);
                reader.readAsDataURL(logoBlob);
            });
            logoId = workbook.addImage({ base64: logoBase64, extension: 'jpeg' });
        } catch (e) { console.warn('Logo not loaded:', e); }

        // ======= WORKSHEET 1: BODENWIDERSTAND =======
        const ws = workbook.addWorksheet('Bodenwiderstand');
        const totalBaseCols = 6 + customCols.length;
        const totalDepthCols = depthsToExport.length * 9;
        const totalCols = totalBaseCols + totalDepthCols + 1;

        // Title (rows 1-6)
        ws.mergeCells(3, 1, 4, totalCols);
        const titleCell = ws.getCell('A3');
        titleCell.value = 'Auswertung Bodenwiderstandsmessung - ' + projectName;
        titleCell.font = titleFont;
        titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        for (let r = 1; r <= 6; r++) ws.getRow(r).height = 18;
        ws.getRow(3).height = 28;

        if (logoId !== null) {
            ws.addImage(logoId, { tl: { col: 0, row: 0 }, ext: { width: 280, height: 80 } });
        }

        // Headers at row 7-8
        const HEADER_ROW = 7;
        const DATA_START = 9;

        const h1 = ['Kennzeichen', 'Alt-Kz.', 'Typ', 'Örtlichkeit', 'Meter [m]', 'Datum', ...customCols];
        const h2 = Array(h1.length).fill('');
        depthsToExport.forEach(d => {
            h1.push(depthLabels[d], '', '', '', '', '', '', '', '');
            h2.push('R1 [Ω]', 'R2 [Ω]', 'R3 [Ω]', 'ρ1 [Ωm]', 'ρ2 [Ωm]', 'ρ3 [Ωm]', 'Mittelwert [Ωm]', 'Std.Abw. [Ωm]', 'Anhang');
        });
        h1.push('Anhang (Gesamt)');
        h2.push('');

        ws.getRow(HEADER_ROW).values = h1;
        ws.getRow(HEADER_ROW + 1).values = h2;
        ws.getRow(HEADER_ROW + 1).height = 30;

        [ws.getRow(HEADER_ROW), ws.getRow(HEADER_ROW + 1)].forEach(row => {
            row.eachCell((cell, colNum) => {
                cell.font = headerFont;
                cell.alignment = center;
                cell.border = border;
                let fillColor = 'FFE2E8F0';
                if (colNum > totalBaseCols && colNum <= h1.length - 1) {
                    const depthIdx = Math.floor((colNum - totalBaseCols - 1) / 9);
                    if (depthIdx < depthsToExport.length) fillColor = depthArgbColors[depthsToExport[depthIdx]];
                }
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
            });
        });

        for (let i = 1; i <= totalBaseCols; i++) ws.mergeCells(HEADER_ROW, i, HEADER_ROW + 1, i);
        ws.mergeCells(HEADER_ROW, h1.length, HEADER_ROW + 1, h1.length);
        let mc = totalBaseCols + 1;
        depthsToExport.forEach(() => { ws.mergeCells(HEADER_ROW, mc, HEADER_ROW, mc + 8); mc += 9; });

        // Data rows - only export rows that have at least Kennzeichen filled
        const exportData = AppState.data.filter(row => {
            return (row['Kennzeichen'] && row['Kennzeichen'].trim() !== '') ||
                   depthsToExport.some(depth => {
                       const sfx = depth === '08' ? '0.8' : (depth === '16' ? '1.6' : '3.2');
                       return parseFloat(row['R1 [Ω]_' + sfx]) > 0 || parseFloat(row['MW [Ωm]_' + sfx]) > 0;
                   });
        });
        for (let i = 0; i < exportData.length; i++) {
            const d = exportData[i];
            const exRow = ws.getRow(DATA_START + i);
            const vals = [d['Kennzeichen'] || '', d['Alt-Kz.'] || '', d['Typ'] || '', d['Örtlichkeit'] || '', d['Meter [m]'] || '', d['Datum'] || ''];
            customCols.forEach(c => vals.push(d[c] || ''));
            depthsToExport.forEach(depth => {
                const sfx = depth === '08' ? '0.8' : (depth === '16' ? '1.6' : '3.2');
                const get = (p) => d[p + ' [Ω]_' + sfx] || d[p + ' [Ωm]_' + sfx] || '';
                vals.push(get('R1'), get('R2'), get('R3'), get('ρ1'), get('ρ2'), get('ρ3'), get('MW'), get('SD'), '');
            });
            vals.push('');
            exRow.values = vals;
            exRow.eachCell(cell => { cell.font = dataFont; cell.alignment = center; cell.border = border; });

            depthsToExport.forEach((depth, di) => {
                const sfx = depth === '08' ? '0.8' : (depth === '16' ? '1.6' : '3.2');
                const imgData = d['Anhang_' + sfx];
                if (imgData && imgData.startsWith('data:image')) {
                    const imgId2 = workbook.addImage({ base64: imgData, extension: 'png' });
                    const anhangCol = totalBaseCols + di * 9 + 8;
                    // Image fits exactly in one cell: col width 32 chars (~230px), row height 170pt (~230px)
                    ws.addImage(imgId2, {
                        tl: { col: anhangCol, row: DATA_START - 1 + i },
                        br: { col: anhangCol + 1, row: DATA_START + i }
                    });
                    exRow.height = 170;
                    ws.getColumn(anhangCol + 1).width = 32;
                }
            });
            const gesImg = d['Anhang_Global'];
            if (gesImg && gesImg.startsWith('data:image')) {
                const imgId2 = workbook.addImage({ base64: gesImg, extension: 'png' });
                ws.addImage(imgId2, {
                    tl: { col: h1.length - 1, row: DATA_START - 1 + i },
                    br: { col: h1.length, row: DATA_START + i }
                });
                exRow.height = 170;
                ws.getColumn(h1.length).width = 32;
            }
        }

        // Auto column widths - fit to content, skip Anhang columns
        const anhangColIndices = new Set();
        depthsToExport.forEach((depth, di) => { anhangColIndices.add(totalBaseCols + di * 9 + 8); });
        anhangColIndices.add(h1.length - 1); // Anhang Global

        for (let colIdx = 0; colIdx < totalCols; colIdx++) {
            if (anhangColIndices.has(colIdx)) continue;
            const col = ws.getColumn(colIdx + 1);
            let maxLen = 8;
            // Check header text
            const hVal = String(h1[colIdx] || h2[colIdx] || '');
            maxLen = Math.max(maxLen, hVal.length * 1.2 + 2);
            // Check all data rows
            for (let r = 0; r < exportData.length; r++) {
                const cell = ws.getRow(DATA_START + r).getCell(colIdx + 1);
                const val = cell.value ? String(cell.value) : '';
                maxLen = Math.max(maxLen, val.length * 1.1 + 2);
            }
            col.width = Math.max(10, Math.min(Math.ceil(maxLen), 35));
        }

        // ======= WORKSHEET 2: STATISTIK-WEDAL =======
        const statSheet = workbook.addWorksheet('Statistik-' + projectName);
        const statTotalCols = 1 + depthsToExport.length * 3;

        // Title
        statSheet.mergeCells(3, 1, 4, statTotalCols);
        const statTitle = statSheet.getCell('A3');
        statTitle.value = 'Statistik Bodenwiderstandsmessung - ' + projectName;
        statTitle.font = titleFont;
        statTitle.alignment = { horizontal: 'center', vertical: 'middle' };
        for (let r = 1; r <= 6; r++) statSheet.getRow(r).height = 18;
        statSheet.getRow(3).height = 28;

        if (logoId !== null) {
            statSheet.addImage(logoId, { tl: { col: 0, row: 0 }, ext: { width: 280, height: 80 } });
        }

        // Headers at row 7-9
        const SH_ROW = 7;
        const sH1 = ['Kennzeichen'];
        const sH2 = [''];
        const sH3 = [''];
        depthsToExport.forEach(d => {
            sH1.push(depthLabels[d], '', '');
            sH2.push('Einzelwerte', 'Mittelwert', 'Standard Abweichung');
            sH3.push('ρ [Ωm]', '', '');
        });
        statSheet.getRow(SH_ROW).values = sH1;
        statSheet.getRow(SH_ROW + 1).values = sH2;
        statSheet.getRow(SH_ROW + 2).values = sH3;

        [statSheet.getRow(SH_ROW), statSheet.getRow(SH_ROW + 1), statSheet.getRow(SH_ROW + 2)].forEach(row => {
            row.eachCell(cell => {
                cell.font = { name: 'Inter', bold: true, size: 14 }; cell.alignment = center; cell.border = border;
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
            });
            row.height = 24;
        });

        statSheet.mergeCells(SH_ROW, 1, SH_ROW + 2, 1);
        let sMC = 2;
        depthsToExport.forEach(() => { statSheet.mergeCells(SH_ROW, sMC, SH_ROW, sMC + 2); sMC += 3; });

        // Data - only rows with actual values
        const statData = AppState.data.filter(row => {
            return depthsToExport.some(d => {
                const sfx = d === '08' ? '0.8' : (d === '16' ? '1.6' : '3.2');
                return parseFloat(row['MW [Ωm]_' + sfx]) > 0;
            });
        });
        let curRow = SH_ROW + 3;
        for (let i = 0; i < statData.length; i++) {
            const dataRow = statData[i];
            for (let r = 1; r <= 3; r++) {
                const exRow = statSheet.getRow(curRow + r - 1);
                const values = [r === 1 ? (dataRow['Kennzeichen'] || '') : ''];
                depthsToExport.forEach(d => {
                    const sfx = d === '08' ? '0.8' : (d === '16' ? '1.6' : '3.2');
                    values.push(dataRow['ρ' + r + ' [Ωm]_' + sfx] || '');
                    values.push(r === 1 ? (dataRow['MW [Ωm]_' + sfx] || '') : '');
                    values.push(r === 1 ? (dataRow['SD [Ωm]_' + sfx] || '') : '');
                });
                exRow.values = values;
                exRow.eachCell(cell => { cell.font = { name: 'Inter', size: 13 }; cell.alignment = center; cell.border = border; });
            }
            statSheet.mergeCells(curRow, 1, curRow + 2, 1);
            let mCol = 3;
            depthsToExport.forEach(() => {
                statSheet.mergeCells(curRow, mCol, curRow + 2, mCol);
                statSheet.mergeCells(curRow, mCol + 1, curRow + 2, mCol + 1);
                mCol += 3;
            });
            curRow += 3;
        }

        // Statistik column widths
        const statWidths = [{ width: 22 }];
        depthsToExport.forEach(() => { statWidths.push({ width: 16 }, { width: 18 }, { width: 24 }); });
        statSheet.columns = statWidths;

        // ======= CHARTS FOR STATISTIK =======
        // Only include rows that have actual MW data
        const filledData = AppState.data.filter(row => {
            return depthsToExport.some(d => {
                const sfx = d === '08' ? '0.8' : (d === '16' ? '1.6' : '3.2');
                return parseFloat(row['MW [Ωm]_' + sfx]) > 0;
            });
        });

        if (filledData.length > 0) {
            const depthBarColors = { '08': ['#e879f9', '#a21caf'], '16': ['#fbbf24', '#b45309'], '32': ['#22d3ee', '#0e7490'] };
            const trendColors = { '08': '#a21caf', '16': '#b45309', '32': '#0e7490' };

            // --- BALKENDIAGRAMM (Bar Chart with Error Bars + Trend Lines) ---
            const barCanvas = document.createElement('canvas');
            barCanvas.width = 1600;
            barCanvas.height = 800;
            const bCtx = barCanvas.getContext('2d');

            // Background
            bCtx.fillStyle = '#ffffff';
            bCtx.fillRect(0, 0, 1600, 800);

            const bPad = { top: 70, bottom: 90, left: 90, right: 50 };
            const bW = 1600 - bPad.left - bPad.right;
            const bH = 800 - bPad.top - bPad.bottom;

            // Max value
            let bMax = 0;
            filledData.forEach(row => {
                depthsToExport.forEach(d => {
                    const sfx = d === '08' ? '0.8' : (d === '16' ? '1.6' : '3.2');
                    const mw = parseFloat(row['MW [Ωm]_' + sfx]) || 0;
                    const sd = parseFloat(row['SD [Ωm]_' + sfx]) || 0;
                    bMax = Math.max(bMax, mw + sd * 1.5);
                });
            });
            bMax = Math.ceil(bMax * 1.2 / 50) * 50;
            if (bMax === 0) bMax = 100;

            // Grid
            bCtx.strokeStyle = '#e2e8f0';
            bCtx.lineWidth = 1;
            for (let i = 0; i <= 5; i++) {
                const y = bPad.top + bH - (bH * i / 5);
                bCtx.beginPath(); bCtx.moveTo(bPad.left, y); bCtx.lineTo(bPad.left + bW, y); bCtx.stroke();
                bCtx.fillStyle = '#475569';
                bCtx.font = '14px Inter, sans-serif';
                bCtx.textAlign = 'right';
                bCtx.fillText(Math.round(bMax * i / 5).toLocaleString('de-DE'), bPad.left - 12, y + 4);
            }
            // Baseline
            bCtx.strokeStyle = '#94a3b8';
            bCtx.lineWidth = 2;
            bCtx.beginPath(); bCtx.moveTo(bPad.left, bPad.top + bH); bCtx.lineTo(bPad.left + bW, bPad.top + bH); bCtx.stroke();

            const groupW = bW / filledData.length;
            const barWidth = (groupW * 0.65) / depthsToExport.length;

            // Draw bars
            filledData.forEach((row, i) => {
                const gX = bPad.left + i * groupW + groupW * 0.175;

                depthsToExport.forEach((d, di) => {
                    const sfx = d === '08' ? '0.8' : (d === '16' ? '1.6' : '3.2');
                    const mw = parseFloat(row['MW [Ωm]_' + sfx]) || 0;
                    const sd = parseFloat(row['SD [Ωm]_' + sfx]) || 0;
                    if (mw <= 0) return;

                    const barH = (mw / bMax) * bH;
                    const bx = gX + di * barWidth;
                    const by = bPad.top + bH - barH;

                    // Hatched cylinder bar (matching reference image)
                    bCtx.fillStyle = depthBarColors[d][0] + '30';
                    bCtx.fillRect(bx, by, barWidth - 6, barH);
                    // Diagonal hatching
                    bCtx.save();
                    bCtx.beginPath(); bCtx.rect(bx, by, barWidth - 6, barH); bCtx.clip();
                    bCtx.strokeStyle = depthBarColors[d][0] + '70'; bCtx.lineWidth = 1.5;
                    for (let hi = -barH; hi < barWidth + barH; hi += 8) {
                        bCtx.beginPath(); bCtx.moveTo(bx + hi, by + barH); bCtx.lineTo(bx + hi + barH, by); bCtx.stroke();
                    }
                    bCtx.restore();
                    // Border
                    bCtx.strokeStyle = depthBarColors[d][0]; bCtx.lineWidth = 2;
                    bCtx.strokeRect(bx, by, barWidth - 6, barH);
                    // Top ellipse (hollow)
                    const ecx = bx + (barWidth - 6) / 2;
                    bCtx.fillStyle = '#ffffff';
                    bCtx.beginPath(); bCtx.ellipse(ecx, by, (barWidth - 6) / 2, 8, 0, 0, Math.PI * 2); bCtx.fill();
                    bCtx.strokeStyle = depthBarColors[d][0]; bCtx.lineWidth = 2;
                    bCtx.beginPath(); bCtx.ellipse(ecx, by, (barWidth - 6) / 2, 8, 0, 0, Math.PI * 2); bCtx.stroke();
                    bCtx.strokeStyle = depthBarColors[d][0] + '66'; bCtx.lineWidth = 1;
                    bCtx.beginPath(); bCtx.ellipse(ecx, by, (barWidth - 6) / 2 - 4, 4, 0, 0, Math.PI * 2); bCtx.stroke();

                    // Error bar
                    if (sd > 0) {
                        const sdH = (sd / bMax) * bH;
                        const cx = bx + (barWidth - 6) / 2;
                        bCtx.strokeStyle = '#1f2937';
                        bCtx.lineWidth = 2;
                        bCtx.beginPath(); bCtx.moveTo(cx, by - sdH); bCtx.lineTo(cx, by + Math.min(sdH, barH)); bCtx.stroke();
                        bCtx.beginPath(); bCtx.moveTo(cx - 8, by - sdH); bCtx.lineTo(cx + 8, by - sdH); bCtx.stroke();
                        bCtx.beginPath(); bCtx.moveTo(cx - 8, by + Math.min(sdH, barH)); bCtx.lineTo(cx + 8, by + Math.min(sdH, barH)); bCtx.stroke();
                    }

                    // Value label
                    bCtx.fillStyle = depthBarColors[d][1];
                    bCtx.font = 'bold 18px Inter, sans-serif';
                    bCtx.textAlign = 'center';
                    const labelY = sd > 0 ? by - (sd / bMax) * bH - 20 : by - 14;
                    bCtx.fillText(mw.toFixed(1), bx + (barWidth - 6) / 2, labelY);
                    if (sd > 0) {
                        bCtx.fillStyle = depthBarColors[d][1];
                        bCtx.font = '14px Inter, sans-serif';
                        bCtx.fillText('\u00B1 ' + sd.toFixed(1), bx + (barWidth - 6) / 2, labelY + 13);
                    }
                });

                // X label
                bCtx.fillStyle = '#1f2937';
                bCtx.font = 'bold 16px Inter, sans-serif';
                bCtx.textAlign = 'center';
                bCtx.fillText(row['Kennzeichen'] || (i + 1).toString(), bPad.left + i * groupW + groupW / 2, bPad.top + bH + 25);
            });

            // Trend lines per depth (dashed)
            depthsToExport.forEach((d, di) => {
                bCtx.strokeStyle = trendColors[d];
                bCtx.lineWidth = 2.5;
                bCtx.setLineDash([10, 6]);
                bCtx.beginPath();
                let started = false;
                filledData.forEach((row, i) => {
                    const sfx = d === '08' ? '0.8' : (d === '16' ? '1.6' : '3.2');
                    const mw = parseFloat(row['MW [Ωm]_' + sfx]) || 0;
                    if (mw <= 0) return;
                    const gX = bPad.left + i * groupW + groupW * 0.175;
                    const cx = gX + di * barWidth + (barWidth - 6) / 2;
                    const cy = bPad.top + bH - (mw / bMax) * bH;
                    if (!started) { bCtx.moveTo(cx, cy); started = true; }
                    else bCtx.lineTo(cx, cy);
                });
                bCtx.stroke();
                bCtx.setLineDash([]);

                // Dots on trend
                filledData.forEach((row, i) => {
                    const sfx = d === '08' ? '0.8' : (d === '16' ? '1.6' : '3.2');
                    const mw = parseFloat(row['MW [Ωm]_' + sfx]) || 0;
                    if (mw <= 0) return;
                    const gX = bPad.left + i * groupW + groupW * 0.175;
                    const cx = gX + di * barWidth + (barWidth - 6) / 2;
                    const cy = bPad.top + bH - (mw / bMax) * bH;
                    bCtx.fillStyle = trendColors[d];
                    bCtx.beginPath(); bCtx.arc(cx, cy, 5, 0, Math.PI * 2); bCtx.fill();
                });
            });

            // Title & labels
            bCtx.fillStyle = '#0f172a';
            bCtx.font = 'bold 22px Inter, sans-serif';
            bCtx.textAlign = 'center';
            bCtx.fillText('Bodenwiderstand - ' + projectName, 800, 40);
            bCtx.save(); bCtx.translate(25, bPad.top + bH / 2); bCtx.rotate(-Math.PI / 2);
            bCtx.fillStyle = '#374151'; bCtx.font = 'bold 16px Inter, sans-serif'; bCtx.fillText('Bodenwiderstand [\u03A9m]', 0, 0);
            bCtx.restore();
            bCtx.fillStyle = '#374151'; bCtx.font = 'bold 15px Inter, sans-serif'; bCtx.textAlign = 'center';
            bCtx.fillText('Kennzeichen [km]', bPad.left + bW / 2, 800 - 25);

            // Legend
            let lx = bPad.left;
            depthsToExport.forEach((d, di) => {
                bCtx.fillStyle = depthBarColors[d][0]; bCtx.fillRect(lx, 800 - 60, 16, 16);
                bCtx.strokeStyle = depthBarColors[d][1]; bCtx.strokeRect(lx, 800 - 60, 16, 16);
                bCtx.fillStyle = '#374151'; bCtx.font = '14px Inter, sans-serif'; bCtx.textAlign = 'left';
                bCtx.fillText(depthLabels[d] + ' (Balken)', lx + 22, 800 - 48);
                bCtx.strokeStyle = trendColors[d]; bCtx.lineWidth = 2.5; bCtx.setLineDash([6, 4]);
                bCtx.beginPath(); bCtx.moveTo(lx + 130, 800 - 52); bCtx.lineTo(lx + 160, 800 - 52); bCtx.stroke();
                bCtx.setLineDash([]);
                bCtx.fillText('Trend', lx + 165, 800 - 48);
                lx += 220;
            });

            // Add bar chart image
            const barImgId = workbook.addImage({ base64: barCanvas.toDataURL('image/png'), extension: 'png' });
            statSheet.addImage(barImgId, { tl: { col: 0, row: curRow + 2 }, ext: { width: 1100, height: 550 } });

            // --- SCATTER PLOT (Scatter with connected dots per depth) ---
            const scCanvas = document.createElement('canvas');
            scCanvas.width = 1600;
            scCanvas.height = 800;
            const sCtx = scCanvas.getContext('2d');

            sCtx.fillStyle = '#ffffff';
            sCtx.fillRect(0, 0, 1600, 800);

            // Grid
            sCtx.strokeStyle = '#e2e8f0'; sCtx.lineWidth = 1;
            for (let i = 0; i <= 5; i++) {
                const y = bPad.top + bH - (bH * i / 5);
                sCtx.beginPath(); sCtx.moveTo(bPad.left, y); sCtx.lineTo(bPad.left + bW, y); sCtx.stroke();
                sCtx.fillStyle = '#475569'; sCtx.font = '14px Inter, sans-serif'; sCtx.textAlign = 'right';
                sCtx.fillText(Math.round(bMax * i / 5).toLocaleString('de-DE'), bPad.left - 12, y + 4);
            }
            sCtx.strokeStyle = '#94a3b8'; sCtx.lineWidth = 2;
            sCtx.beginPath(); sCtx.moveTo(bPad.left, bPad.top + bH); sCtx.lineTo(bPad.left + bW, bPad.top + bH); sCtx.stroke();

            // Scatter lines + dots per depth (different markers + STD)
            const scLineStyles = [[], [14, 8], [6, 6]];
            depthsToExport.forEach((d, di) => {
                const sfx = d === '08' ? '0.8' : (d === '16' ? '1.6' : '3.2');
                const color = trendColors[d];

                // Line with different dash per depth
                sCtx.strokeStyle = color; sCtx.lineWidth = 3; sCtx.setLineDash(scLineStyles[di] || []);
                sCtx.beginPath();
                let started = false;
                filledData.forEach((row, i) => {
                    const mw = parseFloat(row['MW [Ωm]_' + sfx]) || 0;
                    if (mw <= 0) return;
                    const cx = bPad.left + i * groupW + groupW / 2;
                    const cy = bPad.top + bH - (mw / bMax) * bH;
                    if (!started) { sCtx.moveTo(cx, cy); started = true; }
                    else sCtx.lineTo(cx, cy);
                });
                sCtx.stroke();
                sCtx.setLineDash([]);

                // Markers + STD error bars
                filledData.forEach((row, i) => {
                    const mw = parseFloat(row['MW [Ωm]_' + sfx]) || 0;
                    const sd = parseFloat(row['SD [Ωm]_' + sfx]) || 0;
                    if (mw <= 0) return;
                    const cx = bPad.left + i * groupW + groupW / 2;
                    const cy = bPad.top + bH - (mw / bMax) * bH;

                    // STD error bars
                    if (sd > 0) {
                        const sdPx = (sd / bMax) * bH;
                        sCtx.strokeStyle = color; sCtx.lineWidth = 2;
                        sCtx.beginPath(); sCtx.moveTo(cx, cy - sdPx); sCtx.lineTo(cx, cy + sdPx); sCtx.stroke();
                        sCtx.beginPath(); sCtx.moveTo(cx - 7, cy - sdPx); sCtx.lineTo(cx + 7, cy - sdPx); sCtx.stroke();
                        sCtx.beginPath(); sCtx.moveTo(cx - 7, cy + sdPx); sCtx.lineTo(cx + 7, cy + sdPx); sCtx.stroke();
                    }

                    // Different marker per depth
                    sCtx.fillStyle = color; sCtx.strokeStyle = '#ffffff'; sCtx.lineWidth = 2;
                    if (di === 0) {
                        sCtx.beginPath(); sCtx.arc(cx, cy, 11, 0, Math.PI * 2); sCtx.fill(); sCtx.stroke();
                    } else if (di === 1) {
                        sCtx.beginPath(); sCtx.moveTo(cx, cy - 13); sCtx.lineTo(cx + 11, cy); sCtx.lineTo(cx, cy + 13); sCtx.lineTo(cx - 11, cy); sCtx.closePath(); sCtx.fill(); sCtx.stroke();
                    } else {
                        sCtx.fillRect(cx - 10, cy - 10, 20, 20); sCtx.strokeRect(cx - 10, cy - 10, 20, 20);
                    }

                    // Value + STD
                    sCtx.fillStyle = color; sCtx.font = 'bold 16px Inter, sans-serif'; sCtx.textAlign = 'center';
                    const vlY = sd > 0 ? cy - (sd / bMax) * bH - 18 : cy - 18;
                    sCtx.fillText(mw.toFixed(1), cx, vlY);
                    if (sd > 0) {
                        sCtx.font = '13px Inter, sans-serif';
                        sCtx.fillText('\u00B1 ' + sd.toFixed(1), cx, vlY + 15);
                    }
                });
            });

            // X labels
            filledData.forEach((row, i) => {
                sCtx.fillStyle = '#1f2937'; sCtx.font = 'bold 16px Inter, sans-serif'; sCtx.textAlign = 'center';
                sCtx.fillText(row['Kennzeichen'] || (i + 1).toString(), bPad.left + i * groupW + groupW / 2, bPad.top + bH + 25);
            });

            // Title & labels
            sCtx.fillStyle = '#0f172a'; sCtx.font = 'bold 22px Inter, sans-serif'; sCtx.textAlign = 'center';
            sCtx.fillText('Bodenwiderstand - ' + projectName, 800, 40);
            sCtx.save(); sCtx.translate(25, bPad.top + bH / 2); sCtx.rotate(-Math.PI / 2);
            sCtx.fillStyle = '#374151'; sCtx.font = 'bold 16px Inter, sans-serif'; sCtx.fillText('Bodenwiderstand [\u03A9m]', 0, 0);
            sCtx.restore();
            sCtx.fillStyle = '#374151'; sCtx.font = 'bold 15px Inter, sans-serif'; sCtx.textAlign = 'center';
            sCtx.fillText('Kennzeichen [km]', bPad.left + bW / 2, 800 - 25);

            // Legend
            lx = bPad.left;
            depthsToExport.forEach((d) => {
                sCtx.fillStyle = trendColors[d]; sCtx.beginPath(); sCtx.arc(lx + 8, 800 - 52, 8, 0, Math.PI * 2); sCtx.fill();
                sCtx.strokeStyle = trendColors[d]; sCtx.lineWidth = 3;
                sCtx.beginPath(); sCtx.moveTo(lx + 20, 800 - 52); sCtx.lineTo(lx + 50, 800 - 52); sCtx.stroke();
                sCtx.fillStyle = '#374151'; sCtx.font = '14px Inter, sans-serif'; sCtx.textAlign = 'left';
                sCtx.fillText(depthLabels[d], lx + 56, 800 - 48);
                lx += 150;
            });

            // Add scatter chart below bar chart
            const scImgId = workbook.addImage({ base64: scCanvas.toDataURL('image/png'), extension: 'png' });
            const barChartRows = Math.ceil(550 / 15); // approx rows the bar chart takes
            statSheet.addImage(scImgId, { tl: { col: 0, row: curRow + 2 + barChartRows + 2 }, ext: { width: 1100, height: 550 } });
        }

        // Generate file with native clickable Excel chart
        const rawBuffer = await workbook.xlsx.writeBuffer();
        const finalBuffer = await injectNativeBarChart(rawBuffer, depthsToExport, depthLabels, statData, projectName, curRow);
        const blob = new Blob([finalBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'Bodenwiderstand_' + projectName + '_' + new Date().toISOString().split('T')[0] + '.xlsx';
        document.body.appendChild(a);
        a.click();
        URL.revokeObjectURL(url);
        a.remove();
        showToast('Excel exportiert');
    } catch (err) {
        console.error('Export error:', err);
        showToast('Export-Fehler: ' + err.message);
    }
    if (loadingEl) loadingEl.style.display = 'none';
}

// ─── NATIVE EXCEL CHART INJECTION ───
async function injectNativeBarChart(xlsxBuffer, depthsToExport, depthLabels, filledData, projectName, dataEndRow) {
    try {
        if (typeof JSZip === 'undefined') {
            console.warn('JSZip not loaded, skipping native chart');
            return xlsxBuffer;
        }
        var zip = await JSZip.loadAsync(xlsxBuffer);
        var statSheetName = 'Statistik-' + projectName;
        var SH_ROW = 7;
        var STAT_DATA_START = SH_ROW + 3;
        var numRows = filledData.length;
        if (numRows === 0) return xlsxBuffer;

        var depthHex = { '08': 'E879F9', '16': 'FBBF24', '32': '22D3EE' };
        var seriesXml = '';

        depthsToExport.forEach(function(d, di) {
            var mwCol = 2 + di * 3 + 1;
            var colLetter = String.fromCharCode(64 + mwCol);
            var color = depthHex[d] || '94A3B8';
            var catFormula = "'" + statSheetName + "'!$A$" + STAT_DATA_START + ":$A$" + (STAT_DATA_START + (numRows - 1) * 3);
            var valFormula = "'" + statSheetName + "'!$" + colLetter + "$" + STAT_DATA_START + ":$" + colLetter + "$" + (STAT_DATA_START + (numRows - 1) * 3);

            seriesXml += '<c:ser>' +
                '<c:idx val="' + di + '"/><c:order val="' + di + '"/>' +
                '<c:tx><c:strRef><c:f>' + "'" + statSheetName + "'!$" + colLetter + "$" + (SH_ROW + 1) + '</c:f></c:strRef></c:tx>' +
                '<c:spPr><a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill></c:spPr>' +
                '<c:cat><c:strRef><c:f>' + catFormula + '</c:f></c:strRef></c:cat>' +
                '<c:val><c:numRef><c:f>' + valFormula + '</c:f></c:numRef></c:val>' +
                '</c:ser>';
        });

        var chartXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
            '<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="de-DE" sz="2000" b="1"/>' +
            '<a:t>Bodenwiderstand \u03C1 [\u03A9m] - ' + projectName + '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>' +
            '<c:autoTitleDeleted val="0"/><c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:rAngAx val="1"/><c:perspective val="30"/></c:view3D><c:plotArea><c:layout/>' +
            '<c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:shape val="cylinder"/>' +
            seriesXml +
            '<c:axId val="111"/><c:axId val="222"/><c:axId val="333"/></c:bar3DChart>' +
            '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:catAx>' +
            '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111"/>' +
            '<c:title><c:tx><c:rich><a:bodyPr rot="-5400000" vert="horz"/><a:lstStyle/><a:p><a:r><a:rPr lang="de-DE" sz="1400"/><a:t>Bodenwiderstand [\u03A9m]</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx>' +
            '<c:serAx><c:axId val="333"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:crossAx val="222"/></c:serAx>' +
            '</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>';

        zip.file('xl/charts/chart1.xml', chartXml);

        var anchorRow = dataEndRow + 2;
        var drawingXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
            '<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + anchorRow + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
            '<xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + (anchorRow + 20) + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
            '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Diagramm 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
            '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>' +
            '</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>';

        var drawingNum = zip.file('xl/drawings/drawing2.xml') ? 3 : 2;
        var drawingPath = 'xl/drawings/drawing' + drawingNum + '.xml';
        zip.file(drawingPath, drawingXml);

        var drawingRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>';
        zip.file('xl/drawings/_rels/drawing' + drawingNum + '.xml.rels', drawingRels);

        var ct = await zip.file('[Content_Types].xml').async('string');
        if (!ct.includes('chart1.xml')) {
            ct = ct.replace('</Types>',
                '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' +
                '<Override PartName="/' + drawingPath + '" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>');
            zip.file('[Content_Types].xml', ct);
        }

        var sheet2 = await zip.file('xl/worksheets/sheet2.xml').async('string');
        if (!sheet2.includes('drawing r:id')) {
            sheet2 = sheet2.replace('</worksheet>', '<drawing r:id="rIdChart1"/></worksheet>');
            zip.file('xl/worksheets/sheet2.xml', sheet2);
        }

        var s2RelsPath = 'xl/worksheets/_rels/sheet2.xml.rels';
        var s2Rels;
        if (zip.file(s2RelsPath)) {
            s2Rels = await zip.file(s2RelsPath).async('string');
            if (!s2Rels.includes('rIdChart1')) {
                s2Rels = s2Rels.replace('</Relationships>',
                    '<Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing' + drawingNum + '.xml"/></Relationships>');
            }
        } else {
            s2Rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing' + drawingNum + '.xml"/></Relationships>';
        }
        zip.file(s2RelsPath, s2Rels);

        return await zip.generateAsync({ type: 'arraybuffer' });
    } catch (err) {
        console.warn('Chart injection failed:', err);
        return xlsxBuffer;
    }
}

// ─── PLOT ───
function renderAppPlot() {
    const container = $('appPlotContainer');
    if (!container) return;

    const data = AppState.data.filter(row => {
        return ['0.8', '1.6', '3.2'].some(d => parseFloat(row['MW [Ωm]_' + d]) > 0);
    });
    const depths = ['0.8', '1.6', '3.2'].filter(d => !AppState.hiddenColumns.has(d === '0.8' ? '08' : (d === '1.6' ? '16' : '32')));

    if (data.length === 0 || depths.length === 0) {
        container.innerHTML = '<div style="text-align:center;color:var(--text-muted);padding:40px;">Keine Daten für Statistik</div>';
        return;
    }

    container.innerHTML = '<canvas id="liveChartCanvas" style="width:100%;height:320px;"></canvas>';
    const canvas = $('liveChartCanvas');
    const ctx = canvas.getContext('2d');

    const dpr = window.devicePixelRatio || 1;
    canvas.width = canvas.clientWidth * dpr;
    canvas.height = 320 * dpr;
    ctx.scale(dpr, dpr);

    const w = canvas.clientWidth, h = 320;
    const pad = { top: 40, bottom: 55, left: 65, right: 20 };
    const chartW = w - pad.left - pad.right;
    const chartH = h - pad.top - pad.bottom;

    const colors = { '0.8': '#e879f9', '1.6': '#fbbf24', '3.2': '#22d3ee' };

    let maxV = 100;
    data.forEach(row => {
        depths.forEach(d => {
            const mw = parseFloat(row['MW [Ωm]_' + d]) || 0;
            const sd = parseFloat(row['SD [Ωm]_' + d]) || 0;
            maxV = Math.max(maxV, mw + sd);
        });
    });
    maxV *= 1.2;

    // Grid
    ctx.strokeStyle = 'rgba(148,163,184,0.15)';
    ctx.lineWidth = 1;
    ctx.fillStyle = '#94a3b8';
    ctx.font = '10px Inter';
    ctx.textAlign = 'right';
    for (var i = 0; i <= 5; i++) {
        var y = pad.top + chartH - (chartH * (i / 5));
        ctx.beginPath(); ctx.moveTo(pad.left, y); ctx.lineTo(pad.left + chartW, y); ctx.stroke();
        ctx.fillText(Math.round(maxV * (i / 5)), pad.left - 8, y + 3);
    }

    var bG = chartW / data.length;

    if (AppState.chartMode === 'scatter') {
        // SCATTER MODE
        var lineStyles = { '0.8': [], '1.6': [12, 6], '3.2': [5, 5] };
        depths.forEach(function(d) {
            var color = colors[d];
            ctx.strokeStyle = color;
            ctx.lineWidth = 3;
            ctx.setLineDash(lineStyles[d] || []);
            ctx.beginPath();
            var started = false;
            data.forEach(function(row, i) {
                var val = parseFloat(row['MW [Ωm]_' + d]) || 0;
                if (val <= 0) return;
                var cx = pad.left + i * bG + bG / 2;
                var cy = pad.top + chartH - (val / maxV) * chartH;
                if (!started) { ctx.moveTo(cx, cy); started = true; }
                else ctx.lineTo(cx, cy);
            });
            ctx.stroke();
            ctx.setLineDash([]);

            data.forEach(function(row, i) {
                var val = parseFloat(row['MW [Ωm]_' + d]) || 0;
                if (val <= 0) return;
                var cx = pad.left + i * bG + bG / 2;
                var cy = pad.top + chartH - (val / maxV) * chartH;
                ctx.fillStyle = color;
                ctx.strokeStyle = '#0f172a'; ctx.lineWidth = 2;
                // Different marker per depth
                if (d === '0.8') {
                    ctx.beginPath(); ctx.arc(cx, cy, 10, 0, Math.PI * 2); ctx.fill();
                    ctx.beginPath(); ctx.arc(cx, cy, 10, 0, Math.PI * 2); ctx.stroke();
                } else if (d === '1.6') {
                    ctx.beginPath(); ctx.moveTo(cx, cy - 12); ctx.lineTo(cx + 10, cy); ctx.lineTo(cx, cy + 12); ctx.lineTo(cx - 10, cy); ctx.closePath(); ctx.fill(); ctx.stroke();
                } else {
                    ctx.fillRect(cx - 9, cy - 9, 18, 18);
                    ctx.strokeRect(cx - 9, cy - 9, 18, 18);
                }
                // STD error bars on scatter
                var sdScatter = parseFloat(row['SD [\u03A9m]_' + d]) || 0;
                if (sdScatter > 0) {
                    var sdPx = (sdScatter / maxV) * chartH;
                    ctx.strokeStyle = color; ctx.lineWidth = 2;
                    ctx.beginPath(); ctx.moveTo(cx, cy - sdPx); ctx.lineTo(cx, cy + sdPx); ctx.stroke();
                    ctx.beginPath(); ctx.moveTo(cx - 6, cy - sdPx); ctx.lineTo(cx + 6, cy - sdPx); ctx.stroke();
                    ctx.beginPath(); ctx.moveTo(cx - 6, cy + sdPx); ctx.lineTo(cx + 6, cy + sdPx); ctx.stroke();
                }
                // Value + STD text
                ctx.fillStyle = color; ctx.font = 'bold 15px Inter'; ctx.textAlign = 'center';
                var scLabelY = sdScatter > 0 ? cy - (sdScatter / maxV) * chartH - 14 : cy - 18;
                ctx.fillText(val.toFixed(1), cx, scLabelY);
                if (sdScatter > 0) {
                    ctx.font = '12px Inter';
                    ctx.fillText('\u00B1 ' + sdScatter.toFixed(1), cx, scLabelY + 13);
                }
            });

            // Linear regression trend line
            var xVals = [], yVals = [];
            data.forEach(function(row, i) {
                var val = parseFloat(row['MW [Ωm]_' + d]) || 0;
                if (val > 0) { xVals.push(i); yVals.push(val); }
            });
            if (xVals.length >= 2) {
                var n = xVals.length;
                var sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
                for (var j = 0; j < n; j++) { sumX += xVals[j]; sumY += yVals[j]; sumXY += xVals[j] * yVals[j]; sumX2 += xVals[j] * xVals[j]; }
                var slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
                var intercept = (sumY - slope * sumX) / n;

                ctx.strokeStyle = color; ctx.lineWidth = 1.5; ctx.setLineDash([8, 5]);
                ctx.beginPath();
                var x0 = 0, x1 = data.length - 1;
                var y0val = intercept, y1val = slope * x1 + intercept;
                var px0 = pad.left + x0 * bG + bG / 2;
                var py0 = pad.top + chartH - (y0val / maxV) * chartH;
                var px1 = pad.left + x1 * bG + bG / 2;
                var py1 = pad.top + chartH - (y1val / maxV) * chartH;
                ctx.moveTo(px0, Math.max(pad.top, Math.min(pad.top + chartH, py0)));
                ctx.lineTo(px1, Math.max(pad.top, Math.min(pad.top + chartH, py1)));
                ctx.stroke();
                ctx.setLineDash([]);

                // Trend line only (no text label)
            }
        });
    } else {
        // BAR MODE - Cylinder/Tube style with STD display
        var barW = (bG * 0.7) / depths.length;
        data.forEach(function(row, i) {
            var gX = pad.left + i * bG;
            depths.forEach(function(d, di) {
                var mw = parseFloat(row['MW [Ωm]_' + d]) || 0;
                var sd = parseFloat(row['SD [Ωm]_' + d]) || 0;
                if (mw <= 0) return;
                var barH = (mw / maxV) * chartH;
                var bx = gX + di * barW + bG * 0.15;
                var by = pad.top + chartH - barH;
                var bWidth = barW - 4;
                var cx = bx + bWidth / 2;

                // Cylinder with hatching (reference style)
                ctx.fillStyle = colors[d] + '30';
                ctx.fillRect(bx, by, bWidth, barH);
                // Diagonal hatching
                ctx.save();
                ctx.beginPath(); ctx.rect(bx, by, bWidth, barH); ctx.clip();
                ctx.strokeStyle = colors[d] + '70'; ctx.lineWidth = 1;
                for (var hi = -barH; hi < bWidth + barH; hi += 6) {
                    ctx.beginPath(); ctx.moveTo(bx + hi, by + barH); ctx.lineTo(bx + hi + barH, by); ctx.stroke();
                }
                ctx.restore();
                // Border
                ctx.strokeStyle = colors[d]; ctx.lineWidth = 1.5;
                ctx.strokeRect(bx, by, bWidth, barH);
                // Top ellipse (hollow cap)
                ctx.fillStyle = '#0f172a';
                ctx.beginPath(); ctx.ellipse(cx, by, bWidth / 2, 7, 0, 0, Math.PI * 2); ctx.fill();
                ctx.strokeStyle = colors[d]; ctx.lineWidth = 1.5;
                ctx.beginPath(); ctx.ellipse(cx, by, bWidth / 2, 7, 0, 0, Math.PI * 2); ctx.stroke();
                ctx.strokeStyle = colors[d] + '88'; ctx.lineWidth = 1;
                ctx.beginPath(); ctx.ellipse(cx, by, bWidth / 2 - 3, 4, 0, 0, Math.PI * 2); ctx.stroke();
                // Side edges
                ctx.strokeStyle = colors[d]; ctx.lineWidth = 1.5;
                ctx.beginPath(); ctx.moveTo(bx, by); ctx.lineTo(bx, pad.top + chartH); ctx.stroke();
                ctx.beginPath(); ctx.moveTo(bx + bWidth, by); ctx.lineTo(bx + bWidth, pad.top + chartH); ctx.stroke();

                // Error bar (STD)
                if (sd > 0) {
                    var sdH = (sd / maxV) * chartH;
                    ctx.strokeStyle = '#1f2937'; ctx.lineWidth = 2;
                    ctx.beginPath(); ctx.moveTo(cx, by - sdH); ctx.lineTo(cx, by + Math.min(sdH, barH * 0.7)); ctx.stroke();
                    // Caps
                    ctx.beginPath(); ctx.moveTo(cx - 7, by - sdH); ctx.lineTo(cx + 7, by - sdH); ctx.stroke();
                    ctx.beginPath(); ctx.moveTo(cx - 7, by + Math.min(sdH, barH * 0.7)); ctx.lineTo(cx + 7, by + Math.min(sdH, barH * 0.7)); ctx.stroke();
                }

                // MW value label (colored, above bar)
                ctx.fillStyle = colors[d]; ctx.font = 'bold 16px Inter'; ctx.textAlign = 'center';
                var labelY = sd > 0 ? by - (sd / maxV) * chartH - 14 : by - 10;
                ctx.fillText(mw.toFixed(1), cx, labelY);

                // STD value label (colored, below MW)
                if (sd > 0) {
                    ctx.fillStyle = colors[d]; ctx.font = '13px Inter';
                    ctx.fillText('± ' + sd.toFixed(1), cx, labelY + 11);
                }
            });
        });
    }

    // X labels
    data.forEach(function(row, i) {
        if (data.length < 20 || i % Math.ceil(data.length / 12) === 0) {
            ctx.fillStyle = '#94a3b8'; ctx.textAlign = 'center'; ctx.font = 'bold 12px Inter';
            ctx.fillText(row['Kennzeichen'] || (i + 1), pad.left + i * bG + bG / 2, pad.top + chartH + 14);
        }
    });

    // Axis labels
    ctx.fillStyle = '#f1f5f9'; ctx.font = 'bold 14px Inter'; ctx.textAlign = 'center';
    ctx.fillText('Kennzeichen', pad.left + chartW / 2, h - 10);
    ctx.save(); ctx.translate(14, pad.top + chartH / 2); ctx.rotate(-Math.PI / 2);
    ctx.fillText('Bodenwiderstand [Ωm]', 0, 0); ctx.restore();

    // Legend
    var lx = pad.left + 10;
    depths.forEach(function(d) {
        ctx.fillStyle = colors[d]; ctx.fillRect(lx, 8, 14, 14);
        ctx.fillStyle = '#f1f5f9'; ctx.font = 'bold 12px Inter'; ctx.textAlign = 'left';
        ctx.fillText(d + 'm', lx + 18, 20);
        lx += 70;
    });
}