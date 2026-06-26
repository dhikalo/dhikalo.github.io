/* ============================================
   MESSSTELLEN MANAGER PRO - v6.0
   Professional Field Measurement Tool
   Bug-free, Performance-optimized
   ============================================ */

'use strict';

// ─── DATA STRUCTURE ───
const DEPTHS = ['0.8', '1.6', '3.2'];

const TABLE_STRUCTURE = [
    { group: 'Basis', class: 'basis', columns: ['Kennzeichen', 'Alt-Kz.', 'Typ', 'Örtlichkeit', 'Meter [m]', 'Datum', 'Sprachsteuerung'] },
    { group: '0.8m', class: '08', columns: ['R1 [Ω]_0.8', 'R2 [Ω]_0.8', 'R3 [Ω]_0.8', 'ρ1 [Ωm]_0.8', 'ρ2 [Ωm]_0.8', 'ρ3 [Ωm]_0.8', 'MW [Ωm]_0.8', 'SD [Ωm]_0.8', 'Bilder_0.8'] },
    { group: '1.6m', class: '16', columns: ['R1 [Ω]_1.6', 'R2 [Ω]_1.6', 'R3 [Ω]_1.6', 'ρ1 [Ωm]_1.6', 'ρ2 [Ωm]_1.6', 'ρ3 [Ωm]_1.6', 'MW [Ωm]_1.6', 'SD [Ωm]_1.6', 'Bilder_1.6'] },
    { group: '3.2m', class: '32', columns: ['R1 [Ω]_3.2', 'R2 [Ω]_3.2', 'R3 [Ω]_3.2', 'ρ1 [Ωm]_3.2', 'ρ2 [Ωm]_3.2', 'ρ3 [Ωm]_3.2', 'MW [Ωm]_3.2', 'SD [Ωm]_3.2', 'Bilder_3.2'] },
    { group: 'Anhang (Gesamt)', class: 'anhang_global', columns: ['Anhang_Global'] },
    { group: 'GPS-Daten', class: 'special', columns: ['Koordinaten'] },
    { group: 'Potential', class: 'potential', columns: ['Pot. Ein [V]', 'Pot. Aus [V]', 'AC [V]'] },
    { group: 'Spannung', class: 'spannung', columns: ['Spannung Ein', 'Spannung Aus'] },
    { group: 'Strom', class: 'strom', columns: ['Strom Ein', 'Strom Aus', 'Strom Diff.', 'Mikro Diff.'] },
    { group: 'Widerstand', class: 'widerstand', columns: ['R [Ω]', 'Bodenwid. [Ωm]', '+/-', 'Ra', 'Kommentar'] },
    { group: 'Audit', class: 'audit', columns: ['Erderinformationen', 'Existiert die Messstelle', 'Typ korrekt', 'Fernüberwacht'] }
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
    hiddenColumns: new Set(['32', 'potential', 'spannung', 'strom', 'widerstand', 'audit', 'Kennzeichen', 'Alt-Kz.']),
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
    hiddenMapColors: new Set([]),
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

function convertToDMS(coord, isLat) {
    var absCoord = Math.abs(coord);
    var deg = Math.floor(absCoord);
    var minFloat = (absCoord - deg) * 60;
    var min = Math.floor(minFloat);
    var sec = ((minFloat - min) * 60).toFixed(1);
    var dir = isLat ? (coord >= 0 ? 'N' : 'S') : (coord >= 0 ? 'E' : 'W');
    return deg + '°' + min.toString().padStart(2, '0') + "'" + sec.padStart(4, '0') + '"' + dir;
}

function parseCoordinates(str) {
    if (!str) return null;
    str = str.trim();

    // Check DMS format (e.g. 51°31'27.1"N 7°09'49.0"E)
    const dmsRegex = /(\d+)\s*[°dD]\s*(\d+)\s*['’′]\s*(\d+(?:[.,]\d+)?)\s*["”″]?\s*([NSEWnsew])/g;
    const matches = [];
    let match;
    dmsRegex.lastIndex = 0;
    while ((match = dmsRegex.exec(str)) !== null) {
        matches.push(match);
    }

    if (matches.length === 2) {
        const parseDMSVal = (m) => {
            const deg = parseFloat(m[1]);
            const min = parseFloat(m[2]);
            const sec = parseFloat(m[3].replace(',', '.'));
            const hemi = m[4].toUpperCase();
            let decimal = deg + (min / 60) + (sec / 3600);
            if (hemi === 'S' || hemi === 'W') {
                decimal = -decimal;
            }
            return decimal;
        };
        const lat = parseDMSVal(matches[0]);
        const lng = parseDMSVal(matches[1]);
        return { lat, lng };
    }

    // Try standard decimal format e.g. "51.52419, 7.16361"
    const decimalRegex = /([+-]?\d+(?:[.,]\d+)?)/g;
    const decMatches = [];
    let decMatch;
    decimalRegex.lastIndex = 0;
    while ((decMatch = decimalRegex.exec(str)) !== null) {
        decMatches.push(parseFloat(decMatch[1].replace(',', '.')));
    }
    if (decMatches.length >= 2) {
        return { lat: decMatches[0], lng: decMatches[1] };
    }

    return null;
}


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

// Touch pinch-to-zoom for table wrapper
(function() {
    let initialDist = null;
    let initialZoom = 100;

    document.addEventListener('DOMContentLoaded', () => {
        const wrapper = document.querySelector('.table-wrapper');
        if (!wrapper) return;

        wrapper.addEventListener('touchstart', (e) => {
            if (e.touches.length === 2) {
                e.preventDefault();
                initialDist = Math.hypot(
                    e.touches[0].clientX - e.touches[1].clientX,
                    e.touches[0].clientY - e.touches[1].clientY
                );
                initialZoom = AppState.zoomLevel;
            }
        }, { passive: false });

        wrapper.addEventListener('touchmove', (e) => {
            if (e.touches.length === 2 && initialDist !== null) {
                e.preventDefault();
                const dist = Math.hypot(
                    e.touches[0].clientX - e.touches[1].clientX,
                    e.touches[0].clientY - e.touches[1].clientY
                );
                const factor = dist / initialDist;
                let newZoom = Math.round((initialZoom * factor) / 5) * 5;
                newZoom = Math.max(50, Math.min(200, newZoom));

                if (newZoom !== AppState.zoomLevel) {
                    AppState.zoomLevel = newZoom;
                    const display = $('zoomLevelDisplay');
                    if (display) display.textContent = newZoom + '%';
                    wrapper.style.zoom = newZoom / 100;
                }
            }
        }, { passive: false });

        wrapper.addEventListener('touchend', (e) => {
            if (e.touches.length < 2) {
                initialDist = null;
            }
        });
    });
})();

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

// ─── DEBOUNCED SAVE ───
let _saveTimer = null;
function debouncedSave(delay) {
    delay = delay || 800;
    if (_saveTimer) clearTimeout(_saveTimer);
    _saveTimer = setTimeout(function() {
        _saveTimer = null;
        saveToStorage();
        // Flash autosave indicator
        var dot = $('autosaveDot');
        if (dot) {
            dot.classList.add('saving');
            setTimeout(function() { dot.classList.remove('saving'); }, 700);
        }
    }, delay);
}

// ─── INITIALIZATION ───
document.addEventListener('DOMContentLoaded', () => {
    // Force service worker update
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.getRegistrations().then(regs => {
            regs.forEach(reg => reg.update());
        });
        navigator.serviceWorker.register('./sw.js').then(function(reg) {
            reg.update();
            // Check for updates periodically
            reg.addEventListener('updatefound', function() {
                var newWorker = reg.installing;
                if (newWorker) {
                    newWorker.addEventListener('statechange', function() {
                        if (newWorker.state === 'installed') {
                            if (navigator.serviceWorker.controller) {
                                newWorker.postMessage({ action: 'skipWaiting' });
                                showToast('App-Aktualisierung wird installiert...');
                            }
                        }
                    });
                }
            });
        }).catch(function(err) {
            console.warn('SW registration failed:', err);
        });

        let refreshing = false;
        navigator.serviceWorker.addEventListener('controllerchange', function() {
            if (!refreshing) {
                refreshing = true;
                window.location.reload();
            }
        });
    }

    // Initialize Firebase cloud sync
    initCloudSync();

    initUI();
    initConnectionCheck();
    initStabilityFeatures();
    initGlobalErrorHandler();
});

// ─── CLOUD SYNC INITIALIZATION ───
function initCloudSync() {
    // Try to initialize Firebase
    if (typeof initFirebase === 'function') {
        const success = initFirebase();
        if (success) {
            console.log('Cloud sync enabled');
        }
    }

    // Always show role screen — pre-select last used role if available
    let selectedRole = null;
    const savedRole = localStorage.getItem('messstellen_role');
    const roleOverlay = $('roleOverlay');

    if (roleOverlay) {
        roleOverlay.style.display = 'flex';
        if (savedRole && savedRole !== 'local') {
            const cardId = savedRole === 'messhelfer' ? 'roleMesshelfer' : 'rolePruefer';
            const savedCard = $(cardId);
            if (savedCard) {
                savedCard.classList.add('selected');
                selectedRole = savedRole;
                const nameSection = $('roleNameSection');
                if (nameSection) nameSection.style.display = 'flex';
                const nameInput = $('roleNameInput');
                if (nameInput) nameInput.value = localStorage.getItem('messstellen_userName') || '';
            }
        }
    }

    // Role card click handlers
    const roleCards = document.querySelectorAll('.role-card');
    roleCards.forEach(card => {
        card.addEventListener('click', () => {
            roleCards.forEach(c => c.classList.remove('selected'));
            card.classList.add('selected');
            selectedRole = card.id === 'roleMesshelfer' ? 'messhelfer' : 'pruefer';
            const nameSection = $('roleNameSection');
            if (nameSection) nameSection.style.display = 'flex';
            const nameInput = $('roleNameInput');
            if (nameInput) nameInput.focus();
        });
    });

    // Confirm role
    on('roleConfirmBtn', 'click', () => {
        if (!selectedRole) { showToast('Bitte eine Rolle auswählen'); return; }
        const name = ($('roleNameInput') || {}).value || (selectedRole === 'pruefer' ? 'Prüfer' : 'Messhelfer');
        if (typeof setUserRole === 'function') setUserRole(selectedRole, name);
        const roleOverlay = $('roleOverlay');
        if (roleOverlay) roleOverlay.style.display = 'none';
        applyRoleUI(selectedRole);

        if (selectedRole === 'messhelfer') {
            // Skip welcome screen — go straight to main app
            const welcomeOverlay = $('welcomeOverlay');
            if (welcomeOverlay) welcomeOverlay.style.display = 'none';
            const mainApp = $('mainApp');
            if (mainApp) mainApp.style.display = 'flex';
            // Start with clean empty table — no pre-filled rows
            AppState.data = [];
            renderTable();
        } else if (selectedRole === 'pruefer') {
            // Show Prüfer cloud project list
            const welcomeOverlay = $('welcomeOverlay');
            if (welcomeOverlay) welcomeOverlay.style.display = 'flex';
        }
    });

    // Enter key on name input
    on('roleNameInput', 'keypress', (e) => {
        if (e.key === 'Enter') {
            const btn = $('roleConfirmBtn');
            if (btn) btn.click();
        }
    });

    // Skip cloud mode — go straight to main app
    on('roleSkipBtn', 'click', () => {
        localStorage.setItem('messstellen_role', 'local');
        const roleOverlay = $('roleOverlay');
        if (roleOverlay) roleOverlay.style.display = 'none';
        const welcomeOverlay = $('welcomeOverlay');
        if (welcomeOverlay) welcomeOverlay.style.display = 'none';
        const mainApp = $('mainApp');
        if (mainApp) mainApp.style.display = 'flex';
        AppState.data = [];
        renderTable();
        if (typeof updateCloudStatus === 'function') updateCloudStatus('offline');
    });
}

function applyRoleUI(role) {
    const btnSubmit = $('btnSubmitReview');
    const btnReview = $('btnReviewPanel');
    const btnExport = $('btnExportExcel');
    const btnSave = $('btnSaveAll');

    if (role === 'messhelfer') {
        if (btnSubmit) btnSubmit.style.display = 'flex';
        if (btnReview) btnReview.style.display = 'none';
        // Show Messhelfer welcome
        const wM = $('welcomeMesshelfer');
        const wP = $('welcomePruefer');
        if (wM) wM.style.display = '';
        if (wP) wP.style.display = 'none';
        const sub = $('welcomeSubtitle');
        if (sub) sub.textContent = 'Feldmessung & Dokumentation';

    } else if (role === 'pruefer') {
        if (btnSubmit) btnSubmit.style.display = 'none';
        if (btnReview) btnReview.style.display = 'flex';
        // Prüfer: hide local save button (they export after approval instead)
        // Show Prüfer welcome instead of Messhelfer welcome
        const wM = $('welcomeMesshelfer');
        const wP = $('welcomePruefer');
        if (wM) wM.style.display = 'none';
        if (wP) wP.style.display = '';
        // Load cloud projects into the Prüfer welcome list
        loadPrueferWelcomeList('submitted');
        // Start listening for submitted projects (for notification dot)
        if (typeof listenForAllProjects === 'function' && typeof isFirebaseConfigured === 'function' && isFirebaseConfigured()) {
            listenForAllProjects(renderReviewList);
        }

    } else if (role === 'local') {
        if (btnSubmit) btnSubmit.style.display = 'none';
        if (btnReview) btnReview.style.display = 'none';
        const wM = $('welcomeMesshelfer');
        const wP = $('welcomePruefer');
        if (wM) wM.style.display = '';
        if (wP) wP.style.display = 'none';
    }

    // Listen for review status updates on current project (for Messhelfer)
    if (role === 'messhelfer' && typeof listenForCloudUpdates === 'function' && typeof isFirebaseConfigured === 'function' && isFirebaseConfigured()) {
        const pName = ($('projectName') || {}).value;
        if (pName) {
            listenForCloudUpdates(pName, (data) => {
                if (data.reviewStatus && data.reviewStatus !== 'draft') {
                    showReviewStatusBar(data.reviewStatus, data.reviewComment, data.reviewedBy);
                }
            });
        }
    }
}

// ─── PRÜFER WELCOME LIST ───
let _prueferWelcomeFilter = 'submitted';
let _prueferAllProjects = [];

async function loadPrueferWelcomeList(filter) {
    _prueferWelcomeFilter = filter || 'submitted';
    const list = $('prueferProjectList');
    if (!list) return;

    list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:24px;"><i class="fas fa-spinner fa-spin"></i> Lade...</p>';

    if (typeof isFirebaseConfigured === 'function' && isFirebaseConfigured()) {
        try {
            _prueferAllProjects = await getAllCloudProjects();
        } catch (e) {
            _prueferAllProjects = [];
        }
    } else {
        const library = JSON.parse(localStorage.getItem('messstellen_library') || '{}');
        _prueferAllProjects = Object.values(library).map(p => ({
            id: p.name, name: p.name, data: p.data || [],
            rowCount: (p.data || []).length,
            reviewStatus: p.reviewStatus || 'draft',
            updatedAt: p.timestamp ? { toDate: () => new Date(p.timestamp) } : null,
            updatedBy: 'Lokal', reviewComment: ''
        }));
    }

    renderPrueferWelcomeList();
}

function renderPrueferWelcomeList() {
    const list = $('prueferProjectList');
    if (!list) return;

    let filtered = _prueferAllProjects;
    if (_prueferWelcomeFilter !== 'all') {
        filtered = _prueferAllProjects.filter(p => p.reviewStatus === _prueferWelcomeFilter);
    }

    if (filtered.length === 0) {
        const msgs = { submitted: 'Keine Projekte zur Prüfung', approved: 'Keine genehmigten Projekte', rejected: 'Keine abgelehnten Projekte', all: 'Keine Projekte in der Cloud' };
        list.innerHTML = `<p style="color:var(--text-muted);text-align:center;padding:24px;font-size:13px;"><i class="fas fa-inbox" style="font-size:20px;display:block;margin-bottom:8px;opacity:0.3;"></i>${msgs[_prueferWelcomeFilter] || 'Keine Projekte'}</p>`;
        return;
    }

    list.innerHTML = '';
    filtered.forEach(project => {
        const statusLabels = { draft: '📝', submitted: '⏳', approved: '✅', rejected: '❌' };
        const statusClasses = { draft: 'status-draft', submitted: 'status-submitted', approved: 'status-approved', rejected: 'status-rejected' };
        let dateStr = '';
        if (project.updatedAt) {
            try { dateStr = (project.updatedAt.toDate ? project.updatedAt.toDate() : new Date(project.updatedAt)).toLocaleString('de-DE'); } catch(e){}
        }

        const card = document.createElement('div');
        card.className = 'review-project-card';
        card.innerHTML = `
            <div class="review-project-info">
                <div class="review-project-name">${project.name || project.id}</div>
                <div class="review-project-meta">
                    ${project.updatedBy ? `<span>${project.updatedBy}</span>` : ''}
                    ${dateStr ? `<span>·</span><span>${dateStr}</span>` : ''}
                </div>
                ${project.reviewComment ? `<div style="font-size:11px;color:var(--danger);margin-top:3px;"><i class="fas fa-comment"></i> ${project.reviewComment}</div>` : ''}
            </div>
            <span class="review-status-chip ${statusClasses[project.reviewStatus] || ''}">${statusLabels[project.reviewStatus] || '?'} ${project.reviewStatus || ''}</span>
        `;

        // Click → open project in main app for review
        card.onclick = () => openProjectForReview(project);
        list.appendChild(card);
    });
}

async function openProjectForReview(project) {
    // Load project data into the app
    const pName = project.name || project.id;

    // Try to get full data with images from localStorage first
    const lib = JSON.parse(localStorage.getItem('messstellen_library') || '{}');
    const localProject = lib[pName];

    // Use cloud data as base, but restore images from localStorage if available
    let projectData = project.data || [];
    if (localProject && localProject.data && localProject.data.length > 0) {
        // Merge: use cloud data but fill in images from local where cloud has __IMAGE_REF__
        projectData = project.data.map((cloudRow, idx) => {
            const localRow = localProject.data[idx] || {};
            const mergedRow = Object.assign({}, cloudRow);
            Object.keys(mergedRow).forEach(key => {
                if (mergedRow[key] === '__IMAGE_REF__' && localRow[key] && localRow[key].startsWith('data:image')) {
                    mergedRow[key] = localRow[key];
                }
            });
            return mergedRow;
        });
    }

    AppState.data = projectData;
    const nameInput = $('projectName');
    if (nameInput) nameInput.value = pName;

    if (project.newCols) {
        AppState.newCols = new Set(project.newCols);
        let zusatzGroup = TABLE_STRUCTURE.find(g => g.group === 'Zusatz');
        if (!zusatzGroup) {
            zusatzGroup = { group: 'Zusatz', class: 'zusatz', columns: [] };
            TABLE_STRUCTURE.push(zusatzGroup);
        }
        AppState.newCols.forEach(col => { if (!zusatzGroup.columns.includes(col)) zusatzGroup.columns.push(col); });
    }

    // Store map data for restoreMapDrawings
    if (project.mapData || project.starData) {
        if (!lib[pName]) lib[pName] = {};
        lib[pName].mapData = project.mapData || null;
        lib[pName].starData = project.starData || [];
        lib[pName].data = projectData;
        lib[pName].name = pName;
        lib[pName].hiddenMapColors = project.hiddenMapColors || [];
        lib[pName].newCols = project.newCols || [];
        localStorage.setItem('messstellen_library', JSON.stringify(lib));
        localStorage.setItem('current_project_id', pName);
    }

    // Hide welcome, show main app
    const welcomeOverlay = $('welcomeOverlay');
    if (welcomeOverlay) welcomeOverlay.style.display = 'none';
    const mainApp = $('mainApp');
    if (mainApp) mainApp.style.display = 'flex';

    renderTable();

    // If submitted → open review detail immediately
    if (project.reviewStatus === 'submitted') {
        setTimeout(() => openReviewDetail(project), 400);
    }
}

function initConnectionCheck() {
    var offlineEl = $('offlineIndicator');
    
    function updateOnlineStatus() {
        if (navigator.onLine) {
            if (offlineEl) offlineEl.classList.remove('show');
            showToast('Online');
        } else {
            if (offlineEl) offlineEl.classList.add('show');
            showToast('Offline-Modus aktiv — Daten werden lokal gespeichert');
        }
    }

    window.addEventListener('online', updateOnlineStatus);
    window.addEventListener('offline', updateOnlineStatus);

    // Initial check
    if (!navigator.onLine && offlineEl) {
        offlineEl.classList.add('show');
    }
}

function initStabilityFeatures() {
    // Auto-save every 30 seconds
    setInterval(function() {
        if (AppState.data && AppState.data.length > 0) {
            try {
                saveToStorage();
                var dot = $('autosaveDot');
                if (dot) {
                    dot.classList.add('saving');
                    setTimeout(function() { dot.classList.remove('saving'); }, 700);
                }
            } catch (e) {
                console.warn('Auto-save failed:', e);
            }
        }
    }, 30000);

    // Save when app goes to background (critical for field: user switches to camera)
    document.addEventListener('visibilitychange', function() {
        if (document.visibilityState === 'hidden') {
            try {
                if (_saveTimer) {
                    clearTimeout(_saveTimer);
                    _saveTimer = null;
                }
                saveToStorage();
            } catch (e) {
                console.warn('Visibility save failed:', e);
            }
        }
    });

    // Warn before closing/navigating away with unsaved data
    window.addEventListener('beforeunload', function(e) {
        if (AppState.data && AppState.data.length > 0) {
            try { saveToStorage(); } catch (err) { /* best effort */ }
            e.preventDefault();
            e.returnValue = 'Ungespeicherte Daten gehen verloren!';
        }
    });

    // Prevent iOS rubber-banding on main app (causes accidental refreshes)
    document.body.addEventListener('touchmove', function(e) {
        var t = e.target;
        // Allow scrolling in scrollable containers
        while (t && t !== document.body) {
            var style = window.getComputedStyle(t);
            if (style.overflow === 'auto' || style.overflow === 'scroll' ||
                style.overflowY === 'auto' || style.overflowY === 'scroll' ||
                style.overflowX === 'auto' || style.overflowX === 'scroll') {
                return; // Allow scroll in scrollable areas
            }
            if (t.classList && (t.classList.contains('table-wrapper') ||
                t.classList.contains('map-view') ||
                t.classList.contains('leaflet-container') ||
                t.classList.contains('modal-content') ||
                t.classList.contains('map-sidebar'))) {
                return;
            }
            t = t.parentElement;
        }
        // Prevent bounce/pull-to-refresh on non-scrollable areas
        if (document.body.scrollHeight <= window.innerHeight) {
            e.preventDefault();
        }
    }, { passive: false });
}

function initGlobalErrorHandler() {
    window.addEventListener('error', function(event) {
        console.error('Unhandled error:', event.error);
        showToast('Fehler: ' + (event.message || 'Unbekannter Fehler'));
        event.preventDefault();
    });

    window.addEventListener('unhandledrejection', function(event) {
        console.error('Unhandled promise rejection:', event.reason);
        showToast('Fehler: ' + (event.reason && event.reason.message ? event.reason.message : 'Async-Fehler'));
        event.preventDefault();
    });
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

    // Prüfer welcome filter tabs
    document.querySelectorAll('[data-pfilter]').forEach(tab => {
        tab.addEventListener('click', () => {
            document.querySelectorAll('[data-pfilter]').forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            loadPrueferWelcomeList(tab.dataset.pfilter);
        });
    });
    on('btnPrueferRefresh', 'click', () => loadPrueferWelcomeList(_prueferWelcomeFilter));

    // Header buttons
    on('btnHome', 'click', () => {
        $('mainApp').style.display = 'none';
        const role = localStorage.getItem('messstellen_role');
        if (role === 'pruefer') {
            $('welcomeOverlay').style.display = 'flex';
            loadPrueferWelcomeList(_prueferWelcomeFilter);
        } else {
            // Messhelfer/local: back to role screen
            const roleOverlay = $('roleOverlay');
            if (roleOverlay) roleOverlay.style.display = 'flex';
        }
    });
    on('btnBrowse', 'click', () => $('fileInput').click());
    on('btnStartBlank', 'click', () => {
        closeMap();
        AppState.data = [];
        AppState.newCols = new Set();
        const nameInput = $('projectName');
        if (nameInput) nameInput.value = '';
        renderTable();
    });
    on('btnUndo', 'click', undo);
    on('btnRedo', 'click', redo);
    on('btnSaveAll', 'click', () => { saveToStorage(); showToast('Gespeichert'); });

    // Cloud sync buttons
    on('btnSubmitReview', 'click', handleSubmitForReview);
    on('btnReviewPanel', 'click', () => { openReviewPanel(); });
    on('btnCloseReview', 'click', () => { $('reviewModal').style.display = 'none'; });
    on('reviewStatusClose', 'click', () => { $('reviewStatusBar').style.display = 'none'; });

    // Review tab filters
    document.querySelectorAll('.review-tab').forEach(tab => {
        tab.addEventListener('click', () => {
            document.querySelectorAll('.review-tab').forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            const filter = tab.dataset.filter;
            renderReviewListFiltered(filter);
        });
    });
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
    on('btnToggleSidebar', 'click', toggleMapSidebar);
    on('btnMapStandard', 'click', () => switchLayer('standard'));
    on('btnMapSatellite', 'click', () => switchLayer('satellite'));
    on('btnUploadPlan', 'click', () => $('inputPlanImage').click());
    on('inputPlanImage', 'change', handlePlanUpload);
    on('btnGoToLocation', 'click', handleMapSearch);
    on('inpMapSearch', 'keypress', (e) => { if (e.key === 'Enter') handleMapSearch(); });
    on('btnLocateMe', 'click', startGPS);
    on('btnAutoGPS', 'click', () => {
        AppState.autoGPS = !AppState.autoGPS;
        const btn = $('btnAutoGPS');
        if (AppState.autoGPS) {
            btn.classList.add('active-tool');
            btn.style.background = 'rgba(16, 185, 129, 0.2)';
            btn.style.borderColor = '#10b981';
            btn.style.color = '#10b981';
            // Also start GPS if not already running
            if (!AppState.userMarker) startGPS();
            showToast('🛰️ Auto-GPS aktiv — Koordinaten werden automatisch eingetragen');
        } else {
            btn.classList.remove('active-tool');
            btn.style.background = '';
            btn.style.borderColor = '';
            btn.style.color = '';
            showToast('Auto-GPS deaktiviert');
        }
    });
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
    on('chkFilter08', 'change', () => toggleLayerColor('#e879f9', 'chkFilter08'));
    on('chkFilter16', 'change', () => toggleLayerColor('#fbbf24', 'chkFilter16'));
    on('chkFilter32', 'change', () => toggleLayerColor('#22d3ee', 'chkFilter32'));

    // Numbered markers
    document.querySelectorAll('.num-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            activateNumberedPlacement(btn.dataset.num);
        });
    });

    // Snipping
    on('btnAutoSnip', 'click', () => triggerAusschnitt(null));
    on('btnClearShapes', 'click', () => {
        if (layers.drawItems) layers.drawItems.clearLayers();
        // Only clear freehand drawings — stars are deleted individually by clicking them
        saveMapData();
        showToast('Freihand-Objekte gelöscht');
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
            let projectName = '';
            if (pInput) {
                pInput.value = file.name.split(/[._\s-]/)[0];
                projectName = pInput.value;
            } else {
                projectName = file.name.split(/[._\s-]/)[0];
            }

            // Check if project exists in library
            let library = JSON.parse(localStorage.getItem('messstellen_library') || '{}');
            const existingProject = library[projectName];
            if (existingProject) {
                if (existingProject.hiddenMapColors) {
                    AppState.hiddenMapColors = new Set(existingProject.hiddenMapColors);
                }
                if (existingProject.newCols) {
                    AppState.newCols = new Set(existingProject.newCols);
                }
                
                // Merge coordinates and custom fields
                if (existingProject.data && existingProject.data.length > 0) {
                    const existingRowMap = new Map();
                    existingProject.data.forEach(row => {
                        if (row['Kennzeichen']) {
                            existingRowMap.set(row['Kennzeichen'].toString().trim(), row);
                        }
                    });
                    
                    AppState.data.forEach(row => {
                        const kz = row['Kennzeichen'] ? row['Kennzeichen'].toString().trim() : '';
                        if (kz && existingRowMap.has(kz)) {
                            const exRow = existingRowMap.get(kz);
                            if (!row['Koordinaten'] && exRow['Koordinaten']) row['Koordinaten'] = exRow['Koordinaten'];
                            if (!row['GPS-Lat'] && exRow['GPS-Lat']) row['GPS-Lat'] = exRow['GPS-Lat'];
                            if (!row['GPS-Lng'] && exRow['GPS-Lng']) row['GPS-Lng'] = exRow['GPS-Lng'];
                            
                            Object.keys(exRow).forEach(key => {
                                if (row[key] === undefined || row[key] === '') {
                                    row[key] = exRow[key];
                                }
                            });
                        }
                    });
                }
                saveToStorage();
                if (map) restoreMapDrawings();
                showToast(`Projekt "${projectName}" verknüpft — Zeichnungen & Koordinaten geladen`);
            } else {
                showToast(`${mapped.length} Messstellen geladen`);
            }
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

    // Auto-add 100 empty rows if table is completely empty (Messhelfer starting fresh)
    if (AppState.data.length === 0 && $('mainApp') && $('mainApp').style.display !== 'none') {
        for (let i = 0; i < 100; i++) AppState.data.push({});
    }

    // Build header
    thead.innerHTML = '';
    const tr1 = document.createElement('tr');
    const tr2 = document.createElement('tr');
    tr1.innerHTML = '<th rowspan="2" style="width:30px;text-align:center;">#</th>';

    TABLE_STRUCTURE.forEach(g => {
        if (AppState.hiddenColumns.has(g.class)) return;
        const groupCols = g.columns.filter(c => !AppState.hiddenColumns.has(c));
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
            const isAnhang = c.includes('Anhang');
            const defaultWidth = isAnhang ? '56px' : '36px';
            sh.style.minWidth = AppState.columnWidths[c] || defaultWidth;
            sh.style.color = cs.text;
            // First column of each group gets colored left border
            if (cIdx === 0) {
                sh.style.borderLeft = `3px solid ${cs.border}`;
            }
            // ALL sub-header cells get bottom border in group color
            sh.style.borderBottom = `2px solid ${cs.border}`;
            
            let label = c.includes('_') ? c.split('_')[0] : c;
            label = label.replace(' [Ω]', '').replace(' [Ωm]', ''); // Remove units to save space
            if (c === 'Sprachsteuerung') {
                sh.innerHTML = '<i class="fas fa-microphone" title="Sprachsteuerung" style="cursor:help;"></i>';
                sh.style.width = '36px';
                sh.style.minWidth = '36px';
            } else {
                sh.textContent = label;
            }

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

        const numTd = document.createElement('td');
        numTd.style.cssText = 'text-align:center;font-weight:700;color:var(--text-muted);width:30px;white-space:nowrap;';
        numTd.textContent = idx + 1;
        tr.appendChild(numTd);

        TABLE_STRUCTURE.forEach(g => {
            if (AppState.hiddenColumns.has(g.class)) return;
            const cs = DEPTH_COLORS[g.class] || DEPTH_COLORS.basis;
            const groupCols = g.columns.filter(c => !AppState.hiddenColumns.has(c));

            groupCols.forEach((col, colIdx) => {
                const td = document.createElement('td');

                // First cell of each group gets a colored left border
                if (colIdx === 0) {
                    td.style.borderLeft = `3px solid ${cs.border}`;
                }

                if (col.toLowerCase().includes('anhang') || col.toLowerCase().includes('bilder')) {
                    td.style.textAlign = 'center';
                    const isBilder = col.toLowerCase().includes('bilder');
                    
                    if (originalRow[col]) {
                        const imgContainer = document.createElement('div');
                        imgContainer.className = 'table-img-container';

                        const imgPreview = document.createElement('img');
                        imgPreview.src = originalRow[col];
                        imgPreview.className = 'table-img-preview';
                        imgPreview.title = isBilder ? 'Foto bearbeiten' : 'Ansehen';
                        imgPreview.onclick = () => {
                            if (isBilder) openImageEditor(originalRow[col], idx, col);
                            else openImagePreview(originalRow[col]);
                        };
                        imgContainer.appendChild(imgPreview);

                        const delBtn = document.createElement('button');
                        delBtn.className = 'toolbar-btn';
                        delBtn.style.color = 'var(--danger)';
                        delBtn.innerHTML = '<i class="fas fa-trash"></i>';
                        delBtn.title = 'Löschen';
                        delBtn.onclick = () => {
                            if (confirm('Foto löschen?')) {
                                delete originalRow[col];
                                renderTable();
                                saveToStorage();
                            }
                        };
                        imgContainer.appendChild(delBtn);
                        td.appendChild(imgContainer);
                    } else {
                        if (isBilder) {
                            const camBtn = document.createElement('button');
                            camBtn.className = 'toolbar-btn';
                            camBtn.innerHTML = '<i class="fas fa-camera"></i>';
                            camBtn.title = 'Foto aufnehmen';
                            camBtn.onclick = () => triggerCamera(idx, col);
                            td.appendChild(camBtn);
                        } else {
                            td.textContent = '-';
                            td.style.color = 'var(--text-muted)';
                        }
                    }
                } else if (col === 'Sprachsteuerung') {
                    td.style.textAlign = 'center';
                    const micBtn = document.createElement('button');
                    micBtn.className = 'row-mic-btn';
                    micBtn.dataset.row = idx;
                    micBtn.title = 'Zeilen-Diktat (0.8m & 1.6m)';
                    micBtn.style.cssText = 'background:transparent;border:none;color:var(--text-muted);cursor:pointer;font-size:13px;padding:2px 4px;';
                    micBtn.innerHTML = '<i class="fas fa-microphone"></i>';
                    micBtn.addEventListener('click', (e) => {
                        e.stopPropagation();
                        if (typeof startRowVoiceDictation === 'function') {
                            startRowVoiceDictation(idx, micBtn);
                        }
                    });
                    td.appendChild(micBtn);
                } else {
                    const inp = document.createElement('input');
                    inp.type = 'text';
                    let val = originalRow[col] || '';
                    if (col === 'Koordinaten' && val) {
                        const parsed = parseCoordinates(String(val));
                        if (parsed) {
                            val = convertToDMS(parsed.lat, true) + ' ' + convertToDMS(parsed.lng, false);
                            originalRow[col] = val;
                        }
                    }
                    inp.value = val;
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
                        if (col === 'Koordinaten') {
                            syncMapWithTable();
                        }
                        debouncedSave(800);
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
    
    if (typeof syncMapWithTable === 'function') syncMapWithTable();
}

function syncMapWithTable() {
    if (!map) return;
    if (!layers.tableMarkers) {
        layers.tableMarkers = L.featureGroup().addTo(map);
    }
    layers.tableMarkers.clearLayers();

    AppState.data.forEach((row, idx) => {
        let lat = null, lng = null;
        if (row['Koordinaten']) {
            const coordStr = String(row['Koordinaten']);
            const parsed = parseCoordinates(coordStr);
            if (parsed) {
                lat = parsed.lat;
                lng = parsed.lng;
                row['GPS-Lat'] = lat.toFixed(5);
                row['GPS-Lng'] = lng.toFixed(5);
            }
        }
        if ((lat === null || lng === null || isNaN(lat) || isNaN(lng)) && row['GPS-Lat'] && row['GPS-Lng']) {
            const latStr = String(row['GPS-Lat']);
            const lngStr = String(row['GPS-Lng']);
            lat = parseFloat(latStr.replace(',', '.'));
            lng = parseFloat(lngStr.replace(',', '.'));
        }
        if (lat !== null && lng !== null && !isNaN(lat) && !isNaN(lng)) {
            
            var color = '#3b82f6'; // Blue pin
            var marker = L.marker([lat, lng], {
                interactive: true,
                draggable: true,
                icon: L.divIcon({
                    html: `<div style="background:${color};width:14px;height:14px;border-radius:50%;border:2px solid #fff;box-shadow:0 2px 6px rgba(0,0,0,0.4);"></div>`,
                    iconSize: [14, 14], iconAnchor: [7, 7], className: 'table-synced-marker'
                })
            });

            marker.on('dragend', function(e) {
                const newLatLng = e.target.getLatLng();
                const newLat = newLatLng.lat;
                const newLng = newLatLng.lng;
                const dmsVal = convertToDMS(newLat, true) + ' ' + convertToDMS(newLng, false);
                
                row['Koordinaten'] = dmsVal;
                row['GPS-Lat'] = newLat.toFixed(5);
                row['GPS-Lng'] = newLng.toFixed(5);
                
                debouncedSave(800);
                setTimeout(() => {
                    renderTable();
                }, 100);
                showToast(`Messstelle verschoben auf: ${dmsVal}`);
            });
            var label = 'Messstelle';
            if (!AppState.hiddenColumns.has('Kennzeichen') && row['Kennzeichen']) {
                label = row['Kennzeichen'];
            } else if (!AppState.hiddenColumns.has('Alt-Kz.') && row['Alt-Kz.']) {
                label = row['Alt-Kz.'];
            }
            var mapsLink = `https://www.google.com/maps/search/?api=1&query=${lat},${lng}`;
            
            var dmsFormat = convertToDMS(lat, true) + ' ' + convertToDMS(lng, false);

            marker.bindTooltip(`${label}`, { direction: 'top', offset: [0, -10] });
            marker.bindPopup(`
                <div style="font-family:'Inter',sans-serif;font-size:13px;color:#334155;">
                    <strong style="font-size:14px;color:#0f172a;">#${idx + 1} - ${label}</strong>
                    <div style="margin:8px 0;padding:6px;background:#f1f5f9;border-radius:4px;font-family:monospace;font-size:12px;text-align:center;border:1px solid #cbd5e1;">
                        <span style="font-weight:bold;font-size:13px;color:#1e293b;">${dmsFormat}</span><br>
                        <span style="color:#64748b;font-size:11px;margin-top:2px;display:block;">${lat.toFixed(6)}, ${lng.toFixed(6)}</span>
                    </div>
                    <a href="${mapsLink}" target="_blank" style="display:block;text-align:center;background:#3b82f6;color:white;text-decoration:none;padding:6px 10px;border-radius:4px;font-weight:600;">
                        <i class="fas fa-external-link-alt"></i> In Google Maps öffnen
                    </a>
                </div>
            `);
            layers.tableMarkers.addLayer(marker);
        }
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
            try {
                layers.drawItems.eachLayer(l => {
                    if (!l.feature) l.feature = { type: 'Feature', properties: {} };
                    if (l.options && l.options.markerColor) l.feature.properties.markerColor = l.options.markerColor;
                    if (!l.feature.properties.markerColor && l.options && l.options.color) l.feature.properties.markerColor = l.options.color;
                    if (l.options && l.options.isNumbered) {
                        l.feature.properties.isNumbered = true;
                        l.feature.properties.markerNumber = l.options.markerNumber;
                    }
                });
            } catch (layerErr) {
                console.warn('Map layer serialization warning:', layerErr);
            }
        }

        // Collect star layer data for cross-device restore
        const starData = [];
        ['star08', 'star16', 'star32'].forEach(function(layerKey) {
            if (!layers[layerKey]) return;
            const color = layerKey === 'star08' ? '#e879f9' : (layerKey === 'star16' ? '#fbbf24' : '#22d3ee');
            const depthKey = layerKey === 'star08' ? '0.8' : (layerKey === 'star16' ? '1.6' : '3.2');
            layers[layerKey].eachLayer(function(group) {
                // Find the center dot (circleMarker) to get the center coords
                var centerLatLng = null;
                group.eachLayer(function(l) {
                    if (l instanceof L.CircleMarker && l.getRadius && l.getRadius() === 4) {
                        centerLatLng = l.getLatLng();
                    }
                });
                if (centerLatLng) {
                    starData.push({ lat: centerLatLng.lat, lng: centerLatLng.lng, depth: depthKey, color: color });
                }
            });
        });

        const project = {
            name: pName,
            data: AppState.data,
            mapData: layers.drawItems ? layers.drawItems.toGeoJSON() : null,
            starData: starData,
            timestamp: Date.now(),
            hiddenMapColors: Array.from(AppState.hiddenMapColors),
            newCols: Array.from(AppState.newCols)
        };

        var libraryStr = localStorage.getItem('messstellen_library');
        var library;
        try {
            library = libraryStr ? JSON.parse(libraryStr) : {};
        } catch (parseErr) {
            console.error('localStorage corrupted, recovering:', parseErr);
            // Attempt recovery: save current data to a fresh library
            library = {};
            showToast('Datenbank repariert — alte Projekte könnten verloren sein');
        }

        library[pName] = project;
        
        try {
            localStorage.setItem('messstellen_library', JSON.stringify(library));
        } catch (quotaErr) {
            if (quotaErr.name === 'QuotaExceededError' || quotaErr.code === 22) {
                // Try to free space by removing oldest projects
                var names = Object.keys(library).filter(function(n) { return n !== pName; });
                names.sort(function(a, b) { return (library[a].timestamp || 0) - (library[b].timestamp || 0); });
                while (names.length > 0) {
                    var oldest = names.shift();
                    delete library[oldest];
                    try {
                        localStorage.setItem('messstellen_library', JSON.stringify(library));
                        showToast('Speicher voll — Projekt "' + oldest + '" gelöscht');
                        break;
                    } catch (e2) { continue; }
                }
                if (names.length === 0) {
                    showToast('Speicher voll! Bitte große Bilder entfernen.');
                }
            } else {
                throw quotaErr;
            }
        }
        
        localStorage.setItem('current_project_id', pName);

        // ── Cloud Sync ──
        if (typeof isFirebaseConfigured === 'function' && isFirebaseConfigured()) {
            saveToCloud(pName, project).catch(err => {
                console.warn('Cloud sync failed (will retry):', err);
            });
        }
    } catch (err) {
        console.error('Save error:', err);
        showToast('Speichern fehlgeschlagen: ' + err.message);
    }
}

function saveMapData() {
    saveToStorage();
}

function loadFromStorage() {
    const pName = localStorage.getItem('current_project_id');
    if (!pName) return;

    var libraryStr = localStorage.getItem('messstellen_library');
    var library;
    try {
        library = libraryStr ? JSON.parse(libraryStr) : {};
    } catch (parseErr) {
        console.error('localStorage corrupted on load:', parseErr);
        showToast('Datenbank beschädigt — bitte Daten neu laden');
        return;
    }
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
        // First invalidate after initial render
        setTimeout(() => {
            if (map) {
                map.invalidateSize();
                refreshMapVisibility();
            }
        }, 350);
        // Second invalidate to catch slow CSS transitions on tablets
        setTimeout(() => { if (map) map.invalidateSize(); }, 800);
    }
}

function closeMap() {
    const m = $('mapLayout');
    if (m) { m.style.display = 'none'; m.classList.remove('show'); }
    if (typeof switchTab === 'function') switchTab('table');
}

function toggleMapSidebar() {
    const sidebar = $('mapSidebar');
    const btn = $('btnToggleSidebar');
    if (!sidebar) return;
    const isCollapsed = sidebar.classList.toggle('collapsed');
    if (btn) {
        btn.querySelector('i').className = isCollapsed ? 'fas fa-chevron-left' : 'fas fa-chevron-right';
    }
    // Let Leaflet recalculate size after animation
    setTimeout(() => { if (map) map.invalidateSize(); }, 320);
}

function initMap() {
    if (map) return;
    const mapContainer = $('map');
    if (!mapContainer) return;

    map = L.map('map', {
        zoomControl: false,
        attributionControl: true,
        doubleClickZoom: false,
        preferCanvas: true,
        zoomSnap: 0.25,
        zoomDelta: 0.5,
        wheelPxPerZoomLevel: 100
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

// Tile layers – use OpenStreetMap for CORS‑friendly screenshots
const USE_OSM = false;
if (USE_OSM) {
    // Standard OSM road map
    layers.standard = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        maxZoom: 22,
        attribution: '&copy; OpenStreetMap contributors'
    });
    // OpenTopoMap as a satellite‑like overlay (free, CORS‑allowed)
    layers.satellite = L.tileLayer('https://{s}.tile.opentopomap.org/{z}/{x}/{y}.png', {
        maxZoom: 22,
        attribution: '&copy; OpenTopoMap contributors'
    });
} else {
    // Fallback to original Google Maps (may cause CORS issues)
    layers.standard = L.tileLayer('https://{s}.google.com/vt/lyrs=m&x={x}&y={y}&z={z}', {
        maxZoom: 22,
        maxNativeZoom: 20,
        subdomains: ['mt0', 'mt1', 'mt2', 'mt3'],
        attribution: '&copy; Google Maps'
    });
    layers.satellite = L.tileLayer('https://{s}.google.com/vt/lyrs=y&x={x}&y={y}&z={z}', {
        maxZoom: 22,
        maxNativeZoom: 20,
        subdomains: ['mt0', 'mt1', 'mt2', 'mt3'],
        attribution: '&copy; Google Maps'
    });
}

    layers.standard.addTo(map);
    layers.drawItems = new L.FeatureGroup().addTo(map);
    layers.star08 = new L.FeatureGroup().addTo(map);
    layers.star16 = new L.FeatureGroup().addTo(map);
    layers.star32 = new L.FeatureGroup().addTo(map);

    // Sync checkboxes with hiddenMapColors state
    const chk08 = $('chkFilter08');
    const chk16 = $('chkFilter16');
    const chk32 = $('chkFilter32');
    if (chk08) chk08.checked = !AppState.hiddenMapColors.has('#e879f9');
    if (chk16) chk16.checked = !AppState.hiddenMapColors.has('#fbbf24');
    if (chk32) chk32.checked = !AppState.hiddenMapColors.has('#22d3ee');

    // Fix: call invalidateSize after sidebar toggle to prevent coordinate offset
    const sidebar = document.querySelector('.map-sidebar');
    if (sidebar) {
        const obs = new ResizeObserver(() => { if (map) map.invalidateSize(); });
        obs.observe(sidebar);
    }

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

function createPinMarker(latlng, color, customText) {
    // Count from depthMarkers array (resets when "Löschen" is clicked)
    var depthKey = color === '#e879f9' ? '0.8' : (color === '#fbbf24' ? '1.6' : '3.2');
    if (!AppState.depthMarkers) AppState.depthMarkers = { '0.8': [], '1.6': [], '3.2': [] };
    var count = AppState.depthMarkers[depthKey].length;
    
    var textToShow = customText !== undefined ? customText : count;

    const xSvg = `<svg width="18" height="18" viewBox="0 0 18 18" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M4 4L14 14M14 4L4 14" stroke="black" stroke-width="3.5" stroke-linecap="round"/>
        <path d="M4 4L14 14M14 4L4 14" stroke="${color}" stroke-width="1.5" stroke-linecap="round"/>
    </svg>`;
    return L.marker(latlng, {
        draggable: true,
        interactive: true,
        zIndexOffset: 1000,
        markerColor: color,
        icon: L.divIcon({
            html: `<div style="display:flex; flex-direction:column; align-items:center; filter:drop-shadow(0 1px 2px rgba(0,0,0,0.3));">
                       <div style="width:18px;height:18px;">${xSvg}</div>
                       <div style="background:rgba(255,255,255,0.9); color:${color}; font-size:9px; font-weight:bold; padding:0 3px; border-radius:3px; margin-top:1px;">${textToShow}</div>
                   </div>`,
            iconSize: [24, 34], iconAnchor: [12, 18], className: 'custom-marker-icon'
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

    const parsed = parseCoordinates(query);

    if (parsed) {
        const lat = parsed.lat, lng = parsed.lng;
        map.setView([lat, lng], 19);
        const btnHtml = `<div style="margin-top: 8px; text-align: center;"><button class="toolbar-btn" style="background:#10b981; color:white; font-size:11px; padding:4px 8px; border-radius:4px; font-weight:600; width:100%;" onclick="applySearchCoords(${lat}, ${lng})"><i class="fas fa-check"></i> Untuk baris aktif</button></div>`;
        const dmsVal = convertToDMS(lat, true) + ' ' + convertToDMS(lng, false);
        searchMarker = L.marker([lat, lng]).addTo(map).bindPopup(`<strong>Koordinaten</strong><br><span style="font-size:11px;color:#64748b;">${dmsVal}</span>${btnHtml}`).openPopup();
        showToast(`Koordinaten: ${dmsVal}`);
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
            const btnHtml = `<div style="margin-top: 8px; text-align: center;"><button class="toolbar-btn" style="background:#10b981; color:white; font-size:11px; padding:4px 8px; border-radius:4px; font-weight:600; width:100%;" onclick="applySearchCoords(${res.lat}, ${res.lon})"><i class="fas fa-check"></i> Untuk baris aktif</button></div>`;
            searchMarker = L.marker([res.lat, res.lon]).addTo(map)
                .bindPopup(`<strong style="font-size:13px;">${res.display_name.split(',')[0]}</strong><br><span style="font-size:11px;color:#64748b;display:block;max-width:200px;word-break:break-word;">${res.display_name}</span>${btnHtml}`)
                .openPopup();
            showToast(`Gefunden: ${res.display_name.split(',')[0]}`);
        } else {
            showToast('Standort nicht gefunden');
        }
    } catch (err) {
        $('loadingOverlay').style.display = 'none';
        showToast('Suche fehlgeschlagen: ' + err.message);
    }
}

// ─── HELPER: Find first row without coordinates ───
function findTargetRowIdx() {
    if (AppState.selectedCell) {
        return parseInt(AppState.selectedCell.split('-')[0]);
    }
    // Find first row that has no coordinates yet
    if (AppState.data && AppState.data.length > 0) {
        const emptyIdx = AppState.data.findIndex(row => !row['Koordinaten'] && !row['GPS-Lat']);
        if (emptyIdx >= 0) return emptyIdx;
        // All rows have coords — use last row
        return AppState.data.length - 1;
    }
    return -1;
}

window.applySearchCoords = function(lat, lon) {
    const latNum = parseFloat(lat);
    const lonNum = parseFloat(lon);
    let rowIdx = findTargetRowIdx();
    
    if (rowIdx >= 0 && AppState.data[rowIdx]) {
        const dmsVal = convertToDMS(latNum, true) + ' ' + convertToDMS(lonNum, false);
        AppState.data[rowIdx]['Koordinaten'] = dmsVal;
        AppState.data[rowIdx]['GPS-Lat'] = latNum.toFixed(5);
        AppState.data[rowIdx]['GPS-Lng'] = lonNum.toFixed(5);
        AppState.hiddenColumns.delete('special');
        renderTable();
        saveToStorage();
        if (searchMarker && map) map.removeLayer(searchMarker);
        showToast('Koordinaten übernommen');
    } else {
        showToast('Keine aktive Zeile ausgewählt');
    }
};

function startGPS() {
    if (!map) return;
    showToast('GPS-Suche aktiv...');
    AppState.firstLocationFound = false;
    AppState._lastGPSUpdate = Date.now();
    AppState._gpsBuffer = []; // Reset position averaging buffer
    map.locate({
        setView: false,
        maxZoom: 18,
        enableHighAccuracy: true,
        watch: true,
        timeout: 15000,
        maximumAge: 0
    });
    initCompass();

    // GPS watchdog — restart if no update for 60 seconds
    if (AppState._gpsWatchdog) clearInterval(AppState._gpsWatchdog);
    AppState._gpsWatchdog = setInterval(function() {
        if (!AppState._lastGPSUpdate) return;
        var elapsed = Date.now() - AppState._lastGPSUpdate;
        var statusDot = $('gpsAccuracyDot');
        var statusText = $('gpsStatusText');

        if (elapsed > 60000) {
            // GPS lost — restart
            if (statusDot) {
                statusDot.classList.remove('pulse');
                statusDot.classList.add('lost');
            }
            if (statusText) statusText.textContent = 'GPS verloren — Neustart...';
            
            // Stop and restart location tracking
            try {
                map.stopLocate();
                setTimeout(function() {
                    map.locate({ 
                        setView: false, maxZoom: 18, 
                        enableHighAccuracy: true, 
                        watch: true, timeout: 15000, maximumAge: 0 
                    });
                    AppState._lastGPSUpdate = Date.now();
                    showToast('GPS wird neu gestartet...');
                }, 2000);
            } catch (e) {
                console.warn('GPS restart error:', e);
            }
        } else if (elapsed > 30000) {
            // GPS weak
            if (statusDot) statusDot.classList.add('lost');
            if (statusText) statusText.textContent = 'GPS-Signal schwach...';
        }
    }, 10000);
}

function handleLocationFound(e) {
    const rawLatlng = e.latlng;
    const acc = Math.round(e.accuracy);
    AppState._lastGPSUpdate = Date.now();

    const statusText = $('gpsStatusText');
    const statusDot = $('gpsAccuracyDot');

    // ── Position averaging: keep last 10 readings, filter bad ones, squared-weight by accuracy ──
    if (!AppState._gpsBuffer) AppState._gpsBuffer = [];
    
    // Only reject truly unusable readings (> 200m)
    // The squared-weighting will still heavily prefer accurate ones
    if (acc < 200) {
        AppState._gpsBuffer.push({ lat: rawLatlng.lat, lng: rawLatlng.lng, acc, ts: Date.now() });
        // Keep last 10 good readings
        if (AppState._gpsBuffer.length > 10) AppState._gpsBuffer.shift();
        
        // Remove readings older than 30 seconds (stale data)
        const now = Date.now();
        AppState._gpsBuffer = AppState._gpsBuffer.filter(p => (now - p.ts) < 30000);
    }

    let latlng;
    let effectiveAcc = acc;

    if (AppState._gpsBuffer.length >= 2) {
        // Squared-weight: much heavier preference for accurate readings
        // A reading with 5m accuracy gets 16x more weight than one with 20m
        const totalWeight = AppState._gpsBuffer.reduce((sum, p) => sum + (1 / (p.acc * p.acc)), 0);
        const avgLat = AppState._gpsBuffer.reduce((sum, p) => sum + p.lat * (1 / (p.acc * p.acc)), 0) / totalWeight;
        const avgLng = AppState._gpsBuffer.reduce((sum, p) => sum + p.lng * (1 / (p.acc * p.acc)), 0) / totalWeight;
        latlng = L.latLng(avgLat, avgLng);
        // Effective accuracy = best reading in buffer (since we weight heavily toward it)
        effectiveAcc = Math.min(...AppState._gpsBuffer.map(p => p.acc));
    } else {
        latlng = L.latLng(rawLatlng.lat, rawLatlng.lng);
    }

    // ── GPS STABILIZATION: Don't move marker if within dead-zone ──
    // Prevents jitter/drift when standing still
    const DEAD_ZONE_M = 3; // Minimum movement in meters before marker updates
    if (AppState._lastStableLatLng) {
        const drift = map.distance(latlng, AppState._lastStableLatLng);
        if (drift < DEAD_ZONE_M && effectiveAcc <= 30) {
            // Position hasn't changed significantly — keep the old stable position
            // but still update the accuracy circle and status
            latlng = AppState._lastStableLatLng;
        } else {
            // Significant movement — update stable position
            AppState._lastStableLatLng = latlng;
        }
    } else {
        AppState._lastStableLatLng = latlng;
    }

    // Reuse existing marker instead of destroy+recreate (smoother)
    if (AppState.userMarker) {
        AppState.userMarker.setLatLng(latlng);
    } else {
        const userIcon = L.divIcon({
            html: `<div style="background:var(--accent,#3b82f6);width:28px;height:28px;border-radius:50%;border:3px solid white;box-shadow:0 2px 8px rgba(0,0,0,0.4);display:flex;align-items:center;justify-content:center;"><i class="fas fa-user" style="color:white;font-size:12px;"></i></div>`,
            iconSize: [28, 28], iconAnchor: [14, 14], className: 'user-icon'
        });
        AppState.userMarker = L.marker(latlng, { icon: userIcon, zIndexOffset: 2000 }).addTo(map);
    }

    if (AppState.accuracyCircle) {
        AppState.accuracyCircle.setLatLng(latlng);
        AppState.accuracyCircle.setRadius(effectiveAcc);
    } else {
        AppState.accuracyCircle = L.circle(latlng, effectiveAcc, {
            color: '#3b82f6', fillColor: 'rgba(59,130,246,0.1)', weight: 1
        }).addTo(map);
    }

    if (!AppState.firstLocationFound) {
        map.panTo(latlng);
        AppState.firstLocationFound = true;
    } else if (AppState.liveFollow) {
        map.panTo(latlng);
    }

    // Update GPS status with buffer info
    const bufLen = AppState._gpsBuffer.length;
    const qualityLabel = effectiveAcc < 10 ? '✓✓ Exzellent' : (effectiveAcc < 20 ? '✓ Sehr gut' : (effectiveAcc < 50 ? '✓ Gut' : effectiveAcc < 100 ? 'OK' : 'Schwach'));
    if (statusText) statusText.textContent = `GPS: ±${effectiveAcc}m ${qualityLabel} (${bufLen}/10)`;
    if (statusDot) {
        statusDot.style.background = effectiveAcc < 10 ? '#10b981' : (effectiveAcc < 20 ? '#22d3ee' : (effectiveAcc < 50 ? '#f59e0b' : '#ef4444'));
        statusDot.classList.add('pulse');
    }

    // Distance calculation
    if (AppState.startPoint) {
        const dist = map.distance(latlng, AppState.startPoint);
        const badge = $('distanceBadge');
        const val = $('distanceValue');
        if (badge && val) { val.textContent = dist.toFixed(1); badge.style.display = 'flex'; }
    }

    // ── AUTO-GPS: Automatically fill coordinates into next empty row ──
    if (AppState.autoGPS && acc <= 20 && AppState.data && AppState.data.length > 0) {
        const emptyIdx = AppState.data.findIndex(row => !row['Koordinaten'] && !row['GPS-Lat']);
        if (emptyIdx >= 0) {
            const dmsVal = convertToDMS(avgLat, true) + ' ' + convertToDMS(avgLng, false);
            AppState.data[emptyIdx]['Koordinaten'] = dmsVal;
            AppState.data[emptyIdx]['GPS-Lat'] = avgLat.toFixed(5);
            AppState.data[emptyIdx]['GPS-Lng'] = avgLng.toFixed(5);
            AppState.hiddenColumns.delete('special');
            renderTable();
            saveToStorage();
            showToast(`✅ Zeile ${emptyIdx + 1} — GPS automatisch übernommen (±${acc}m)`);
            // After filling, turn off auto-GPS so it doesn't overwrite next row immediately
            AppState.autoGPS = false;
            const autoBtn = $('btnAutoGPS');
            if (autoBtn) {
                autoBtn.classList.remove('active-tool');
                autoBtn.style.background = '';
            }
        }
    }
}

function handleLocationError(e) {
    const messages = {
        1: 'Standort-Berechtigung verweigert',
        2: 'Position nicht verfügbar',
        3: 'Zeitüberschreitung — wird erneut versucht'
    };
    
    // Silence timeout errors (code 3) to prevent annoying the field workers with constant toasts
    if (e.code !== 3) {
        showToast('GPS: ' + (messages[e.code] || e.message));
    }

    const statusDot = $('gpsAccuracyDot');
    const statusText = $('gpsStatusText');
    if (statusDot) {
        statusDot.style.background = '#ef4444'; // Red for no signal
        statusDot.classList.remove('pulse');
    }
    if (statusText) statusText.textContent = 'GPS: Kein Signal';

    // Auto-retry on timeout errors (code 3)
    if (e.code === 3 && map) {
        setTimeout(function() {
            try {
                map.stopLocate();
                map.locate({ 
                    setView: false, maxZoom: 18, 
                    enableHighAccuracy: true, 
                    watch: true, timeout: 20000, maximumAge: 0 
                });
                AppState._lastGPSUpdate = Date.now();
            } catch (err) {
                console.warn('GPS retry failed:', err);
            }
        }, 3000);
    }
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

    // Commented out to hide the red start marker dot as requested
    // if (AppState.startMarker) map.removeLayer(AppState.startMarker);
    // AppState.startMarker = L.marker(latlng, {
    //     icon: L.divIcon({
    //         html: '<div style="background:#ef4444;width:20px;height:20px;border-radius:50%;border:3px solid white;box-shadow:0 2px 8px rgba(0,0,0,0.4);"></div>',
    //         iconSize: [20, 20], iconAnchor: [10, 10], className: 'start-marker'
    //     }),
    //     zIndexOffset: 3000
    // }).addTo(map);
    
    // Zoom to current position
    map.setView(latlng, 20, { animate: true });

    // Auto-save coordinates to the active table row
    // Auto-save coordinates to the active table row
    var lat = latlng.lat;
    var lng = latlng.lng;
    var rowIdx = -1;
    if (AppState.selectedCell) {
        rowIdx = parseInt(AppState.selectedCell.split('-')[0]);
    } else if (AppState.data && AppState.data.length > 0) {
        rowIdx = AppState.data.length - 1;
    }

    // If no rows exist yet, create one
    if (rowIdx < 0 || !AppState.data[rowIdx]) {
        if (!AppState.data) AppState.data = [];
        AppState.data.push({});
        rowIdx = AppState.data.length - 1;
    }

    if (rowIdx >= 0 && AppState.data[rowIdx]) {
        const dmsVal = convertToDMS(lat, true) + ' ' + convertToDMS(lng, false);
        AppState.data[rowIdx]['Koordinaten'] = dmsVal;
        AppState.data[rowIdx]['GPS-Lat'] = lat.toFixed(5);
        AppState.data[rowIdx]['GPS-Lng'] = lng.toFixed(5);
        AppState.hiddenColumns.delete('special');
        if (typeof renderTable === 'function') renderTable();
        if (typeof saveToStorage === 'function') saveToStorage();
        showToast('Startpunkt gesetzt & in Tabelle gespeichert ✓');
    } else {
        showToast('Startpunkt an GPS-Position gesetzt');
    }
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
                html: '<div style="color:' + d.color + ';font-size:14px;font-weight:900;white-space:nowrap;text-shadow:1px 1px 2px #000, -1px -1px 2px #000, 1px -1px 2px #000, -1px 1px 2px #000;">' + d.label + '</div>',
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
    AppState.activeDrawToolName = 'sticky-marker';
    AppState.activeColor = color;
    AppState.activeDepthLabel = depthLabel;
    $('map').style.cursor = 'crosshair';

    // Clear any old preview circles
    if (AppState.measureCircles) {
        AppState.measureCircles.forEach(function(c) { try { map.removeLayer(c); } catch(e){} });
    }
    AppState.measureCircles = [];

    map.on('click', handleStickyMarkerClick);
    showToast(depthLabel + ' — auf Karte tippen');
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
    var depthKey = depthLabel.replace('m', '');

    var centerLatLng = e.latlng;

    // Stop clicks while placing
    map.off('click', handleStickyMarkerClick);

    // Remove preview circle immediately — the star has its own dashed ring
    if (AppState.measureCircles && map) {
        AppState.measureCircles.forEach(function(c) { try { map.removeLayer(c); } catch(e){} });
        AppState.measureCircles = [];
    }

    // Auto-save coordinates to the active table row
    var lat = centerLatLng.lat;
    var lng = centerLatLng.lng;
    var rowIdx = -1;
    if (AppState.selectedCell) {
        rowIdx = parseInt(AppState.selectedCell.split('-')[0]);
    } else if (AppState.data && AppState.data.length > 0) {
        rowIdx = AppState.data.length - 1;
    }

    // If no rows exist yet, create one
    if (rowIdx < 0 || !AppState.data[rowIdx]) {
        if (!AppState.data) AppState.data = [];
        AppState.data.push({});
        rowIdx = AppState.data.length - 1;
    }

    if (rowIdx >= 0 && AppState.data[rowIdx]) {
        const dmsVal = convertToDMS(lat, true) + ' ' + convertToDMS(lng, false);
        AppState.data[rowIdx]['Koordinaten'] = dmsVal;
        AppState.data[rowIdx]['GPS-Lat'] = lat.toFixed(5);
        AppState.data[rowIdx]['GPS-Lng'] = lng.toFixed(5);
        AppState.hiddenColumns.delete('special');
        // Use setTimeout to avoid interfering with Leaflet's click event processing
        setTimeout(function() {
            if (typeof renderTable === 'function') renderTable();
            if (typeof saveToStorage === 'function') saveToStorage();
        }, 50);
    }

    if (!AppState.depthMarkers) AppState.depthMarkers = { '0.8': [], '1.6': [], '3.2': [] };
    var markers = AppState.depthMarkers[depthKey];
    var requiredDist = parseFloat(depthKey); // 0.8, 1.6, or 3.2 meters

    // Define bearings for Mercedes Star
    // 0.8m and 3.2m: 90°, 210°, 330°
    // 1.6m: 270°, 30°, 150° (Kebalikannya)
    var bearingsDeg = (depthKey === '1.6') ? [270, 30, 150] : [90, 210, 330];
        
        var groupLayer = L.featureGroup();
        groupLayer.options = { markerColor: color };

        // Optional: Draw a small center dot to indicate the origin (titik 0)
        var centerDot = L.circleMarker(centerLatLng, {
            radius: 4, color: color, weight: 2, fillColor: '#fff', fillOpacity: 1
        });
        groupLayer.addLayer(centerDot);

        var firstMarkerNum = markers.length + 1;

        // Add dashed circle around the star
        var outermostDist = requiredDist * 3;
        var circle = L.circle(centerLatLng, {
            radius: outermostDist,
            color: color, weight: 6, fillOpacity: 0, dashArray: '8, 8'
        });
        groupLayer.addLayer(circle);

        for (var i = 0; i < 3; i++) {
            var spokeNum = i + 1; // Garis bintang 1, 2, 3
            var bearingRad = bearingsDeg[i] * Math.PI / 180;
            
            // Jarak terjauh di ujung garis
            var endLatLng = destPoint(centerLatLng, bearingRad, outermostDist);
            
            // Marker 'X' di tiap jarak elektroda (a, 2a, 3a)
            for (var j = 0; j < 3; j++) {
                var currentDist = requiredDist * (j + 1);
                var pointLatLng = destPoint(centerLatLng, bearingRad, currentDist);
                markers.push(pointLatLng);
                
                var xMarker = L.marker(pointLatLng, {
                    interactive: false,
                    icon: L.divIcon({
                        html: `<div style="display:flex;align-items:center;justify-content:center;width:24px;height:24px;">
                                   <svg width="14" height="14" viewBox="0 0 18 18" fill="none" xmlns="http://www.w3.org/2000/svg">
                                       <path d="M4 4L14 14M14 4L4 14" stroke="black" stroke-width="3.5" stroke-linecap="round"/>
                                       <path d="M4 4L14 14M14 4L4 14" stroke="${color}" stroke-width="1.5" stroke-linecap="round"/>
                                   </svg>
                               </div>`,
                        iconSize: [24, 24], iconAnchor: [12, 12], className: 'x-marker'
                    })
                });
                groupLayer.addLayer(xMarker);
            }

            // Garis lurus putus-putus
            var line = L.polyline([centerLatLng, endLatLng], {
                color: color, weight: 6, opacity: 1.0, dashArray: '8, 8'
            });
            groupLayer.addLayer(line);

            // Spoke number label (1, 2, 3) OUTSIDE the circle
            var labelDist = outermostDist + 0.6; // +0.6 meters to sit just outside the dashed line
            var labelLatLng = destPoint(centerLatLng, bearingRad, labelDist);
            var distLabel = L.marker(labelLatLng, {
                interactive: false,
                isDistLabel: true,
                icon: L.divIcon({
                    html: `<div style="color:${color};font-size:22px;font-weight:900;white-space:nowrap;text-shadow:-2px -2px 0 #000, 2px -2px 0 #000, -2px 2px 0 #000, 2px 2px 0 #000, 0px 0px 8px #000;">` + spokeNum + `</div>`,
                    iconSize: [60, 26], iconAnchor: [30, 13], className: 'dist-label'
                })
            });
            groupLayer.addLayer(distLabel);
        }

        var targetLayer = (depthKey === '1.6') ? layers.star16 : ((depthKey === '3.2') ? layers.star32 : layers.star08);
        groupLayer.addTo(targetLayer);

        // Attach delete handler to the group AND every child layer for reliability
        var _starDeleting = false;
        var deleteHandler = function(ev) {
            L.DomEvent.stopPropagation(ev);
            L.DomEvent.preventDefault(ev);
            if (_starDeleting) return;
            _starDeleting = true;
            if (confirm('Mercedes Star ' + depthLabel + ' löschen?')) {
                targetLayer.removeLayer(groupLayer);
                if (AppState.depthMarkers[depthKey].length >= 9) {
                    AppState.depthMarkers[depthKey].splice(-9, 9);
                } else {
                    AppState.depthMarkers[depthKey] = [];
                }
                saveMapData();
            }
            setTimeout(function() { _starDeleting = false; }, 300);
        };
        groupLayer.on('click', deleteHandler);
        groupLayer.eachLayer(function(l) {
            if (l.on) l.on('click', deleteHandler);
        });

        saveMapData();

        // Re-enable for next star placement — but use a short delay to prevent
        // the current mouseup/click from immediately firing again
        setTimeout(function() {
            if (AppState.activeDrawToolName === 'sticky-marker') {
                map.on('click', handleStickyMarkerClick);
            }
        }, 300);
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
    const btn = $(btnId);
    if (!btn) return;
    
    const isHidden = !btn.checked;

    if (isHidden) {
        AppState.hiddenMapColors.add(colorLower);
    } else {
        AppState.hiddenMapColors.delete(colorLower);
    }

    // Hide/show the entire star group (circles, lines, X markers)
    let targetLayer = null;
    if (colorLower === '#e879f9') targetLayer = layers.star08;
    else if (colorLower === '#fbbf24') targetLayer = layers.star16;
    else if (colorLower === '#22d3ee') targetLayer = layers.star32;

    if (targetLayer && map) {
        if (isHidden) {
            map.removeLayer(targetLayer);
        } else {
            if (!map.hasLayer(targetLayer)) map.addLayer(targetLayer);
        }
    }

    // Also hide/show measure circles (showMeasureLine circles) of this color
    if (AppState.measureCircles && map) {
        AppState.measureCircles.forEach(function(c) {
            const cColor = (c.options && (c.options.color || c.options.fillColor)) || '';
            if (cColor.toLowerCase() === colorLower) {
                if (isHidden) {
                    if (map.hasLayer(c)) map.removeLayer(c);
                } else {
                    if (!map.hasLayer(c)) map.addLayer(c);
                }
            }
        });
    }

    // Also hide/show any freehand drawItems of this color
    if (layers.drawItems) {
        layers.drawItems.eachLayer(layer => {
            const c = (layer.options && (layer.options.markerColor || layer.options.color)) || '';
            if (c.toLowerCase() !== colorLower) return;
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

    saveToStorage();
};

function refreshMapVisibility() {
    if (!map) return;
    
    const depthLayers = [
        { color: '#e879f9', layer: layers.star08 },
        { color: '#fbbf24', layer: layers.star16 },
        { color: '#22d3ee', layer: layers.star32 }
    ];

    depthLayers.forEach(dl => {
        if (!dl.layer) return;
        if (AppState.hiddenMapColors.has(dl.color)) {
            map.removeLayer(dl.layer);
        } else {
            map.addLayer(dl.layer);
        }
    });

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

    // Restore freehand drawItems
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

    // Restore Mercedes Stars
    if (map && project && project.starData && project.starData.length > 0) {
        // Clear existing stars first
        if (layers.star08) layers.star08.clearLayers();
        if (layers.star16) layers.star16.clearLayers();
        if (layers.star32) layers.star32.clearLayers();
        AppState.depthMarkers = { '0.8': [], '1.6': [], '3.2': [] };

        project.starData.forEach(function(star) {
            var centerLatLng = L.latLng(star.lat, star.lng);
            var color = star.color;
            var depthKey = star.depth;
            var requiredDist = parseFloat(depthKey);
            var outermostDist = requiredDist * 3;
            var bearingsDeg = (depthKey === '1.6') ? [270, 30, 150] : [90, 210, 330];

            var groupLayer = L.featureGroup();
            groupLayer.options = { markerColor: color };

            // Center dot
            var centerDot = L.circleMarker(centerLatLng, {
                radius: 4, color: color, weight: 2, fillColor: '#fff', fillOpacity: 1
            });
            groupLayer.addLayer(centerDot);

            // Dashed outer circle
            var circle = L.circle(centerLatLng, {
                radius: outermostDist, color: color, weight: 6, fillOpacity: 0, dashArray: '8, 8'
            });
            groupLayer.addLayer(circle);

            for (var i = 0; i < 3; i++) {
                var spokeNum = i + 1;
                var bearingRad = bearingsDeg[i] * Math.PI / 180;
                var endLatLng = destPoint(centerLatLng, bearingRad, outermostDist);

                for (var j = 0; j < 3; j++) {
                    var currentDist = requiredDist * (j + 1);
                    var pointLatLng = destPoint(centerLatLng, bearingRad, currentDist);
                    AppState.depthMarkers[depthKey].push(pointLatLng);
                    var xMarker = L.marker(pointLatLng, {
                        interactive: false,
                        icon: L.divIcon({
                            html: `<div style="display:flex;align-items:center;justify-content:center;width:24px;height:24px;">
                                       <svg width="14" height="14" viewBox="0 0 18 18" fill="none" xmlns="http://www.w3.org/2000/svg">
                                           <path d="M4 4L14 14M14 4L4 14" stroke="black" stroke-width="3.5" stroke-linecap="round"/>
                                           <path d="M4 4L14 14M14 4L4 14" stroke="${color}" stroke-width="1.5" stroke-linecap="round"/>
                                       </svg>
                                   </div>`,
                            iconSize: [24, 24], iconAnchor: [12, 12], className: 'x-marker'
                        })
                    });
                    groupLayer.addLayer(xMarker);
                }

                var line = L.polyline([centerLatLng, endLatLng], {
                    color: color, weight: 3, opacity: 0.8, dashArray: '4, 4'
                });
                groupLayer.addLayer(line);

                // Spoke number label (1, 2, 3) OUTSIDE the circle
                var labelDist = outermostDist + 0.6; // +0.6 meters to sit just outside the dashed line
                var labelLatLng = destPoint(centerLatLng, bearingRad, labelDist);
                var distLabel = L.marker(labelLatLng, {
                    interactive: false,
                    icon: L.divIcon({
                        html: `<div style="color:${color};font-size:22px;font-weight:900;white-space:nowrap;text-shadow:-2px -2px 0 #000, 2px -2px 0 #000, -2px 2px 0 #000, 2px 2px 0 #000, 0px 0px 8px #000;">` + spokeNum + `</div>`,
                        iconSize: [60, 26], iconAnchor: [30, 13], className: 'dist-label'
                    })
                });
                groupLayer.addLayer(distLabel);
            }

            var targetLayer = (depthKey === '1.6') ? layers.star16 : ((depthKey === '3.2') ? layers.star32 : layers.star08);
            groupLayer.addTo(targetLayer);

            // Delete on click
            var _starDeleting = false;
            var deleteHandler = function(ev) {
                L.DomEvent.stopPropagation(ev);
                if (_starDeleting) return;
                _starDeleting = true;
                if (confirm('Mercedes Star ' + depthKey + 'm löschen?')) {
                    targetLayer.removeLayer(groupLayer);
                    saveMapData();
                }
                setTimeout(function() { _starDeleting = false; }, 300);
            };
            groupLayer.on('click', deleteHandler);
            groupLayer.eachLayer(function(l) { if (l.on) l.on('click', deleteHandler); });
        });

        refreshMapVisibility();
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

        await new Promise(r => setTimeout(r, 300));

        // ── Manual tile compositing (bypasses html2canvas CORS issues) ──
        const fullCanvas = await _captureMapToCanvas(mapEl);

        controls.forEach(c => c.style.display = '');

        if (!fullCanvas || fullCanvas.width === 0 || fullCanvas.height === 0) {
            throw new Error('Karte konnte nicht gerendert werden (Canvas leer).');
        }

        const cropped = targetLayer ? cropToShape(fullCanvas, targetLayer, mapEl) : fullCanvas;
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

/**
 * Manually composite visible Leaflet tile <img> elements and SVG/Canvas
 * overlays onto a fresh canvas. This avoids html2canvas entirely, which
 * cannot handle cross-origin Google Map tiles.
 */
async function _captureMapToCanvas(mapEl) {
    const scale = Math.max(2, window.devicePixelRatio || 2);
    const w = mapEl.offsetWidth;
    const h = mapEl.offsetHeight;
    const canvas = document.createElement('canvas');
    canvas.width = w * scale;
    canvas.height = h * scale;
    const ctx = canvas.getContext('2d');
    ctx.scale(scale, scale);
    ctx.fillStyle = '#1e293b';
    ctx.fillRect(0, 0, w, h);

    const mapRect = mapEl.getBoundingClientRect();

    // 1) Draw tile images (the base map layer)
    const tileImgs = mapEl.querySelectorAll('.leaflet-tile-pane img');
    const tilePromises = [];
    tileImgs.forEach(img => {
        if (!img.src || img.style.visibility === 'hidden' || img.style.display === 'none') return;
        if (parseFloat(getComputedStyle(img).opacity) === 0) return;

        tilePromises.push(
            _loadImageCORS(img.src).then(loaded => {
                const imgRect = img.getBoundingClientRect();
                const x = imgRect.left - mapRect.left;
                const y = imgRect.top - mapRect.top;
                ctx.globalAlpha = parseFloat(getComputedStyle(img).opacity) || 1;
                if (imgRect.width > 0 && imgRect.height > 0) {
                    ctx.drawImage(loaded, x, y, imgRect.width, imgRect.height);
                }
                ctx.globalAlpha = 1;
            }).catch(() => {
                // Fallback: try drawing the original img element directly (may taint canvas)
                try {
                    const imgRect = img.getBoundingClientRect();
                    const x = imgRect.left - mapRect.left;
                    const y = imgRect.top - mapRect.top;
                    if (imgRect.width > 0 && imgRect.height > 0) {
                        ctx.drawImage(img, x, y, imgRect.width, imgRect.height);
                    }
                } catch (_) { /* skip tile */ }
            })
        );
    });
    await Promise.all(tilePromises);

    // 2) Hide tiles and controls, then capture the whole map container natively with html2canvas
    const tilePane = mapEl.querySelector('.leaflet-tile-pane');
    const oldTileDisplay = tilePane ? tilePane.style.display : '';
    if (tilePane) tilePane.style.display = 'none';

    // IMPORTANT: mapEl usually has a CSS background-color. If html2canvas sees it, it draws an opaque 
    // box covering the tiles we just drew! We must force the mapEl background to be transparent during capture.
    const oldMapBg = mapEl.style.background;
    const oldMapBgColor = mapEl.style.backgroundColor;
    mapEl.style.background = 'transparent';
    mapEl.style.backgroundColor = 'transparent';

    // Also hide controls and sidebar which we already hid earlier, but just to be absolutely safe
    try {
        const overlayCanvas = await html2canvas(mapEl, {
            backgroundColor: null,
            scale: scale,
            logging: false,
            useCORS: true
        });
        
        if (overlayCanvas.width > 0 && overlayCanvas.height > 0) {
            ctx.scale(1 / scale, 1 / scale);
            ctx.drawImage(overlayCanvas, 0, 0);
            ctx.scale(scale, scale);
        }
    } catch (err) {
        console.error("Overlay capture failed:", err);
    } finally {
        if (tilePane) tilePane.style.display = oldTileDisplay;
        mapEl.style.background = oldMapBg;
        mapEl.style.backgroundColor = oldMapBgColor;
    }

    return canvas;
}

/** Load an image via fetch+blob to bypass CORS (works for Google tiles). */
function _loadImageCORS(url) {
    return fetch(url, { mode: 'cors' })
        .then(r => {
            if (!r.ok) throw new Error('fetch failed');
            return r.blob();
        })
        .then(blob => {
            const objUrl = URL.createObjectURL(blob);
            return _loadImage(objUrl).then(img => {
                URL.revokeObjectURL(objUrl);
                return img;
            });
        });
}

/** Simple image loader that returns a promise. */
function _loadImage(src) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.crossOrigin = 'anonymous';
        img.onload = () => resolve(img);
        img.onerror = reject;
        img.src = src;
    });
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
    temp.width = Math.round(w * sx);
    temp.height = Math.round(h * sy);
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

    try {
        const dataUrl = canvas.toDataURL('image/jpeg', 0.85);
        AppState.data[targetIdx][targetCol] = dataUrl;

        // ── AUTO-KOORDINATEN: Save current GPS position to the same row ──
        if (AppState.userMarker) {
            const gpsLatLng = AppState.userMarker.getLatLng();
            if (gpsLatLng && !isNaN(gpsLatLng.lat) && !isNaN(gpsLatLng.lng)) {
                const row = AppState.data[targetIdx];
                // Only fill if coordinates are not already set
                if (!row['Koordinaten'] || !row['GPS-Lat']) {
                    const dmsVal = convertToDMS(gpsLatLng.lat, true) + ' ' + convertToDMS(gpsLatLng.lng, false);
                    row['Koordinaten'] = dmsVal;
                    row['GPS-Lat'] = gpsLatLng.lat.toFixed(5);
                    row['GPS-Lng'] = gpsLatLng.lng.toFixed(5);
                    AppState.hiddenColumns.delete('special');
                    showToast(`📍 Zeile ${targetIdx + 1} — Koordinaten automatisch gespeichert`);
                }
            }
        }

        renderTable();
        saveToStorage();
        $('snipConfirmModal').style.display = 'none';

        // Upload to Firebase Storage so Prüfer can see the image
        _uploadImageToCloudIfPossible(targetIdx, targetCol, dataUrl);

    } catch (err) {
        if (err.name === 'SecurityError' || String(err).includes('tainted')) {
            showToast('Karte nicht speicherbar (Cross-Origin). Bitte Screenshot manuell erstellen.');
        } else {
            showToast('Fehler: ' + err.message);
        }
        console.error('saveFinalSnip error:', err);
    }
}

// Upload a single image to Firebase Storage and update the cloud document
async function _uploadImageToCloudIfPossible(rowIdx, colName, base64DataUrl) {
    if (typeof isFirebaseConfigured !== 'function' || !isFirebaseConfigured()) return;
    if (typeof uploadPhotoToStorage !== 'function') return;
    const pName = ($('projectName') || {}).value;
    if (!pName) return;
    try {
        const url = await uploadPhotoToStorage(pName, rowIdx, colName, base64DataUrl);
        if (url) {
            // Replace local base64 with Storage URL in the data
            AppState.data[rowIdx][colName] = url;
            saveToStorage();
        }
    } catch (e) {
        console.warn('Image upload to Storage failed:', e);
    }
}

function downloadSnip() {
    const canvas = $('snipEditorCanvas');
    if (!canvas) return;
    
    const targetSelect = $('snipTargetRow');
    const targetIdx = targetSelect ? parseInt(targetSelect.value) : -1;
    let filename = `Messstelle_${Date.now()}_1.png`;
    
    if (targetIdx > -1 && AppState.data[targetIdx]) {
        const rowData = AppState.data[targetIdx];
        const meterVal = (rowData['Meter [m]'] || '').toString().trim();
        if (meterVal !== '') {
            const cleanMeter = meterVal.replace(/[^a-zA-Z0-9_\-]/g, '_');
            filename = `${cleanMeter}_1.png`;
        }
    }

    const link = document.createElement('a');
    link.download = filename;
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

        if (g.class === 'basis') {
            const subCols = ['Kennzeichen', 'Alt-Kz.', 'Datum'];
            subCols.forEach(col => {
                const isColVisible = !AppState.hiddenColumns.has(col);
                const subItem = document.createElement('div');
                subItem.style.cssText = `display:flex;align-items:center;justify-content:space-between;padding:10px 16px 10px 32px;background:var(--bg-primary);border:1px solid ${isColVisible ? '#94a3b8' : 'var(--border)'};border-radius:var(--radius-sm);cursor:pointer;transition:all 0.15s;opacity:${isVisible ? 1 : 0.5};pointer-events:${isVisible ? 'auto' : 'none'};`;
                subItem.innerHTML = `
                    <span style="font-size:11px;font-weight:500;color:${isColVisible ? '#cbd5e1' : 'var(--text-muted)'}"><span style="margin-right: 8px; font-weight: bold; color: var(--text-muted);">↳</span>${col}</span>
                    <i class="fas ${isColVisible ? 'fa-eye' : 'fa-eye-slash'}" style="color:${isColVisible ? '#cbd5e1' : 'var(--text-muted)'}"></i>
                `;
                subItem.onclick = (e) => {
                    e.stopPropagation();
                    if (AppState.hiddenColumns.has(col)) AppState.hiddenColumns.delete(col);
                    else AppState.hiddenColumns.add(col);
                    renderTable();
                    saveToStorage();
                    openColManager();
                };
                grid.appendChild(subItem);
            });
        }
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
                    // Also delete from cloud
                    if (typeof deleteFromCloud === 'function') deleteFromCloud(name);
                    // If this was the currently open project, clear the table
                    const currentId = localStorage.getItem('current_project_id');
                    if (currentId === name) {
                        localStorage.removeItem('current_project_id');
                        AppState.data = [];
                        AppState.newCols = new Set();
                        const nameInput = $('projectName');
                        if (nameInput) nameInput.value = '';
                        renderTable();
                        // Clear map stars too
                        if (layers.star08) layers.star08.clearLayers();
                        if (layers.star16) layers.star16.clearLayers();
                        if (layers.star32) layers.star32.clearLayers();
                        if (layers.drawItems) layers.drawItems.clearLayers();
                        AppState.depthMarkers = { '0.8': [], '1.6': [], '3.2': [] };
                    }
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

        const headerFont = { name: 'Calibri', bold: true, size: 14 };
        const titleFont = { name: 'Calibri', bold: true, size: 18 };
        const dataFont = { name: 'Calibri', size: 12 };
        const center = { horizontal: 'center', vertical: 'middle', wrapText: true };
        const border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

        const depthsToExport = [];
        if (opts.export08) depthsToExport.push('08');
        if (opts.export16) depthsToExport.push('16');
        if (opts.export32) depthsToExport.push('32');
        const depthLabels = { '08': '0.80 m', '16': '1.60 m', '32': '3.20 m' };
        const depthArgbColors = { '08': 'FFF5D0FE', '16': 'FFFFFBEB', '32': 'FFCFFAFE' };
        const customCols = Array.from(AppState.newCols);

        // Determine dynamic row identifier for Worksheet 2 and Charts
        let idHeader = 'Messstelle';
        let idKey = '';
        if (!AppState.hiddenColumns.has('Kennzeichen')) {
            idHeader = 'Kennzeichen';
            idKey = 'Kennzeichen';
        } else if (!AppState.hiddenColumns.has('Alt-Kz.')) {
            idHeader = 'Alt-Kz.';
            idKey = 'Alt-Kz.';
        } else if (!AppState.hiddenColumns.has('Meter [m]')) {
            idHeader = 'Meter [m]';
            idKey = 'Meter [m]';
        }

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

        const h1 = [];
        const h2 = [];
        const colDefinitions = [];

        TABLE_STRUCTURE.forEach(g => {
            const isDepth = g.class === '08' || g.class === '16' || g.class === '32';
            if (isDepth) {
                if (!depthsToExport.includes(g.class)) return;
                const visibleCols = g.columns.filter(col => !AppState.hiddenColumns.has(col));
                if (visibleCols.length === 0) return;
                visibleCols.forEach((col, cIdx) => {
                    h1.push(cIdx === 0 ? g.group : '');
                    let label = col.split('_')[0];
                    if (label === 'MW [Ωm]') label = 'Mittelwert [Ωm]';
                    if (label === 'SD [Ωm]') label = 'Std.Abw. [Ωm]';
                    h2.push(label);

                    const isAnhang = col.toLowerCase().includes('anhang');
                    const isBilder = col.toLowerCase().includes('bilder');
                    colDefinitions.push({
                        key: col,
                        isImage: isAnhang || isBilder,
                        isAnhang: isAnhang,
                        isBilder: isBilder,
                        groupClass: g.class
                    });
                });
            } else if (g.class === 'basis') {
                const visibleCols = g.columns.filter(col => !AppState.hiddenColumns.has(col) && col !== 'Sprachsteuerung');
                visibleCols.forEach(col => {
                    h1.push(col);
                    h2.push('');
                    colDefinitions.push({ key: col, isBaseMerged: true, groupClass: g.class });
                });
                customCols.filter(col => !AppState.hiddenColumns.has(col)).forEach(col => {
                    h1.push(col);
                    h2.push('');
                    colDefinitions.push({ key: col, isBaseMerged: true, groupClass: g.class });
                });
            } else if (g.class === 'anhang_global') {
                if (AppState.hiddenColumns.has(g.class)) return;
                const visibleCols = g.columns.filter(col => !AppState.hiddenColumns.has(col));
                visibleCols.forEach((col, cIdx) => {
                    const isAnhangGlobal = col === 'Anhang_Global';
                    h1.push(g.group);
                    h2.push('');
                    colDefinitions.push({
                        key: col,
                        isBaseMerged: true,
                        isImage: isAnhangGlobal,
                        isAnhangGlobal: isAnhangGlobal,
                        groupClass: g.class
                    });
                });
            } else {
                if (AppState.hiddenColumns.has(g.class)) return;
                const visibleCols = g.columns.filter(col => !AppState.hiddenColumns.has(col));
                visibleCols.forEach(col => {
                    h1.push(g.group);
                    h2.push('');
                    colDefinitions.push({ key: col, isBaseMerged: true, groupClass: g.class });
                });
            }
        });

        const totalCols = h1.length;

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

        ws.getRow(HEADER_ROW).values = h1;
        ws.getRow(HEADER_ROW + 1).values = h2;
        ws.getRow(HEADER_ROW + 1).height = 30;

        [ws.getRow(HEADER_ROW), ws.getRow(HEADER_ROW + 1)].forEach(row => {
            row.eachCell((cell, colNum) => {
                cell.font = headerFont;
                cell.alignment = center;
                cell.border = border;
                
                const colDef = colDefinitions[colNum - 1];
                const groupClass = colDef ? colDef.groupClass : 'basis';
                
                let fillColor = 'FFE2E8F0'; // Default Slate color
                if (groupClass === '08') fillColor = 'FFF5D0FE';
                else if (groupClass === '16') fillColor = 'FFFFFBEB';
                else if (groupClass === '32') fillColor = 'FFCFFAFE';
                else if (groupClass === 'potential') fillColor = 'FFFFF7ED';
                else if (groupClass === 'spannung') fillColor = 'FFFFF1F2';
                else if (groupClass === 'strom') fillColor = 'FFEFF6FF';
                else if (groupClass === 'widerstand') fillColor = 'FFF3E8FF';
                
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
            });
        });

        // Vertical merges for base columns and 1-column groups
        for (let colNum = 1; colNum <= colDefinitions.length; colNum++) {
            const colDef = colDefinitions[colNum - 1];
            if (colDef.isBaseMerged) {
                ws.mergeCells(HEADER_ROW, colNum, HEADER_ROW + 1, colNum);
            }
        }

        // Set default width for all image/anhang columns upfront
        colDefinitions.forEach((colDef, colIdx) => {
            if (colDef.isImage) {
                ws.getColumn(colIdx + 1).width = 65;
            }
        });

        // Horizontal merges for multi-column groups
        let groupStart = -1;
        for (let colNum = 1; colNum <= h1.length; colNum++) {
            const val = h1[colNum - 1];
            const colDef = colDefinitions[colNum - 1];
            if (colDef.isBaseMerged) {
                if (groupStart !== -1 && colNum - 1 > groupStart) {
                    ws.mergeCells(HEADER_ROW, groupStart, HEADER_ROW, colNum - 1);
                }
                groupStart = -1;
                continue;
            }
            if (val !== '') {
                if (groupStart !== -1 && colNum - 1 > groupStart) {
                    ws.mergeCells(HEADER_ROW, groupStart, HEADER_ROW, colNum - 1);
                }
                groupStart = colNum;
            }
        }
        if (groupStart !== -1 && h1.length > groupStart) {
            ws.mergeCells(HEADER_ROW, groupStart, HEADER_ROW, h1.length);
        }

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
            const vals = [];
            
            colDefinitions.forEach(colDef => {
                if (colDef.isImage) {
                    vals.push(''); // Images are drawn over the cell, so cell value is empty
                } else {
                    let val = d[colDef.key] || '';
                    if (colDef.key === 'Koordinaten' && val) {
                        const parsed = parseCoordinates(String(val));
                        if (parsed) {
                            val = convertToDMS(parsed.lat, true) + ' ' + convertToDMS(parsed.lng, false);
                        }
                    }
                    vals.push(val);
                }
            });
            
            exRow.values = vals;
            exRow.eachCell(cell => { 
                cell.font = dataFont; 
                cell.alignment = center; 
                cell.border = border; 
            });

            // Insert images and set column width / row height
            let rowHasImage = false;
            for (let colIdx = 0; colIdx < colDefinitions.length; colIdx++) {
                const colDef = colDefinitions[colIdx];
                if (colDef.isImage) {
                    const imgData = d[colDef.key];
                    if (imgData && imgData.startsWith('data:image')) {
                        rowHasImage = true;
                        let ext = 'png';
                        if (imgData.includes('image/jpeg') || imgData.includes('image/jpg')) ext = 'jpeg';
                        
                        const img = new Image();
                        await new Promise((resolve) => {
                            img.onload = () => resolve();
                            img.onerror = () => resolve();
                            img.src = imgData;
                        });

                        let w = img.width || 480;
                        let h = img.height || 360;
                        const ratio = w / h;
                        const maxW = 480;
                        const maxH = 380;

                        if (w > maxW || h > maxH) {
                            if (w / maxW > h / maxH) {
                                w = maxW;
                                h = Math.floor(w / ratio);
                            } else {
                                h = maxH;
                                w = Math.floor(h * ratio);
                            }
                        }

                        const imgId = workbook.addImage({ base64: imgData.split(',')[1] || imgData, extension: ext });
                        
                        const colWidthPx = 520; // 65 width is ~520px
                        const rowHeightPx = 400; // 300 height in pt is ~400px
                        const dx = Math.max(0, Math.floor((colWidthPx - w) / 2));
                        const dy = Math.max(0, Math.floor((rowHeightPx - h) / 2));

                        ws.addImage(imgId, {
                            tl: { col: colIdx, row: DATA_START - 1 + i, colOff: dx * 9525, rowOff: dy * 9525 },
                            ext: { width: w, height: h }
                        });
                        ws.getColumn(colIdx + 1).width = 65; // Make image column wider for high-res
                    }
                }
            }
            
            if (rowHasImage) {
                exRow.height = 300; // Larger row height for high-res images
            } else {
                exRow.height = 24; // Compact standard row height
            }
        }

        // Auto column widths - fit to content, skip Anhang/Bilder columns
        const anhangColIndices = new Set();
        colDefinitions.forEach((colDef, colIdx) => {
            if (colDef.isImage) anhangColIndices.add(colIdx);
        });

        for (let colIdx = 0; colIdx < totalCols; colIdx++) {
            if (anhangColIndices.has(colIdx)) continue;
            const col = ws.getColumn(colIdx + 1);
            let maxLen = 10;
            const hVal = String(h1[colIdx] || h2[colIdx] || '');
            maxLen = Math.max(maxLen, hVal.length * 1.35 + 4);
            for (let r = 0; r < exportData.length; r++) {
                const cell = ws.getRow(DATA_START + r).getCell(colIdx + 1);
                const val = cell.value ? String(cell.value) : '';
                maxLen = Math.max(maxLen, val.length * 1.25 + 3);
            }
            col.width = Math.max(12, Math.min(Math.ceil(maxLen), 40));
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
        const sH1 = [idHeader];
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
                cell.font = headerFont; cell.alignment = center; cell.border = border;
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
            });
            row.height = 24;
        });

        statSheet.mergeCells(SH_ROW, 1, SH_ROW + 2, 1);
        let sMC = 2;
        depthsToExport.forEach(() => {
            statSheet.mergeCells(SH_ROW, sMC, SH_ROW, sMC + 2);
            // Merge Mittelwert and Standard Abweichung headers vertically
            statSheet.mergeCells(SH_ROW + 1, sMC + 1, SH_ROW + 2, sMC + 1);
            statSheet.mergeCells(SH_ROW + 1, sMC + 2, SH_ROW + 2, sMC + 2);
            sMC += 3;
        });

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
                const rowLabelVal = r === 1 ? ((idKey ? dataRow[idKey] : '') || (i + 1).toString()) : '';
                const values = [rowLabelVal];
                depthsToExport.forEach(d => {
                    const sfx = d === '08' ? '0.8' : (d === '16' ? '1.6' : '3.2');
                    values.push(dataRow['ρ' + r + ' [Ωm]_' + sfx] || '');
                    values.push(r === 1 ? (dataRow['MW [Ωm]_' + sfx] || '') : '');
                    values.push(r === 1 ? (dataRow['SD [Ωm]_' + sfx] || '') : '');
                });
                exRow.values = values;
                exRow.eachCell(cell => { cell.font = dataFont; cell.alignment = center; cell.border = border; });
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
        const statWidths = [{ width: 25 }];
        depthsToExport.forEach(() => { statWidths.push({ width: 18 }, { width: 22 }, { width: 30 }); });
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
            const trendColors = { '08': '#c026d3', '16': '#e28a16', '32': '#0284c7' };
            const depthColors = {
                '08': { solid: '#c026d3', light: '#f5d0fe' },
                '16': { solid: '#e28a16', light: '#fef3c7' },
                '32': { solid: '#0284c7', light: '#e0f2fe' }
            };

            // --- BALKENDIAGRAMM (Bar Chart with Error Bars) ---
            const barCanvas = document.createElement('canvas');
            barCanvas.width = 1600;
            barCanvas.height = 800;
            const bCtx = barCanvas.getContext('2d');

            // Background
            bCtx.fillStyle = '#ffffff';
            bCtx.fillRect(0, 0, 1600, 800);

            const bPad = { top: 90, bottom: 130, left: 110, right: 110 };
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

            const numIntervals = 6;
            const stepVal = bMax / numIntervals;

            const groupW = bW / filledData.length;
            const barWidth = (groupW * 0.65) / depthsToExport.length;

            // Draw solid body and cylinder borders
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

                    const colors = depthColors[d] || { solid: '#94a3b8', light: '#e2e8f0' };
                    const ecx = bx + (barWidth - 6) / 2;
                    const rx = (barWidth - 6) / 2;
                    const ry = 6;

                    // 1. Draw bottom ellipse
                    bCtx.fillStyle = colors.solid;
                    bCtx.beginPath();
                    bCtx.ellipse(ecx, bPad.top + bH, rx, ry, 0, 0, Math.PI * 2);
                    bCtx.fill();
                    bCtx.strokeStyle = '#000000';
                    bCtx.lineWidth = 1.5;
                    bCtx.beginPath();
                    bCtx.ellipse(ecx, bPad.top + bH, rx, ry, 0, 0, Math.PI);
                    bCtx.stroke();

                    // 2. Draw body rectangle
                    bCtx.fillStyle = colors.solid;
                    bCtx.fillRect(bx, by, barWidth - 6, barH);

                    // 3. Draw top ellipse (light filled)
                    bCtx.fillStyle = colors.light;
                    bCtx.beginPath();
                    bCtx.ellipse(ecx, by, rx, ry, 0, 0, Math.PI * 2);
                    bCtx.fill();
                    bCtx.strokeStyle = '#000000';
                    bCtx.lineWidth = 1.5;
                    bCtx.beginPath();
                    bCtx.ellipse(ecx, by, rx, ry, 0, 0, Math.PI * 2);
                    bCtx.stroke();

                    // 4. Draw vertical sides
                    bCtx.strokeStyle = '#000000';
                    bCtx.lineWidth = 1.5;
                    bCtx.beginPath();
                    bCtx.moveTo(bx, by);
                    bCtx.lineTo(bx, bPad.top + bH);
                    bCtx.stroke();
                    bCtx.beginPath();
                    bCtx.moveTo(bx + (barWidth - 6), by);
                    bCtx.lineTo(bx + (barWidth - 6), bPad.top + bH);
                    bCtx.stroke();

                    // 5. Draw error bar
                    if (sd > 0) {
                        const sdH = (sd / bMax) * bH;
                        bCtx.strokeStyle = '#000000';
                        bCtx.lineWidth = 1.5;
                        bCtx.beginPath(); bCtx.moveTo(ecx, by - sdH); bCtx.lineTo(ecx, by + Math.min(sdH, barH)); bCtx.stroke();
                        bCtx.beginPath(); bCtx.moveTo(ecx - 6, by - sdH); bCtx.lineTo(ecx + 6, by - sdH); bCtx.stroke();
                        bCtx.beginPath(); bCtx.moveTo(ecx - 6, by + Math.min(sdH, barH)); bCtx.lineTo(ecx + 6, by + Math.min(sdH, barH)); bCtx.stroke();
                    }
                });
            });

            // Outer border box (Individual colored lines)
            bCtx.strokeStyle = '#000000'; bCtx.lineWidth = 1.5;
            bCtx.beginPath(); bCtx.moveTo(bPad.left, bPad.top); bCtx.lineTo(bPad.left + bW, bPad.top); bCtx.stroke(); // top
            bCtx.beginPath(); bCtx.moveTo(bPad.left, bPad.top + bH); bCtx.lineTo(bPad.left + bW, bPad.top + bH); bCtx.stroke(); // bottom

            bCtx.strokeStyle = '#c026d3'; bCtx.lineWidth = 1.5;
            bCtx.beginPath(); bCtx.moveTo(bPad.left, bPad.top); bCtx.lineTo(bPad.left, bPad.top + bH); bCtx.stroke(); // left (magenta)

            bCtx.strokeStyle = '#e28a16'; bCtx.lineWidth = 1.5;
            bCtx.beginPath(); bCtx.moveTo(bPad.left + bW, bPad.top); bCtx.lineTo(bPad.left + bW, bPad.top + bH); bCtx.stroke(); // right (orange)

            // Ticks (Left, Right, Bottom, Top)
            for (let i = 0; i <= numIntervals; i++) {
                const val = i * stepVal;
                const y = bPad.top + bH - (bH * val / bMax);

                // Left major tick
                bCtx.strokeStyle = '#c026d3'; bCtx.lineWidth = 1.5;
                bCtx.beginPath(); bCtx.moveTo(bPad.left, y); bCtx.lineTo(bPad.left + 8, y); bCtx.stroke();

                // Right major tick
                bCtx.strokeStyle = '#e28a16'; bCtx.lineWidth = 1.5;
                bCtx.beginPath(); bCtx.moveTo(bPad.left + bW, y); bCtx.lineTo(bPad.left + bW - 8, y); bCtx.stroke();

                // Minor ticks
                if (i < numIntervals) {
                    for (let j = 1; j <= 4; j++) {
                        const mVal = val + j * (stepVal / 5);
                        const my = bPad.top + bH - (mVal / bMax) * bH;

                        bCtx.strokeStyle = '#c026d3'; bCtx.lineWidth = 1.0;
                        bCtx.beginPath(); bCtx.moveTo(bPad.left, my); bCtx.lineTo(bPad.left + 4, my); bCtx.stroke();

                        bCtx.strokeStyle = '#e28a16'; bCtx.lineWidth = 1.0;
                        bCtx.beginPath(); bCtx.moveTo(bPad.left + bW, my); bCtx.lineTo(bPad.left + bW - 4, my); bCtx.stroke();
                    }
                }
            }

            // X axis ticks (Bottom & Top)
            filledData.forEach((row, i) => {
                const cx = bPad.left + i * groupW + groupW / 2;

                // Bottom major tick
                bCtx.strokeStyle = '#000000'; bCtx.lineWidth = 1.5;
                bCtx.beginPath(); bCtx.moveTo(cx, bPad.top + bH); bCtx.lineTo(cx, bPad.top + bH - 8); bCtx.stroke();

                // Top major tick
                bCtx.beginPath(); bCtx.moveTo(cx, bPad.top); bCtx.lineTo(cx, bPad.top + 8); bCtx.stroke();

                // Minor ticks at boundaries
                const bx = bPad.left + i * groupW;
                bCtx.lineWidth = 1.0;
                bCtx.beginPath(); bCtx.moveTo(bx, bPad.top + bH); bCtx.lineTo(bx, bPad.top + bH - 4); bCtx.stroke();
                bCtx.beginPath(); bCtx.moveTo(bx, bPad.top); bCtx.lineTo(bx, bPad.top + 4); bCtx.stroke();

                if (i === filledData.length - 1) {
                    const endX = bPad.left + bW;
                    bCtx.beginPath(); bCtx.moveTo(endX, bPad.top + bH); bCtx.lineTo(endX, bPad.top + bH - 4); bCtx.stroke();
                    bCtx.beginPath(); bCtx.moveTo(endX, bPad.top); bCtx.lineTo(endX, bPad.top + 4); bCtx.stroke();
                }
            });

            // Rotated X labels
            filledData.forEach((row, i) => {
                const cx = bPad.left + i * groupW + groupW / 2;
                const label = (idKey ? row[idKey] : '') || (i + 1).toString();

                bCtx.save();
                bCtx.translate(cx, bPad.top + bH + 15);
                bCtx.rotate(-Math.PI / 4);
                bCtx.fillStyle = '#000000';
                bCtx.font = 'bold 14px Inter, sans-serif';
                bCtx.textAlign = 'right';
                bCtx.textBaseline = 'middle';
                bCtx.fillText(label, 0, 0);
                bCtx.restore();
            });

            // Left Y tick values (Magenta)
            for (let i = 0; i <= numIntervals; i++) {
                const val = i * stepVal;
                const y = bPad.top + bH - (bH * val / bMax);
                bCtx.fillStyle = '#c026d3';
                bCtx.font = 'bold 14px Inter, sans-serif';
                bCtx.textAlign = 'right';
                bCtx.textBaseline = 'middle';
                bCtx.fillText(Math.round(val).toLocaleString('de-DE'), bPad.left - 15, y);
            }

            // Right Y tick values (Orange)
            for (let i = 0; i <= numIntervals; i++) {
                const val = i * stepVal;
                const y = bPad.top + bH - (bH * val / bMax);
                bCtx.fillStyle = '#e28a16';
                bCtx.font = 'bold 14px Inter, sans-serif';
                bCtx.textAlign = 'left';
                bCtx.textBaseline = 'middle';
                bCtx.fillText(Math.round(val).toLocaleString('de-DE'), bPad.left + bW + 15, y);
            }

            // Y Axis Titles
            // Left Y axis title (Magenta)
            bCtx.save();
            bCtx.translate(35, bPad.top + bH / 2);
            bCtx.rotate(-Math.PI / 2);
            bCtx.fillStyle = '#c026d3'; bCtx.font = 'bold 14px Inter, sans-serif'; bCtx.textAlign = 'center';
            bCtx.fillText('ρ [Ω · m]', 0, 0);
            bCtx.restore();

            // Right Y axis title (Orange)
            bCtx.save();
            bCtx.translate(1600 - 35, bPad.top + bH / 2);
            bCtx.rotate(Math.PI / 2);
            bCtx.fillStyle = '#e28a16'; bCtx.font = 'bold 14px Inter, sans-serif'; bCtx.textAlign = 'center';
            bCtx.fillText('ρ [Ω · m]', 0, 0);
            bCtx.restore();

            // X Axis Title
            bCtx.fillStyle = '#374151'; bCtx.font = 'bold 14px Inter, sans-serif'; bCtx.textAlign = 'center';
            bCtx.fillText('Kennzeichen', bPad.left + bW / 2, 800 - 20);

            // Legend box (horizontal, centered above the plot area)
            bCtx.fillStyle = '#ffffff';
            bCtx.strokeStyle = '#000000';
            bCtx.lineWidth = 1.5;
            const legW = depthsToExport.length * 110 + 20;
            const legH = 35;
            const legX = bPad.left + (bW - legW) / 2;
            const legY = bPad.top - 50;
            bCtx.fillRect(legX, legY, legW, legH);
            bCtx.strokeRect(legX, legY, legW, legH);

            depthsToExport.forEach((d, di) => {
                const lx = legX + 15 + di * 110;
                const ly = legY + legH / 2;

                bCtx.fillStyle = depthColors[d].solid;
                bCtx.strokeStyle = '#000000';
                bCtx.lineWidth = 1.5;
                bCtx.fillRect(lx, ly - 7, 16, 14);
                bCtx.strokeRect(lx, ly - 7, 16, 14);

                bCtx.fillStyle = '#000000';
                bCtx.font = '14px Inter, sans-serif';
                bCtx.textAlign = 'left';
                bCtx.textBaseline = 'middle';
                bCtx.fillText(depthLabels[d], lx + 24, ly);
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

            // Connect dots with solid lines first
            depthsToExport.forEach((d, di) => {
                const sfx = d === '08' ? '0.8' : (d === '16' ? '1.6' : '3.2');
                const colors = depthColors[d] || { solid: '#94a3b8' };

                sCtx.strokeStyle = colors.solid; sCtx.lineWidth = 3;
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
            });

            // Draw markers + STD error bars
            depthsToExport.forEach((d, di) => {
                const sfx = d === '08' ? '0.8' : (d === '16' ? '1.6' : '3.2');
                const colors = depthColors[d] || { solid: '#94a3b8' };

                filledData.forEach((row, i) => {
                    const mw = parseFloat(row['MW [Ωm]_' + sfx]) || 0;
                    const sd = parseFloat(row['SD [Ωm]_' + sfx]) || 0;
                    if (mw <= 0) return;
                    const cx = bPad.left + i * groupW + groupW / 2;
                    const cy = bPad.top + bH - (mw / bMax) * bH;

                    // STD error bars (thin grey vertical line)
                    if (sd > 0) {
                        const sdPx = (sd / bMax) * bH;
                        sCtx.strokeStyle = '#4b5563'; sCtx.lineWidth = 1.5;
                        sCtx.beginPath(); sCtx.moveTo(cx, cy - sdPx); sCtx.lineTo(cx, cy + sdPx); sCtx.stroke();
                        sCtx.beginPath(); sCtx.moveTo(cx - 6, cy - sdPx); sCtx.lineTo(cx + 6, cy - sdPx); sCtx.stroke();
                        sCtx.beginPath(); sCtx.moveTo(cx - 6, cy + sdPx); sCtx.lineTo(cx + 6, cy + sdPx); sCtx.stroke();
                    }

                    // Marker draw
                    sCtx.fillStyle = colors.solid;
                    sCtx.strokeStyle = '#000000';
                    sCtx.lineWidth = 1.5;
                    if (d === '08') {
                        sCtx.beginPath(); sCtx.arc(cx, cy, 6, 0, Math.PI * 2); sCtx.fill(); sCtx.stroke();
                    } else if (d === '16') {
                        sCtx.fillRect(cx - 6, cy - 6, 12, 12); sCtx.strokeRect(cx - 6, cy - 6, 12, 12);
                    } else {
                        sCtx.beginPath(); sCtx.moveTo(cx, cy - 7); sCtx.lineTo(cx + 7, cy + 6); sCtx.lineTo(cx - 7, cy + 6); sCtx.closePath(); sCtx.fill(); sCtx.stroke();
                    }
                });
            });

            // Outer border box (Solid black border box)
            sCtx.strokeStyle = '#000000'; sCtx.lineWidth = 1.5;
            sCtx.strokeRect(bPad.left, bPad.top, bW, bH);

            // Ticks (Left, Right, Bottom, Top)
            for (let i = 0; i <= numIntervals; i++) {
                const val = i * stepVal;
                const y = bPad.top + bH - (bH * val / bMax);

                // Left major tick
                sCtx.strokeStyle = '#000000'; sCtx.lineWidth = 1.5;
                sCtx.beginPath(); sCtx.moveTo(bPad.left, y); sCtx.lineTo(bPad.left + 8, y); sCtx.stroke();

                // Right major tick
                sCtx.beginPath(); sCtx.moveTo(bPad.left + bW, y); sCtx.lineTo(bPad.left + bW - 8, y); sCtx.stroke();

                // Minor ticks
                if (i < numIntervals) {
                    for (let j = 1; j <= 4; j++) {
                        const mVal = val + j * (stepVal / 5);
                        const my = bPad.top + bH - (mVal / bMax) * bH;

                        sCtx.lineWidth = 1.0;
                        sCtx.beginPath(); sCtx.moveTo(bPad.left, my); sCtx.lineTo(bPad.left + 4, my); sCtx.stroke();
                        sCtx.beginPath(); sCtx.moveTo(bPad.left + bW, my); sCtx.lineTo(bPad.left + bW - 4, my); sCtx.stroke();
                    }
                }
            }

            // X ticks (Bottom & Top)
            filledData.forEach((row, i) => {
                const cx = bPad.left + i * groupW + groupW / 2;

                // Bottom major tick
                sCtx.strokeStyle = '#000000'; sCtx.lineWidth = 1.5;
                sCtx.beginPath(); sCtx.moveTo(cx, bPad.top + bH); sCtx.lineTo(cx, bPad.top + bH - 8); sCtx.stroke();

                // Top major tick
                sCtx.beginPath(); sCtx.moveTo(cx, bPad.top); sCtx.lineTo(cx, bPad.top + 8); sCtx.stroke();

                // Minor ticks at boundaries
                const bx = bPad.left + i * groupW;
                sCtx.lineWidth = 1.0;
                sCtx.beginPath(); sCtx.moveTo(bx, bPad.top + bH); sCtx.lineTo(bx, bPad.top + bH - 4); sCtx.stroke();
                sCtx.beginPath(); sCtx.moveTo(bx, bPad.top); sCtx.lineTo(bx, bPad.top + 4); sCtx.stroke();

                if (i === filledData.length - 1) {
                    const endX = bPad.left + bW;
                    sCtx.beginPath(); sCtx.moveTo(endX, bPad.top + bH); sCtx.lineTo(endX, bPad.top + bH - 4); sCtx.stroke();
                    sCtx.beginPath(); sCtx.moveTo(endX, bPad.top); sCtx.lineTo(endX, bPad.top + 4); sCtx.stroke();
                }
            });

            // Rotated X labels
            filledData.forEach((row, i) => {
                const cx = bPad.left + i * groupW + groupW / 2;
                const label = (idKey ? row[idKey] : '') || (i + 1).toString();

                sCtx.save();
                sCtx.translate(cx, bPad.top + bH + 15);
                sCtx.rotate(-Math.PI / 4);
                sCtx.fillStyle = '#000000';
                sCtx.font = 'bold 14px Inter, sans-serif';
                sCtx.textAlign = 'right';
                sCtx.textBaseline = 'middle';
                sCtx.fillText(label, 0, 0);
                sCtx.restore();
            });

            // Left Y tick values (Black)
            for (let i = 0; i <= numIntervals; i++) {
                const val = i * stepVal;
                const y = bPad.top + bH - (bH * val / bMax);
                sCtx.fillStyle = '#000000';
                sCtx.font = 'bold 14px Inter, sans-serif';
                sCtx.textAlign = 'right';
                sCtx.textBaseline = 'middle';
                sCtx.fillText(Math.round(val).toLocaleString('de-DE'), bPad.left - 15, y);
            }

            // Left Y Axis Title (Black)
            sCtx.save();
            sCtx.translate(35, bPad.top + bH / 2);
            sCtx.rotate(-Math.PI / 2);
            sCtx.fillStyle = '#000000'; sCtx.font = 'bold 14px Inter, sans-serif'; sCtx.textAlign = 'center';
            sCtx.fillText('ρ [Ω · m]', 0, 0);
            sCtx.restore();

            // X Axis Title
            sCtx.fillStyle = '#374151'; sCtx.font = 'bold 14px Inter, sans-serif'; sCtx.textAlign = 'center';
            sCtx.fillText('Kennzeichen', bPad.left + bW / 2, 800 - 20);

            // Legend box (vertical, in top right of plot area)
            sCtx.fillStyle = '#ffffff';
            sCtx.strokeStyle = '#000000';
            sCtx.lineWidth = 1.5;
            const legW2 = 130;
            const legH2 = depthsToExport.length * 30 + 10;
            const legX2 = bPad.left + bW - legW2 - 20;
            const legY2 = bPad.top + 20;
            sCtx.fillRect(legX2, legY2, legW2, legH2);
            sCtx.strokeRect(legX2, legY2, legW2, legH2);

            depthsToExport.forEach((d, di) => {
                const ly = legY2 + 20 + di * 30;
                const tx = legX2 + 45;
                const mx = legX2 + 22;

                // Line behind marker
                sCtx.strokeStyle = depthColors[d].solid;
                sCtx.lineWidth = 2.0;
                sCtx.beginPath(); sCtx.moveTo(mx - 15, ly); sCtx.lineTo(mx + 15, ly); sCtx.stroke();

                // Marker
                sCtx.fillStyle = depthColors[d].solid;
                sCtx.strokeStyle = '#000000';
                sCtx.lineWidth = 1.5;
                if (d === '08') {
                    sCtx.beginPath(); sCtx.arc(mx, ly, 6, 0, Math.PI * 2); sCtx.fill(); sCtx.stroke();
                } else if (d === '16') {
                    sCtx.fillRect(mx - 6, ly - 6, 12, 12); sCtx.strokeRect(mx - 6, ly - 6, 12, 12);
                } else {
                    sCtx.beginPath(); sCtx.moveTo(mx, ly - 7); sCtx.lineTo(mx + 7, ly + 6); sCtx.lineTo(mx - 7, ly + 6); sCtx.closePath(); sCtx.fill(); sCtx.stroke();
                }

                sCtx.fillStyle = '#000000';
                sCtx.font = '14px Inter, sans-serif';
                sCtx.textAlign = 'left';
                sCtx.textBaseline = 'middle';
                sCtx.fillText(depthLabels[d], tx, ly);
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

        // Download individual photos from the exported rows simultaneously
        let downloadCount = 0;
        exportData.forEach((d, rowIdx) => {
            const imageCols = [];
            if (opts.export08) imageCols.push({ key: 'Bilder_0.8', suffix: '0.8' });
            if (opts.export16) imageCols.push({ key: 'Bilder_1.6', suffix: '1.6' });
            if (opts.export32) imageCols.push({ key: 'Bilder_3.2', suffix: '3.2' });
            imageCols.push({ key: 'Anhang_Global', suffix: 'Gesamt' });

            imageCols.forEach((colInfo) => {
                const imgData = d[colInfo.key];
                if (imgData && imgData.startsWith('data:image')) {
                    let ext = 'jpg';
                    if (imgData.includes('image/png')) ext = 'png';
                    else if (imgData.includes('image/webp')) ext = 'webp';

                    let filename = `Foto_${rowIdx + 1}_${colInfo.suffix}.${ext}`;
                    const meterVal = (d['Meter [m]'] || '').toString().trim();
                    if (meterVal !== '') {
                        const cleanMeter = meterVal.replace(/[^a-zA-Z0-9_\-]/g, '_');
                        filename = `${cleanMeter}_${colInfo.suffix}.${ext}`;
                    }

                    downloadCount++;
                    // Trigger download with a slight staggered delay to prevent browser download congestion
                    setTimeout(() => {
                        const link = document.createElement('a');
                        link.download = filename;
                        link.href = imgData;
                        document.body.appendChild(link);
                        link.click();
                        document.body.removeChild(link);
                    }, downloadCount * 120);
                }
            });
        });
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
            '<a:t>Bodenwiderstand \u03C1 [\u03A9 \u00B7 m] - ' + projectName + '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>' +
            '<c:autoTitleDeleted val="0"/><c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:rAngAx val="1"/><c:perspective val="30"/></c:view3D><c:plotArea><c:layout/>' +
            '<c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:shape val="cylinder"/>' +
            seriesXml +
            '<c:axId val="111"/><c:axId val="222"/><c:axId val="333"/></c:bar3DChart>' +
            '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/>' +
            '<c:title><c:tx><c:rich><a:bodyPr rot="0" vert="horz"/><a:lstStyle/><a:p><a:r><a:rPr lang="de-DE" sz="1400"/><a:t>Kennzeichen</a:t></a:r></a:p></c:rich></c:tx></c:title></c:catAx>' +
            '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111"/>' +
            '<c:title><c:tx><c:rich><a:bodyPr rot="-5400000" vert="horz"/><a:lstStyle/><a:p><a:r><a:rPr lang="de-DE" sz="1400"/><a:t>Bodenwiderstand \u03C1 [\u03A9 \u00B7 m]</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx>' +
            '<c:serAx><c:axId val="333"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:crossAx val="222"/></c:serAx>' +
            '</c:plotArea><c:legend><c:legendPos val="tr"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>';

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
    if (!canvas || canvas.clientWidth < 10) return;
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
            let label = '';
            if (!AppState.hiddenColumns.has('Kennzeichen')) {
                label = row['Kennzeichen'];
            } else if (!AppState.hiddenColumns.has('Alt-Kz.')) {
                label = row['Alt-Kz.'];
            }
            ctx.fillText(label || (i + 1), pad.left + i * bG + bG / 2, pad.top + chartH + 14);
        }
    });

    // Axis labels
    ctx.fillStyle = '#f1f5f9'; ctx.font = 'bold 14px Inter'; ctx.textAlign = 'center';
    let axisLabel = 'Messstelle';
    if (!AppState.hiddenColumns.has('Kennzeichen')) {
        axisLabel = 'Kennzeichen';
    } else if (!AppState.hiddenColumns.has('Alt-Kz.')) {
        axisLabel = 'Alt-Kz.';
    }
    ctx.fillText(axisLabel, pad.left + chartW / 2, h - 10);
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

// --- IMAGE EDITOR LOGIC ---
let currentBilderRow = -1;
let currentBilderCol = '';
let imgEditorActiveTool = 'pencil';
let isImgDrawing = false;
let savedImgData = null;
let lastImgPos = {x:0, y:0};
let startImgPos = {x:0, y:0};

let imgAnnotations = [];
let selectedAnnotation = null;
let draggingAnnotation = null;
let draggingHandle = null;
let dragOffset = {};
let bgImage = null;
let currentPath = null;

function triggerCamera(rowIdx, colName) {
    currentBilderRow = rowIdx;
    currentBilderCol = colName;
    const input = $('cameraInput');
    if (input) {
        input.value = '';
        input.click();
    }
}

function openImageEditor(dataUrl, rowIdx, colName) {
    currentBilderRow = rowIdx;
    currentBilderCol = colName;
    initImageEditorCanvas(dataUrl);
}

function initImageEditorCanvas(imgSrc) {
    const canvas = $('imgEditorCanvas');
    const ctx = canvas.getContext('2d');
    const img = new Image();
    img.onload = () => {
        let w = img.width;
        let h = img.height;
        const maxDim = 4096;
        if (w > maxDim || h > maxDim) {
            if (w > h) { h = Math.floor((h/w)*maxDim); w = maxDim; }
            else { w = Math.floor((w/h)*maxDim); h = maxDim; }
        }
        canvas.width = w;
        canvas.height = h;
        bgImage = img;
        imgAnnotations = [];
        selectedAnnotation = null;
        draggingAnnotation = null;
        draggingHandle = null;
        redrawCanvas();
        $('imageEditorModal').style.display = 'flex';
    };
    img.src = imgSrc;
}

function distToSegment(p, a, b) {
    const l2 = Math.pow(a.x - b.x, 2) + Math.pow(a.y - b.y, 2);
    if (l2 === 0) return Math.sqrt(Math.pow(p.x - a.x, 2) + Math.pow(p.y - a.y, 2));
    let t = ((p.x - a.x) * (b.x - a.x) + (p.y - a.y) * (b.y - a.y)) / l2;
    t = Math.max(0, Math.min(1, t));
    return Math.sqrt(Math.pow(p.x - (a.x + t * (b.x - a.x)), 2) + Math.pow(p.y - (a.y + t * (b.y - a.y)), 2));
}

function redrawCanvas() {
    const canvas = $('imgEditorCanvas');
    if (!canvas || !bgImage) return;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(bgImage, 0, 0, canvas.width, canvas.height);

    const scale = canvas.width / 1200;

    imgAnnotations.forEach(ann => {
        ctx.strokeStyle = ann.color;
        ctx.fillStyle = ann.color;
        ctx.lineWidth = (ann.width || 3) * scale;
        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';

        if (ann.type === 'path') {
            if (ann.isEraser) {
                ctx.globalCompositeOperation = 'destination-out';
            } else if (ann.isHighlighter) {
                ctx.globalCompositeOperation = 'source-over';
                const r = parseInt(ann.color.slice(1,3), 16);
                const g = parseInt(ann.color.slice(3,5), 16);
                const b = parseInt(ann.color.slice(5,7), 16);
                ctx.strokeStyle = `rgba(${r},${g},${b},0.3)`;
            } else {
                ctx.globalCompositeOperation = 'source-over';
            }
            ctx.beginPath();
            ann.points.forEach((pt, idx) => {
                if (idx === 0) ctx.moveTo(pt.x, pt.y);
                else ctx.lineTo(pt.x, pt.y);
            });
            ctx.stroke();
            ctx.globalCompositeOperation = 'source-over';
        } else if (ann.type === 'line') {
            ctx.beginPath();
            ctx.moveTo(ann.x1, ann.y1);
            ctx.lineTo(ann.x2, ann.y2);
            ctx.stroke();
        } else if (ann.type === 'circle') {
            ctx.beginPath();
            ctx.arc(ann.cx, ann.cy, ann.r, 0, 2 * Math.PI);
            ctx.stroke();
        } else if (ann.type === 'text') {
            const fontSize = Math.round(20 * scale);
            ctx.font = `bold ${fontSize}px Inter, sans-serif`;
            ctx.textBaseline = 'middle';
            ctx.textAlign = 'center';
            
            const txtWidth = ctx.measureText(ann.text).width;
            const padX = 6 * scale;
            const boxH = 28 * scale;
            const boxYOffset = 14 * scale;
            
            ctx.fillStyle = 'rgba(15, 23, 42, 0.8)';
            ctx.fillRect(ann.x - txtWidth/2 - padX, ann.y - boxYOffset, txtWidth + padX * 2, boxH);
            ctx.strokeStyle = ann.color;
            ctx.lineWidth = 1.5 * scale;
            ctx.strokeRect(ann.x - txtWidth/2 - padX, ann.y - boxYOffset, txtWidth + padX * 2, boxH);
            
            ctx.fillStyle = '#ffffff';
            ctx.fillText(ann.text, ann.x, ann.y);
        }

        // Draw handles for selection in move mode
        if (ann === selectedAnnotation && imgEditorActiveTool === 'move') {
            ctx.strokeStyle = '#3b82f6';
            ctx.lineWidth = 1.5 * scale;
            ctx.setLineDash([4 * scale, 4 * scale]);
            if (ann.type === 'line') {
                ctx.beginPath(); ctx.moveTo(ann.x1, ann.y1); ctx.lineTo(ann.x2, ann.y2); ctx.stroke();
                drawHandle(ann.x1, ann.y1, scale);
                drawHandle(ann.x2, ann.y2, scale);
            } else if (ann.type === 'circle') {
                ctx.beginPath(); ctx.arc(ann.cx, ann.cy, ann.r, 0, 2 * Math.PI); ctx.stroke();
                drawHandle(ann.cx, ann.cy, scale);
                drawHandle(ann.cx + ann.r, ann.cy, scale);
            } else if (ann.type === 'text') {
                const txtWidth = ctx.measureText(ann.text).width;
                ctx.strokeRect(ann.x - txtWidth/2 - 8 * scale, ann.y - 16 * scale, txtWidth + 16 * scale, 32 * scale);
                drawHandle(ann.x, ann.y, scale);
            }
            ctx.setLineDash([]);
        }
    });
}

function drawHandle(x, y, scale) {
    const canvas = $('imgEditorCanvas');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    ctx.save();
    ctx.globalCompositeOperation = 'source-over';
    ctx.fillStyle = '#3b82f6';
    ctx.strokeStyle = '#ffffff';
    ctx.lineWidth = 2 * scale;
    ctx.beginPath();
    ctx.arc(x, y, 6 * scale, 0, 2 * Math.PI);
    ctx.fill();
    ctx.stroke();
    ctx.restore();
}

document.addEventListener('DOMContentLoaded', () => {
    const camInput = $('cameraInput');
    if (camInput) {
        camInput.addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = function(event) {
                initImageEditorCanvas(event.target.result);
            };
            reader.readAsDataURL(file);
        });
    }

    // Tools Setup
    const tools = ['toolPencil', 'toolHighlighter', 'toolLine', 'toolCircle', 'toolText', 'toolMove', 'toolEraser'];
    tools.forEach(id => {
        const btn = $(id);
        if(btn) {
            btn.addEventListener('click', () => {
                tools.forEach(t => {
                    const el = $(t);
                    if (el) el.classList.remove('active');
                });
                btn.classList.add('active');
                imgEditorActiveTool = id.replace('tool', '').toLowerCase();
                selectedAnnotation = null;
                redrawCanvas();
            });
        }
    });

    on('toolDelete', 'click', () => {
        if (selectedAnnotation) {
            imgAnnotations = imgAnnotations.filter(ann => ann !== selectedAnnotation);
            selectedAnnotation = null;
            redrawCanvas();
            showToast('Element gelöscht');
        } else {
            showToast('Zuerst Verschieben-Werkzeug wählen und Element anklicken');
        }
    });

    // Update selected annotation color when color picker changes
    const colorInput = $('imgEditorColor');
    if (colorInput) {
        colorInput.addEventListener('input', (e) => {
            if (selectedAnnotation) {
                selectedAnnotation.color = e.target.value;
                redrawCanvas();
            }
        });
    }

    const canvas = $('imgEditorCanvas');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');

    function getPos(e) {
        const rect = canvas.getBoundingClientRect();
        let clientX = e.clientX;
        let clientY = e.clientY;
        if (e.touches && e.touches.length > 0) {
            clientX = e.touches[0].clientX;
            clientY = e.touches[0].clientY;
        } else if (e.changedTouches && e.changedTouches.length > 0) {
            clientX = e.changedTouches[0].clientX;
            clientY = e.changedTouches[0].clientY;
        }
        return {
            x: (clientX - rect.left) * (canvas.width / rect.width),
            y: (clientY - rect.top) * (canvas.height / rect.height)
        };
    }

    let lastTap = 0;
    function handleEditText(pos, e) {
        const scale = canvas.width / 1200;
        for (let i = imgAnnotations.length - 1; i >= 0; i--) {
            const ann = imgAnnotations[i];
            if (ann.type === 'text') {
                ctx.save();
                ctx.font = `bold ${Math.round(20 * scale)}px Inter, sans-serif`;
                const txtWidth = ctx.measureText(ann.text).width;
                ctx.restore();
                if (Math.abs(pos.x - ann.x) < (txtWidth / 2 + 15 * scale) && Math.abs(pos.y - ann.y) < 24 * scale) {
                    if (e) e.preventDefault();
                    setTimeout(() => {
                        const newTxt = prompt('Text ändern:', ann.text);
                        if (newTxt !== null && newTxt.trim() !== '') {
                            ann.text = newTxt.trim();
                            redrawCanvas();
                        }
                    }, 50);
                    return true;
                }
            }
        }
        return false;
    }

    function startDraw(e) {
        const scale = canvas.width / 1200;
        if (e.touches) {
            const currentTime = new Date().getTime();
            const tapLength = currentTime - lastTap;
            lastTap = currentTime;
            if (tapLength < 300 && tapLength > 0) {
                if (imgEditorActiveTool === 'move') {
                    const pos = getPos(e);
                    if (handleEditText(pos, e)) return;
                }
            }
        }

        e.preventDefault();
        isImgDrawing = true;
        startImgPos = getPos(e);
        lastImgPos = startImgPos;
        const color = $('imgEditorColor').value;

        if (imgEditorActiveTool === 'pencil' || imgEditorActiveTool === 'highlighter' || imgEditorActiveTool === 'eraser') {
            currentPath = {
                type: 'path',
                points: [startImgPos],
                color: color,
                width: imgEditorActiveTool === 'pencil' ? 3 : (imgEditorActiveTool === 'highlighter' ? 15 : 20),
                isHighlighter: imgEditorActiveTool === 'highlighter',
                isEraser: imgEditorActiveTool === 'eraser'
            };
        } else if (imgEditorActiveTool === 'text') {
            isImgDrawing = false;
            const txt = prompt('Text eingeben (z.B. Nummer oder Messwert):');
            if (txt && txt.trim()) {
                imgAnnotations.push({
                    type: 'text',
                    text: txt.trim(),
                    x: startImgPos.x,
                    y: startImgPos.y,
                    color: color
                });
                redrawCanvas();
            }
        } else if (imgEditorActiveTool === 'move') {
            draggingHandle = null;
            draggingAnnotation = null;

            if (selectedAnnotation) {
                const sa = selectedAnnotation;
                if (sa.type === 'line') {
                    if (Math.hypot(startImgPos.x - sa.x1, startImgPos.y - sa.y1) < 20 * scale) draggingHandle = 'x1y1';
                    else if (Math.hypot(startImgPos.x - sa.x2, startImgPos.y - sa.y2) < 20 * scale) draggingHandle = 'x2y2';
                } else if (sa.type === 'circle') {
                    if (Math.hypot(startImgPos.x - sa.cx, startImgPos.y - sa.cy) < 20 * scale) draggingHandle = 'center';
                    else if (Math.hypot(startImgPos.x - (sa.cx + sa.r), startImgPos.y - sa.cy) < 20 * scale) draggingHandle = 'radius';
                } else if (sa.type === 'text') {
                    if (Math.hypot(startImgPos.x - sa.x, startImgPos.y - sa.y) < 20 * scale) draggingHandle = 'center';
                }
            }

            if (!draggingHandle) {
                selectedAnnotation = null;
                for (let i = imgAnnotations.length - 1; i >= 0; i--) {
                    const ann = imgAnnotations[i];
                    if (ann.type === 'text') {
                        ctx.save();
                        ctx.font = `bold ${Math.round(20 * scale)}px Inter, sans-serif`;
                        const txtWidth = ctx.measureText(ann.text).width;
                        ctx.restore();
                        if (Math.abs(startImgPos.x - ann.x) < (txtWidth / 2 + 15 * scale) && Math.abs(startImgPos.y - ann.y) < 24 * scale) {
                            selectedAnnotation = ann;
                            draggingAnnotation = ann;
                            dragOffset = { x: startImgPos.x - ann.x, y: startImgPos.y - ann.y };
                            break;
                        }
                    } else if (ann.type === 'circle') {
                        const dist = Math.hypot(startImgPos.x - ann.cx, startImgPos.y - ann.cy);
                        if (Math.abs(dist - ann.r) < 25 * scale || dist < 25 * scale) {
                            selectedAnnotation = ann;
                            draggingAnnotation = ann;
                            dragOffset = { x: startImgPos.x - ann.cx, y: startImgPos.y - ann.cy };
                            break;
                        }
                    } else if (ann.type === 'line') {
                        const dist = distToSegment(startImgPos, { x: ann.x1, y: ann.y1 }, { x: ann.x2, y: ann.y2 });
                        if (dist < 25 * scale) {
                            selectedAnnotation = ann;
                            draggingAnnotation = ann;
                            dragOffset = {
                                x1: startImgPos.x - ann.x1, y1: startImgPos.y - ann.y1,
                                x2: startImgPos.x - ann.x2, y2: startImgPos.y - ann.y2
                            };
                            break;
                        }
                    }
                }
            }
            redrawCanvas();
        }
    }

    function moveDraw(e) {
        if (!isImgDrawing && !draggingAnnotation && !draggingHandle) return;
        e.preventDefault();
        const pos = getPos(e);
        const color = $('imgEditorColor').value;
        const scale = canvas.width / 1200;

        if (imgEditorActiveTool === 'pencil' || imgEditorActiveTool === 'highlighter' || imgEditorActiveTool === 'eraser') {
            if (currentPath) {
                currentPath.points.push(pos);
                ctx.lineCap = 'round';
                ctx.lineJoin = 'round';
                ctx.lineWidth = currentPath.width * scale;
                if (currentPath.isEraser) {
                    ctx.globalCompositeOperation = 'destination-out';
                    ctx.strokeStyle = 'rgba(0,0,0,1)';
                } else if (currentPath.isHighlighter) {
                    ctx.globalCompositeOperation = 'source-over';
                    const r = parseInt(color.slice(1,3), 16);
                    const g = parseInt(color.slice(3,5), 16);
                    const b = parseInt(color.slice(5,7), 16);
                    ctx.strokeStyle = `rgba(${r},${g},${b},0.3)`;
                } else {
                    ctx.globalCompositeOperation = 'source-over';
                    ctx.strokeStyle = color;
                }
                ctx.beginPath();
                ctx.moveTo(lastImgPos.x, lastImgPos.y);
                ctx.lineTo(pos.x, pos.y);
                ctx.stroke();
                ctx.globalCompositeOperation = 'source-over';
                lastImgPos = pos;
            }
        } else if (imgEditorActiveTool === 'line' || imgEditorActiveTool === 'circle') {
            redrawCanvas();
            ctx.globalCompositeOperation = 'source-over';
            ctx.lineWidth = 3 * scale;
            ctx.strokeStyle = color;
            ctx.beginPath();
            if (imgEditorActiveTool === 'line') {
                ctx.moveTo(startImgPos.x, startImgPos.y);
                ctx.lineTo(pos.x, pos.y);
            } else if (imgEditorActiveTool === 'circle') {
                const radius = Math.sqrt(Math.pow(pos.x - startImgPos.x, 2) + Math.pow(pos.y - startImgPos.y, 2));
                ctx.arc(startImgPos.x, startImgPos.y, radius, 0, 2 * Math.PI);
            }
            ctx.stroke();
        } else if (imgEditorActiveTool === 'move') {
            const sa = selectedAnnotation;
            if (draggingHandle && sa) {
                if (sa.type === 'line') {
                    if (draggingHandle === 'x1y1') { sa.x1 = pos.x; sa.y1 = pos.y; }
                    else if (draggingHandle === 'x2y2') { sa.x2 = pos.x; sa.y2 = pos.y; }
                } else if (sa.type === 'circle') {
                    if (draggingHandle === 'center') { sa.cx = pos.x; sa.cy = pos.y; }
                    else if (draggingHandle === 'radius') {
                        sa.r = Math.hypot(pos.x - sa.cx, pos.y - sa.cy);
                    }
                } else if (sa.type === 'text') {
                    if (draggingHandle === 'center') { sa.x = pos.x; sa.y = pos.y; }
                }
                redrawCanvas();
            } else if (draggingAnnotation) {
                const da = draggingAnnotation;
                if (da.type === 'text') {
                    da.x = pos.x - dragOffset.x;
                    da.y = pos.y - dragOffset.y;
                } else if (da.type === 'circle') {
                    da.cx = pos.x - dragOffset.x;
                    da.cy = pos.y - dragOffset.y;
                } else if (da.type === 'line') {
                    da.x1 = pos.x - dragOffset.x1;
                    da.y1 = pos.y - dragOffset.y1;
                    da.x2 = pos.x - dragOffset.x2;
                    da.y2 = pos.y - dragOffset.y2;
                }
                redrawCanvas();
            }
        }
    }

    function stopDraw(e) {
        if (!isImgDrawing && !draggingAnnotation && !draggingHandle) return;
        e.preventDefault();
        isImgDrawing = false;
        draggingAnnotation = null;
        draggingHandle = null;
        
        const pos = getPos(e);
        const color = $('imgEditorColor').value;

        if (imgEditorActiveTool === 'pencil' || imgEditorActiveTool === 'highlighter' || imgEditorActiveTool === 'eraser') {
            if (currentPath) {
                imgAnnotations.push(currentPath);
                currentPath = null;
            }
        } else if (imgEditorActiveTool === 'line') {
            imgAnnotations.push({
                type: 'line',
                x1: startImgPos.x,
                y1: startImgPos.y,
                x2: pos.x,
                y2: pos.y,
                color: color
            });
            redrawCanvas();
        } else if (imgEditorActiveTool === 'circle') {
            const radius = Math.sqrt(Math.pow(pos.x - startImgPos.x, 2) + Math.pow(pos.y - startImgPos.y, 2));
            imgAnnotations.push({
                type: 'circle',
                cx: startImgPos.x,
                cy: startImgPos.y,
                r: radius,
                color: color
            });
            redrawCanvas();
        }
    }

    canvas.addEventListener('mousedown', startDraw);
    canvas.addEventListener('mousemove', moveDraw);
    window.addEventListener('mouseup', stopDraw);
    
    canvas.addEventListener('touchstart', startDraw, {passive: false});
    canvas.addEventListener('touchmove', moveDraw, {passive: false});
    window.addEventListener('touchend', stopDraw);

    canvas.addEventListener('dblclick', (e) => {
        if (imgEditorActiveTool !== 'move') return;
        const pos = getPos(e);
        handleEditText(pos, e);
    });

    const btnSave = $('btnSaveImageEdit');
    if (btnSave) {
        btnSave.addEventListener('click', () => {
            selectedAnnotation = null;
            redrawCanvas();
            
            const tempCanvas = document.createElement('canvas');
            tempCanvas.width = canvas.width;
            tempCanvas.height = canvas.height;
            const tCtx = tempCanvas.getContext('2d');
            tCtx.fillStyle = '#ffffff';
            tCtx.fillRect(0, 0, canvas.width, canvas.height);
            tCtx.drawImage(canvas, 0, 0);
            
            const dataUrl = tempCanvas.toDataURL('image/jpeg', 0.98);
            if (currentBilderRow > -1 && currentBilderCol) {
                AppState.data[currentBilderRow][currentBilderCol] = dataUrl;
                saveToStorage();
                renderTable();
                // Upload to Firebase Storage so Prüfer can see the image
                _uploadImageToCloudIfPossible(currentBilderRow, currentBilderCol, dataUrl);
            }
            $('imageEditorModal').style.display = 'none';
        });
    }

    const btnDownload = $('btnDownloadImageEdit');
    if (btnDownload) {
        btnDownload.addEventListener('click', () => {
            const prevSelection = selectedAnnotation;
            selectedAnnotation = null;
            redrawCanvas();
            
            const tempCanvas = document.createElement('canvas');
            tempCanvas.width = canvas.width;
            tempCanvas.height = canvas.height;
            const tCtx = tempCanvas.getContext('2d');
            tCtx.fillStyle = '#ffffff';
            tCtx.fillRect(0, 0, canvas.width, canvas.height);
            tCtx.drawImage(canvas, 0, 0);
            
            let filename = `Foto_${Date.now()}_1.jpg`;
            if (currentBilderRow > -1 && AppState.data[currentBilderRow]) {
                const rowData = AppState.data[currentBilderRow];
                const meterVal = (rowData['Meter [m]'] || '').toString().trim();
                if (meterVal !== '') {
                    const cleanMeter = meterVal.replace(/[^a-zA-Z0-9_\-]/g, '_');
                    filename = `${cleanMeter}_1.jpg`;
                }
            }
            
            const dataUrl = tempCanvas.toDataURL('image/jpeg', 0.98);
            const link = document.createElement('a');
            link.download = filename;
            link.href = dataUrl;
            link.click();
            
            selectedAnnotation = prevSelection;
            redrawCanvas();
            showToast('Foto heruntergeladen');
        });
    }
});

// ============================================
// CLOUD REVIEW / PRÜFUNG FUNCTIONS
// ============================================

let _allCloudProjects = [];
let _currentReviewFilter = 'submitted';
let _currentReviewProject = null;

// ─── SUBMIT FOR REVIEW (Messhelfer) ───
async function handleSubmitForReview() {
    const pName = ($('projectName') || {}).value;
    if (!pName || pName.trim() === '') {
        showToast('Bitte zuerst einen Projektnamen eingeben');
        return;
    }

    if (AppState.data.length === 0) {
        showToast('Keine Messdaten vorhanden');
        return;
    }

    if (typeof isFirebaseConfigured === 'function' && isFirebaseConfigured()) {
        const confirmed = confirm(`Projekt "${pName}" zur Prüfung einreichen?\n\nDer Prüfer wird benachrichtigt und kann die Messdaten überprüfen.`);
        if (!confirmed) return;

        // Save latest data to cloud first, then mark as submitted
        saveToStorage();
        showToast('Wird hochgeladen...');

        // Wait for cloud save to complete before submitting
        const project = {
            name: pName,
            data: AppState.data,
            mapData: layers && layers.drawItems ? layers.drawItems.toGeoJSON() : null,
            hiddenMapColors: Array.from(AppState.hiddenMapColors),
            newCols: Array.from(AppState.newCols)
        };

        try {
            await saveToCloud(pName, project);
        } catch(e) {
            console.warn('Pre-submit cloud save failed:', e);
        }

        const success = await submitForReview(pName);
        if (success) {
            showReviewStatusBar('submitted', '', '');
        }
    } else {
        showToast('Cloud nicht konfiguriert — Firebase-Einrichtung erforderlich');
    }
}

// ─── REVIEW STATUS BAR (shown at bottom for Messhelfer) ───
function showReviewStatusBar(status, comment, reviewedBy) {
    const bar = $('reviewStatusBar');
    if (!bar) return;

    bar.style.display = 'block';
    bar.className = 'review-status-bar status-' + status;

    const icon = $('reviewStatusIcon');
    const text = $('reviewStatusText');

    if (status === 'submitted') {
        if (icon) icon.innerHTML = '<i class="fas fa-clock"></i>';
        if (text) text.textContent = 'Zur Prüfung eingereicht — warte auf Freigabe...';
    } else if (status === 'approved') {
        if (icon) icon.innerHTML = '<i class="fas fa-check-circle"></i>';
        if (text) text.textContent = `Genehmigt${reviewedBy ? ' von ' + reviewedBy : ''}! ✅`;
    } else if (status === 'rejected') {
        if (icon) icon.innerHTML = '<i class="fas fa-times-circle"></i>';
        if (text) text.textContent = `Abgelehnt${reviewedBy ? ' von ' + reviewedBy : ''}: ${comment || 'Kein Grund angegeben'}`;
    }
}

// ─── OPEN REVIEW PANEL (Prüfer) ───
async function openReviewPanel() {
    const modal = $('reviewModal');
    if (!modal) return;

    if (typeof isFirebaseConfigured === 'function' && isFirebaseConfigured()) {
        // Cloud mode — fetch from Firestore
        try {
            _allCloudProjects = await getAllCloudProjects();
        } catch (err) {
            console.error('Failed to fetch cloud projects:', err);
            _allCloudProjects = [];
        }
    } else {
        // Local mode — use localStorage
        const library = JSON.parse(localStorage.getItem('messstellen_library') || '{}');
        _allCloudProjects = Object.values(library).map(p => ({
            id: p.name,
            name: p.name,
            data: p.data || [],
            rowCount: (p.data || []).length,
            reviewStatus: p.reviewStatus || 'draft',
            updatedAt: p.timestamp ? { toDate: () => new Date(p.timestamp) } : null,
            updatedBy: 'Lokal',
            reviewComment: '',
            reviewedBy: ''
        }));
    }

    renderReviewListFiltered(_currentReviewFilter);
    modal.style.display = 'flex';
}

// ─── RENDER REVIEW LIST (called by real-time listener or manually) ───
function renderReviewList(projects) {
    _allCloudProjects = projects || [];
    renderReviewListFiltered(_currentReviewFilter);

    // Update notification dot
    const pending = _allCloudProjects.filter(p => p.reviewStatus === 'submitted');
    const btnReview = $('btnReviewPanel');
    if (btnReview) {
        const existingDot = btnReview.querySelector('.notification-dot');
        if (pending.length > 0) {
            if (!existingDot) {
                const dot = document.createElement('span');
                dot.className = 'notification-dot';
                btnReview.appendChild(dot);
            }
        } else {
            if (existingDot) existingDot.remove();
        }
    }
}

function renderReviewListFiltered(filter) {
    _currentReviewFilter = filter;
    const list = $('reviewProjectList');
    if (!list) return;

    let filtered = _allCloudProjects;
    if (filter !== 'all') {
        filtered = _allCloudProjects.filter(p => p.reviewStatus === filter);
    }

    if (filtered.length === 0) {
        const emptyMessages = {
            submitted: 'Keine Projekte zur Prüfung eingereicht',
            approved: 'Keine genehmigten Projekte',
            rejected: 'Keine abgelehnten Projekte',
            all: 'Keine Projekte in der Cloud'
        };
        list.innerHTML = `<p style="color:var(--text-muted);text-align:center;padding:32px;font-size:13px;">
            <i class="fas fa-inbox" style="font-size:24px;display:block;margin-bottom:8px;opacity:0.3;"></i>
            ${emptyMessages[filter] || 'Keine Projekte'}
        </p>`;
        return;
    }

    list.innerHTML = '';
    filtered.forEach(project => {
        const card = document.createElement('div');
        card.className = 'review-project-card';

        const statusClass = 'status-' + (project.reviewStatus || 'draft');
        const statusLabels = {
            draft: '📝 Entwurf',
            submitted: '⏳ Zur Prüfung',
            approved: '✅ Genehmigt',
            rejected: '❌ Abgelehnt'
        };
        const statusLabel = statusLabels[project.reviewStatus] || 'Unbekannt';

        let updatedStr = '';
        if (project.updatedAt) {
            try {
                const d = project.updatedAt.toDate ? project.updatedAt.toDate() : new Date(project.updatedAt);
                updatedStr = d.toLocaleString('de-DE');
            } catch (e) {
                updatedStr = '';
            }
        }

        card.innerHTML = `
            <div class="review-project-info">
                <div class="review-project-name">${project.name || project.id}</div>
                <div class="review-project-meta">
                    <span>${project.updatedBy || ''}</span>
                    ${updatedStr ? `<span>·</span><span>${updatedStr}</span>` : ''}
                </div>
                ${project.reviewComment ? `<div style="font-size:11px;color:var(--danger);margin-top:4px;"><i class="fas fa-comment"></i> ${project.reviewComment}</div>` : ''}
            </div>
            <span class="review-status-chip ${statusClass}">${statusLabel}</span>
        `;

        card.onclick = () => openReviewDetail(project);
        list.appendChild(card);
    });
}

// ─── REVIEW DETAIL VIEW ───
function openReviewDetail(project) {
    _currentReviewProject = project;
    const modal = $('reviewDetailModal');
    const title = $('reviewDetailTitle');
    const content = $('reviewDetailContent');
    const actionsBar = $('reviewActionsBar');

    if (!modal || !content) return;

    if (title) title.textContent = project.name || 'Projekt prüfen';

    // Build data table
    const data = project.data || [];
    if (data.length === 0) {
        content.innerHTML = '<p style="color:var(--text-muted);text-align:center;padding:24px;">Keine Messdaten vorhanden</p>';
    } else {
        // Determine which columns to show (non-empty ones)
        const allCols = new Set();
        data.forEach(row => {
            if (row && typeof row === 'object') {
                Object.keys(row).forEach(k => {
                    if (k !== '_isNew' && k !== '_originalIndex' && row[k] !== '' && row[k] !== null && row[k] !== undefined) {
                        // Skip large base64 images
                        if (typeof row[k] === 'string' && (row[k].startsWith('data:image') || row[k] === '__IMAGE_REF__')) return;
                        allCols.add(k);
                    }
                });
            }
        });

        const cols = Array.from(allCols);
        const priorityCols = ['Kennzeichen', 'Typ', 'Örtlichkeit', 'Meter [m]', 'Koordinaten'];
        const sortedCols = [
            ...priorityCols.filter(c => cols.includes(c)),
            ...cols.filter(c => !priorityCols.includes(c))
        ];

        let html = '<div style="overflow-x:auto;"><table class="review-detail-table"><thead><tr>';
        html += '<th>#</th>';
        sortedCols.forEach(c => {
            const label = c.includes('_') ? c.split('_')[0] : c;
            html += `<th>${label}</th>`;
        });
        html += '</tr></thead><tbody>';

        data.forEach((row, idx) => {
            if (!row || typeof row !== 'object') return;
            html += '<tr>';
            html += `<td style="text-align:center;font-weight:700;color:var(--text-muted);">${idx + 1}</td>`;
            sortedCols.forEach(c => {
                let val = row[c] || '';
                if (typeof val === 'string' && val.length > 50) val = val.substring(0, 50) + '...';
                html += `<td>${val}</td>`;
            });
            html += '</tr>';
        });

        html += '</tbody></table></div>';
        content.innerHTML = html;
    }

    // Show/hide approve/reject buttons based on status
    if (actionsBar) {
        if (project.reviewStatus === 'submitted') {
            actionsBar.style.display = 'block';
        } else {
            actionsBar.style.display = 'none';
        }
    }

    // Wire up approve/reject buttons
    const btnApprove = $('btnApproveProject');
    const btnReject = $('btnRejectProject');

    if (btnApprove) {
        btnApprove.onclick = async () => {
            const projectName = project.name || project.id;
            if (typeof isFirebaseConfigured === 'function' && isFirebaseConfigured()) {
                await approveProject(projectName);
            } else {
                // Local mode
                let library = JSON.parse(localStorage.getItem('messstellen_library') || '{}');
                if (library[projectName]) {
                    library[projectName].reviewStatus = 'approved';
                    localStorage.setItem('messstellen_library', JSON.stringify(library));
                    showToast('✅ Projekt genehmigt!');
                }
            }
            modal.style.display = 'none';
            openReviewPanel(); // Refresh the list
        };
    }

    if (btnReject) {
        btnReject.onclick = async () => {
            const reason = ($('reviewRejectReason') || {}).value || '';
            if (!reason.trim()) {
                showToast('Bitte einen Ablehnungsgrund eingeben');
                $('reviewRejectReason').focus();
                return;
            }
            const projectName = project.name || project.id;
            if (typeof isFirebaseConfigured === 'function' && isFirebaseConfigured()) {
                await rejectProject(projectName, reason);
            } else {
                let library = JSON.parse(localStorage.getItem('messstellen_library') || '{}');
                if (library[projectName]) {
                    library[projectName].reviewStatus = 'rejected';
                    library[projectName].reviewComment = reason;
                    localStorage.setItem('messstellen_library', JSON.stringify(library));
                    showToast('❌ Projekt abgelehnt');
                }
            }
            modal.style.display = 'none';
            openReviewPanel();
        };
    }

    modal.style.display = 'flex';
}


// ─── VOICE DICTATION (Spracheingabe) — per row, sequential ───
let _voiceRecognition = null;
let _voiceActive = false;
let _voiceRowIdx = -1;
let _voiceColIdx = 0;
let _voiceMicBtn = null;

// The sequence of columns to fill, in order: 0.8m (R1,R2,R3) then 1.6m (R1,R2,R3)
const VOICE_SEQUENCE = [
    'R1 [Ω]_0.8', 'R2 [Ω]_0.8', 'R3 [Ω]_0.8',
    'R1 [Ω]_1.6', 'R2 [Ω]_1.6', 'R3 [Ω]_1.6'
];

function startRowVoiceDictation(rowIdx, micBtn) {
    if (!('webkitSpeechRecognition' in window) && !('SpeechRecognition' in window)) {
        showToast('Spracheingabe nur in Chrome/Edge verfügbar');
        return;
    }

    // If already listening on this row, stop
    if (_voiceActive && _voiceRowIdx === rowIdx) {
        stopRowVoiceDictation();
        return;
    }

    // If listening on another row, stop that first
    if (_voiceActive) stopRowVoiceDictation();

    if (!AppState.data[rowIdx]) {
        showToast('Zeile existiert nicht');
        return;
    }

    _voiceRowIdx = rowIdx;
    _voiceMicBtn = micBtn;

    // Determine which column to fill based on selected cell, default R1/0.8
    let targetCol = 'R1 [Ω]_0.8';
    if (AppState.selectedCell) {
        const selCol = AppState.selectedCell.split('-').slice(1).join('-');
        if (selCol && selCol.includes('[Ω]')) {
            targetCol = selCol;
        }
    }
    const seqIdx = VOICE_SEQUENCE.indexOf(targetCol);
    if (seqIdx >= 0) {
        _voiceColIdx = seqIdx;
    } else {
        _voiceColIdx = 0;
        targetCol = VOICE_SEQUENCE[0];
    }
    window._voiceTargetCol = targetCol;

    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    _voiceRecognition = new SpeechRecognition();
    _voiceRecognition.lang = 'de-DE';
    _voiceRecognition.continuous = true;
    _voiceRecognition.interimResults = false;
    _voiceRecognition.maxAlternatives = 1;

    _voiceActive = true;
    if (micBtn) {
        micBtn.classList.add('listening');
        micBtn.style.color = '#ef4444';
        micBtn.querySelector('i').className = 'fas fa-stop';
    }
    showVoiceHUD();

    _voiceRecognition.onresult = function(event) {
        const lastResult = event.results[event.results.length - 1];
        if (!lastResult.isFinal) return;
        const transcript = lastResult[0].transcript.trim().toLowerCase();

        // Stop commands
        if (/\b(stopp|stop|fertig|ende)\b/.test(transcript)) {
            stopRowVoiceDictation();
            return;
        }

        const value = parseVoiceNumber(transcript);
        if (value === null) {
            updateVoiceHUDMessage('Nicht erkannt: "' + transcript + '"');
            return;
        }

        // Fill selected column
        const targetCol = window._voiceTargetCol || 'R1 [Ω]_0.8';
        pushUndo();
        AppState.data[_voiceRowIdx][targetCol] = value;
        updateLiveCalculations(AppState.data[_voiceRowIdx], targetCol, null);
        renderTable();
        debouncedSave(1000);

        const label = targetCol.replace(' [Ω]', '').replace('_', '/').replace('.', ',');
        showToast('✓ Zeile ' + (_voiceRowIdx + 1) + ' ' + label + ' = ' + value);

        // Advance to next column in sequence
        if (_voiceColIdx < VOICE_SEQUENCE.length - 1) {
            _voiceColIdx++;
            const nextCol = VOICE_SEQUENCE[_voiceColIdx];
            window._voiceTargetCol = nextCol;
            updateVoiceHUD(targetCol, value);
        } else {
            updateVoiceHUDMessage('✓ ' + label + ' = ' + value);
            setTimeout(stopRowVoiceDictation, 1200);
        }
    };

    _voiceRecognition.onerror = function(event) {
        if (event.error === 'no-speech') return;
        if (event.error === 'not-allowed') {
            showToast('Mikrofon-Zugriff verweigert');
            stopRowVoiceDictation();
        }
    };

    _voiceRecognition.onend = function() {
        if (_voiceActive) {
            try { _voiceRecognition.start(); } catch(e) {}
        }
    };

    try {
        _voiceRecognition.start();
        showToast('🎤 Zeile ' + (rowIdx + 1) + ' — Werte nacheinander sprechen');
    } catch (e) {
        showToast('Mikrofon-Fehler');
        _voiceActive = false;
    }
}

// Parse a spoken number from German speech → returns string or null
function parseVoiceNumber(transcript) {
    let cleaned = transcript
        .replace(/komma/g, '.').replace(/punkt/g, '.')
        .replace(/,/g, '.')
        .replace(/\bist\b|\bgleich\b|\bergibt\b/g, '')
        .replace(/\br\s*[123]\b/gi, '')  // remove "R1" "R2" "R3" references
        .replace(/\beins\b/g, '1').replace(/\bzwei\b/g, '2').replace(/\bdrei\b/g, '3')
        .replace(/\bvier\b/g, '4').replace(/\bfünf\b/g, '5').replace(/\bsechs\b/g, '6')
        .replace(/\bsieben\b/g, '7').replace(/\bacht\b/g, '8').replace(/\bneun\b/g, '9')
        .replace(/\bnull\b/g, '0').replace(/\bzehn\b/g, '10')
        .replace(/\belf\b/g, '11').replace(/\bzwölf\b/g, '12')
        .replace(/\bdreizehn\b/g, '13').replace(/\bvierzehn\b/g, '14').replace(/\bfünfzehn\b/g, '15')
        .replace(/\bzwanzig\b/g, '20').replace(/\bdreißig\b/g, '30');

    const match = cleaned.match(/\d+\.?\d*/);
    return match ? match[0] : null;
}

function stopRowVoiceDictation() {
    _voiceActive = false;
    if (_voiceRecognition) {
        try { _voiceRecognition.stop(); } catch(e) {}
        _voiceRecognition = null;
    }
    if (_voiceMicBtn) {
        _voiceMicBtn.classList.remove('listening');
        _voiceMicBtn.style.color = '';
        const icon = _voiceMicBtn.querySelector('i');
        if (icon) icon.className = 'fas fa-microphone';
        _voiceMicBtn = null;
    }
    // Also reset any other mic buttons
    document.querySelectorAll('.row-mic-btn').forEach(b => {
        b.classList.remove('listening');
        b.style.color = '';
        const i = b.querySelector('i');
        if (i) i.className = 'fas fa-microphone';
    });
    removeVoiceHUD();
    _voiceRowIdx = -1;
    _voiceColIdx = 0;
}

function showVoiceHUD() {
    removeVoiceHUD();
    const targetCol = window._voiceTargetCol || 'R1 [Ω]_0.8';
    const label = targetCol.replace(' [Ω]', '').replace('_', '/').replace('.', ',');
    const hud = document.createElement('div');
    hud.id = 'voiceHUD';
    hud.style.cssText = 'position:fixed;bottom:60px;left:50%;transform:translateX(-50%);background:rgba(15,23,42,0.97);border:2px solid #ef4444;border-radius:12px;padding:14px 22px;z-index:999999;display:flex;align-items:center;gap:14px;font-size:13px;color:#f1f5f9;box-shadow:0 8px 32px rgba(0,0,0,0.5);';
    hud.innerHTML = `<i class="fas fa-microphone-alt" style="color:#ef4444;font-size:20px;"></i>
        <div>
            <div id="voiceHUDStatus" style="font-weight:600;">Zeile ${_voiceRowIdx + 1}: <span style="color:#fbbf24;">${label}</span></div>
            <div id="voiceHUDLast" style="font-size:11px;color:#64748b;margin-top:3px;">Einfach Zahl sprechen</div>
        </div>
        <button onclick="stopRowVoiceDictation()" style="background:#ef4444;color:white;border:none;border-radius:6px;padding:8px 14px;font-size:12px;font-weight:600;cursor:pointer;">Stopp</button>`;
    document.body.appendChild(hud);
}

function removeVoiceHUD() {
    const hud = document.getElementById('voiceHUD');
    if (hud) hud.remove();
}

function updateVoiceHUD(col, value) {
    const last = document.getElementById('voiceHUDLast');
    if (last) last.textContent = `✓ ${col.replace(' [Ω]', '').replace('_', ' / ')} = ${value}`;
    const status = document.getElementById('voiceHUDStatus');
    const nextCol = window._voiceTargetCol;
    if (status && nextCol) {
        const label = nextCol.replace(' [Ω]', '').replace('_', ' / ');
        status.innerHTML = `Zeile ${_voiceRowIdx + 1} — sage Wert für: <span style="color:#fbbf24;">${label}</span>`;
    }
}

function updateVoiceHUDMessage(msg) {
    const el = document.getElementById('voiceHUDLast');
    if (el) el.textContent = msg;
}

// Old toolbar button → start dictation on first empty/selected row
function openVoiceDictation() {
    let rowIdx = -1;
    if (AppState.selectedCell) {
        rowIdx = parseInt(AppState.selectedCell.split('-')[0]);
    } else {
        rowIdx = AppState.data.findIndex(row => !row['R1 [Ω]_0.8'] && !row['R2 [Ω]_0.8'] && !row['R3 [Ω]_0.8']);
        if (rowIdx < 0) rowIdx = 0;
    }
    if (rowIdx < 0) { showToast('Bitte zuerst eine Zeile auswählen'); return; }
    const micBtn = document.querySelector(`.row-mic-btn[data-row="${rowIdx}"]`);
    startRowVoiceDictation(rowIdx, micBtn);
}

window.openVoiceDictation = openVoiceDictation;
window.startRowVoiceDictation = startRowVoiceDictation;
window.stopRowVoiceDictation = stopRowVoiceDictation;

