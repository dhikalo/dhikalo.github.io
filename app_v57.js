/* ===================================
   MESSSTELLEN MANAGER PRO - v5.4 (STABLE)
   Full Audit Features + High-End Mapping
   =================================== */
console.log("🚀 Messstellen Manager Pro v5.4.200 (Updated) Loaded!");

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

const COLUMN_UNITS = {
    'R1': 'Ω', 'R2': 'Ω', 'R3': 'Ω', 'ρ1': 'Ωm', 'ρ2': 'Ωm', 'ρ3': 'Ωm',
    'MW': 'Ωm', 'SD': 'Ωm',
    'Potential Ein': 'V', 'Potential Aus': 'V', 'Potential AC': 'V',
    'Spannung Ein': 'V', 'Spannung Aus': 'V',
    'Strom Ein': 'A', 'Strom Aus': 'A', 'Strom Diff.': 'A', 'Mikro Diff.': 'µA',
    'R': 'Ω', 'Bodenwid. [Ωm]': 'Ωm', 'RaAusbreitung': 'Ω'
};

const DEPTH_COLORS = {
    'basis': { bg: 'rgba(230, 237, 243, 0.1)', border: '#8b949e', text: '#e6edf3' },
    '08': { bg: 'rgba(255, 0, 255, 0.1)', border: '#ff00ff', text: '#ff00ff' },
    '16': { bg: 'rgba(255, 255, 0, 0.1)', border: '#ffff00', text: '#ffff00' },
    '32': { bg: 'rgba(0, 255, 255, 0.1)', border: '#00ffff', text: '#00ffff' },
    'zusatz': { bg: 'rgba(57, 255, 20, 0.1)', border: '#39ff14', text: '#39ff14' },
    'potential': { bg: 'rgba(255, 140, 0, 0.1)', border: '#ff8c00', text: '#ff8c00' },
    'spannung': { bg: 'rgba(255, 0, 60, 0.1)', border: '#ff003c', text: '#ff003c' },
    'strom': { bg: 'rgba(0, 242, 255, 0.1)', border: '#00f2ff', text: '#00f2ff' },
    'widerstand': { bg: 'rgba(188, 19, 254, 0.1)', border: '#bc13fe', text: '#bc13fe' },
    'audit': { bg: 'rgba(230, 237, 243, 0.05)', border: '#8b949e', text: '#8b949e' },
    'special': { bg: 'rgba(0, 242, 255, 0.05)', border: '#00f2ff', text: '#00f2ff' },
    'anhang_global': { bg: 'rgba(255, 255, 255, 0.1)', border: '#ffffff', text: '#ffffff' }
};

const AppState = {
    data: [],
    hiddenColumns: new Set(['32', 'potential', 'spannung', 'strom', 'widerstand', 'audit', 'special']),
    zoomLevel: 100,
    selectedCell: null,
    columnWidths: {},
    activeColor: '#FF00FF',
    liveFollow: false,      // ✔ Manual by default: wait for user to click Standortermittlung
    newCols: new Set(),
    filters: {
        'Kennzeichen': ''
    },
    userMarker: null,
    drawItems: null,
    hiddenMapColors: new Set([]),
    watchId: null,
    firstLocationFound: false, // Critical for initial zoom
    undoStack: [],
    redoStack: [],
    chartType: 'balken',
    showStd: false
};

function pushState() {
    AppState.undoStack.push(JSON.stringify(AppState.data));
    if (AppState.undoStack.length > 50) AppState.undoStack.shift();
    AppState.redoStack = []; // Clear redo on new action
}

function undo() {
    if (AppState.undoStack.length === 0) {
        showToast("ℹ️ Kein weiterer Verlauf verfügbar");
        return;
    }
    AppState.redoStack.push(JSON.stringify(AppState.data));
    AppState.data = JSON.parse(AppState.undoStack.pop());
    
    // Auto-switch to table tab so user sees the change
    switchTab('table'); 
    renderTable();
    saveToStorage();
    showToast("⟲ Rückgängig gemacht");
}

function redo() {
    if (AppState.redoStack.length === 0) {
        showToast("ℹ️ Keine weiteren Schritte zum Wiederholen");
        return;
    }
    AppState.undoStack.push(JSON.stringify(AppState.data));
    AppState.data = JSON.parse(AppState.redoStack.pop());
    
    // Auto-switch to table tab so user sees the change
    switchTab('table');
    renderTable();
    saveToStorage();
    showToast("⟳ Wiederholt");
}

function switchTab(tabId) {
    document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('#btnTabTable, #btnTabPlot').forEach(b => b.classList.remove('active'));

    if (tabId === 'table') {
        document.getElementById('tabTable').classList.add('active');
        document.getElementById('btnTabTable').classList.add('active');
        renderTable();
    } else if (tabId === 'plot') {
        document.getElementById('tabPlot').classList.add('active');
        document.getElementById('btnTabPlot').classList.add('active');
        // Delay ensures canvas has correct width from DOM before drawing
        setTimeout(renderAppPlot, 100); 
    }
}

// GLOBAL ZOOM HANDLER (Full Layout Scale)
window.handleZoom = function(delta) {
    if(!AppState) return;
    AppState.zoomLevel = Math.max(50, Math.min(200, AppState.zoomLevel + delta));
    const val = AppState.zoomLevel;
    const display = document.getElementById('zoomLevelDisplay');
    if(display) display.innerText = val + '%';
    
    const wrapper = document.querySelector('.table-wrapper');
    if(wrapper) {
        // "Full Zoom" - scales everything (borders, boxes, text)
        wrapper.style.zoom = val / 100;
        wrapper.style.webkitZoom = val / 100;
        showToast(`Ansicht: ${val}%`, 500);
    }
};

const STORAGE_KEY = 'messstellen_v54_stable_peak';

let map, layers = {}, captureMode = false;
let snipTexts = [], draggingTextIdx = -1;
// currentSnipWidth entfernt zur Vermeidung von Konflikten

const safeOn = (id, evt, cb) => document.getElementById(id)?.addEventListener(evt, cb);

document.addEventListener('DOMContentLoaded', () => {
    // REGISTER SERVICE WORKER FOR OFFLINE
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('./sw.js')
            .then(() => console.log("PWA: Offline-Modus aktiv ✓"))
            .catch(err => console.log("PWA: Fehler", err));
    }

    // loadFromStorage(); // DISABLED: user wants empty start
    initUI();
    initConnectionCheck();
});

function initConnectionCheck() {
    const updateStats = () => showToast(window.navigator.onLine ? "System: ONLINE" : "System: OFFLINE (Eingeschränkte Karte)");
    window.addEventListener('online', updateStats);
    window.addEventListener('offline', updateStats);
}

function initUI() {
    safeOn('btnHome', 'click', () => {
        document.getElementById('welcomeOverlay').style.display = 'flex';
        document.getElementById('mainApp').style.display = 'none';
        document.getElementById('fileStatus').innerText = "Keine Datei ausgewählt";
        document.getElementById('btnFinalStart').disabled = true;
        showToast("🏠 Zurück zum Hauptmenü");
    });

    safeOn('btnWelcomeBrowse', 'click', () => document.getElementById('fileInput').click());
    
    safeOn('btnWelcomeBlank', 'click', () => {
        AppState.data = [];
        createEmptyRows(50);
        document.getElementById('fileStatus').innerText = "✓ NEUES PROJEKT ERSTELLT";
        document.getElementById('btnFinalStart').disabled = false;
        showToast("✓ Leeres Projekt bereit");
    });

    safeOn('btnFinalStart', 'click', () => {
        document.getElementById('welcomeOverlay').style.display = 'none';
        document.getElementById('mainApp').style.display = 'flex';
        renderTable();
        showToast("🚀 Projekt gestartet");
    });

    safeOn('btnBrowse', 'click', () => document.getElementById('fileInput').click());
    safeOn('fileInput', 'change', (e) => { 
        if(e.target.files[0]) {
            handleFileUpload(e.target.files[0]);
        } 
    });
    
    safeOn('btnStartBlank', 'click', () => { 
        document.getElementById('mapLayout').classList.remove('show');
        if (AppState.data.length === 0) createEmptyRows(50);
        renderTable(); 
        showToast("✓ App bereit");
    });

    safeOn('btnColManager', 'click', openColManager);
    safeOn('btnCloseColManager', 'click', () => {
        document.getElementById('colManagerModal').style.display = 'none';
    });

    safeOn('btnExportExcel', 'click', openExportModal);
    safeOn('btnOpenProjects', 'click', () => {
        renderProjectList();
        document.getElementById('standortModal').style.display = 'block';
    });
    safeOn('btnCloseStandortModal', 'click', () => {
        document.getElementById('standortModal').style.display = 'none';
    });
    safeOn('btnSaveAll', 'click', () => { 
        saveToStorage(); 
        showToast("✅ Alles sicher gespeichert (Tabelle & Karte)!", 3000); 
    });

    safeOn('btnDeleteRow', 'click', (e) => { 
        if(e) e.preventDefault();
        if(AppState.selectedCell) { 
            pushState();
            AppState.data.splice(parseInt(AppState.selectedCell.split('-')[0]), 1); 
            renderTable(); 
            saveToStorage();
            showToast("✓ Zeile gelöscht");
        } 
    });

    safeOn('btnAddRow', 'click', (e) => {
        if(e) e.preventDefault();
        pushState();
        if (AppState.selectedCell) {
            const rowIndex = parseInt(AppState.selectedCell.split('-')[0]);
            AppState.data.splice(rowIndex + 1, 0, { _isNew: true }); 
        } else {
            AppState.data.push({ _isNew: true });
        }
        renderTable();
        saveToStorage();
        showToast("✓ Zeile hinzugefügt");
    });

    safeOn('btnAddCol', 'click', (e) => {
        if(e) e.preventDefault();
        const cName = prompt("Name der neuen Spalte eingeben:");
        if (!cName || cName.trim() === "") return;
        const colName = cName.trim();
        
        let targetGroup = null;
        let insertIdx = -1;

        if (AppState.selectedCell && typeof AppState.selectedCell === 'string') {
            const [rowI, selectedCol] = AppState.selectedCell.split('-');
            for (let g of TABLE_STRUCTURE) {
                const idx = g.columns.indexOf(selectedCol);
                if (idx !== -1) {
                    targetGroup = g;
                    insertIdx = idx + 1;
                    break;
                }
            }
        }

        if (!targetGroup) {
            targetGroup = TABLE_STRUCTURE.find(g => g.group === 'Zusatz');
            if (!targetGroup) {
                targetGroup = { group: 'Zusatz', class: 'zusatz', columns: [] };
                TABLE_STRUCTURE.push(targetGroup);
            }
            insertIdx = targetGroup.columns.length;
        }

        if (!targetGroup.columns.includes(colName)) {
            pushState();
            targetGroup.columns.splice(insertIdx, 0, colName);
            AppState.newCols.add(colName);
            AppState.data.forEach(row => { if(row[colName] === undefined) row[colName] = ""; });
            renderTable();
            saveToStorage();
            showToast(`✓ Spalte '${colName}' eingefügt.`);
        } else {
            showToast("⚠️ Spalte existiert bereits!");
        }
    });

    safeOn('btnDeleteCol', 'click', (e) => {
        if(e) e.preventDefault();
        let colToDelete = null;
        if (AppState.selectedCell && typeof AppState.selectedCell === 'string') {
            colToDelete = AppState.selectedCell.split('-')[1];
        }
        if (!colToDelete) {
            showToast("⚠️ Bitte zuerst eine Spalte auswählen!");
            return;
        }
        let foundGroup = null;
        for (let g of TABLE_STRUCTURE) {
            if (g.columns.includes(colToDelete)) {
                foundGroup = g;
                break;
            }
        }
        if (foundGroup) {
            if (foundGroup.class !== 'zusatz') {
                alert("Standard-Spalten (Basis, 0.8, Potential, etc.) dürfen nicht gelöscht werden!");
                return;
            }
            if (!confirm(`Möchten Sie '${colToDelete}' wirklich löschen?`)) return;
            pushState();
            const idx = foundGroup.columns.indexOf(colToDelete);
            foundGroup.columns.splice(idx, 1);
            AppState.newCols.delete(colToDelete);
            renderTable();
            saveToStorage();
            showToast(`✓ Spalte '${colToDelete}' gelöscht.`);
        }
    });


    safeOn('btnTabTable', 'click', () => switchTab('table'));
    safeOn('btnTabPlot', 'click', () => { 
        switchTab('plot');
        updateChartUI();
    });

    safeOn('btnChartBalken', 'click', () => { 
        AppState.chartType = 'balken'; 
        updateChartUI(); 
        renderAppPlot(); 
    });
    safeOn('btnChartScatter', 'click', () => { 
        AppState.chartType = 'scatter'; 
        updateChartUI(); 
        renderAppPlot(); 
    });
    safeOn('btnStdOn', 'click', () => { 
        AppState.showStd = true; 
        updateChartUI(); 
        renderAppPlot(); 
    });
    safeOn('btnStdOff', 'click', () => { 
        AppState.showStd = false; 
        updateChartUI(); 
        renderAppPlot(); 
    });

    safeOn('btnToggleMap', 'click', () => { 
        const m = document.getElementById('mapLayout');
        if(m) {
            m.style.display = 'flex';
            m.classList.add('show'); 
            
            // Karten-Redraw erzwingen (Force Redraw)
            if (!map) {
                initMap();
            }

            setTimeout(() => {
                if(map) {
                    map.invalidateSize();
                    map.setView(map.getCenter(), map.getZoom(), { animate: false });
                    setTimeout(() => map.invalidateSize(), 500);
                    refreshMapVisibility(); // ✔ Sync visibility on load
                    
                    // AUTO-START GPS WATCHING
                    // Automatic GPS start removed to prevent permission prompt
                    
                    showToast("Karte & Audit geladen ✓");
                }
            }, 300);
        }
    });
    safeOn('btnCloseMap', 'click', () => {
        const m = document.getElementById('mapLayout');
        if(m) {
            m.style.display = 'none';
            m.classList.remove('show');
        }
    });
    
    // Sidebar Map Buttons
    safeOn('btnMapOffline', 'click', () => switchLayer('standard'));
    safeOn('btnMapOnline', 'click', () => switchLayer('satellite'));

    // Top Toolbar Map Buttons (Header)
    safeOn('btnMapStandard', 'click', () => switchLayer('standard'));
    safeOn('btnMapSatellite', 'click', () => switchLayer('satellite'));

    // Global Visibility Sync for Toolbar Toggles (if any)
    safeOn('btnToolbarToggle08', 'click', () => toggleLayerColor('#ff00ff', 'btnToggle08'));
    safeOn('btnToolbarToggle16', 'click', () => toggleLayerColor('#ffff00', 'btnToggle16'));
    safeOn('btnToolbarToggle32', 'click', () => toggleLayerColor('#00ffff', 'btnToggle32'));
    
    safeOn('btnRefreshTable', 'click', () => {
        const pId = localStorage.getItem('current_project_id');
        if (pId) loadProject(pId);
        else renderTable();
        showToast("✓ Tabelle neu geladen");
    });

    safeOn('btnUndo', 'click', undo);
    safeOn('btnRedo', 'click', redo);

    safeOn('btnUploadPlan', 'click', () => document.getElementById('inputPlanImage').click());
    safeOn('inputPlanImage', 'change', (e) => {
        const file = e.target.files[0];
        if(!file) return;
        const url = URL.createObjectURL(file);
        const img = new Image();
        img.onload = () => {
            // ONLY remove tile layers, NOT drawItems
            if (layers.standard) map.removeLayer(layers.standard);
            if (layers.satellite) map.removeLayer(layers.satellite);
            if (layers.offlinePlan) map.removeLayer(layers.offlinePlan);

            const center = map.getCenter();
            const latOffset = (img.height / 20) / 111320; 
            const lngOffset = (img.width / 20) / (111320 * Math.cos(center.lat * Math.PI / 180));
            const boundsLL = L.latLngBounds([center.lat - latOffset, center.lng - lngOffset], [center.lat + latOffset, center.lng + lngOffset]);
            
            // Add as a proper layer that stays BELOW markers
            layers.offlinePlan = L.imageOverlay(url, boundsLL, { 
                zIndex: 1, 
                interactive: false 
            }).addTo(map);

            map.fitBounds(boundsLL);
            toggleMapUI(true); // Clean UI for Offline Plan
            showToast("✓ Offline-Plan geladen. Audit-Tools aktiv!");
        };
        img.src = url;
    });

    safeOn('drawColorPicker', 'input', (e) => {
        const c = e.target.value;
        if(window.drawControl) {
            drawControl.setDrawingOptions({
                polyline: { shapeOptions: { color: c, weight: 3 } },
                polygon: { shapeOptions: { color: c } },
                rectangle: { shapeOptions: { color: c } },
                circle: { shapeOptions: { color: c } },
                marker: { 
                    icon: L.divIcon({
                        html: `<i class="fas fa-map-marker-alt" style="color: ${c}; font-size: 30px; text-shadow: 0 0 3px white;"></i>`,
                        iconSize: [30, 30],
                        iconAnchor: [15, 30],
                        className: 'custom-marker-icon'
                    })
                }
            });
            showToast(`Farbe: ${c}`);
        }
    });

    safeOn('btnGoToLocation', 'click', handleMapSearch);
    safeOn('inpMapSearch', 'keypress', (e) => { if(e.key === 'Enter') handleMapSearch(); });

    let searchMarker = null;
    async function handleMapSearch() {
        if(searchMarker) map.removeLayer(searchMarker);
        const query = document.getElementById('inpMapSearch').value.trim();
        if(!query) return;
        const cleanQuery = query.replace(',', '.'); 
        const coordMatch = cleanQuery.match(/^(-?\d+\.?\d*)\s*[\s,]\s*(-?\d+\.?\d*)$/);
        if(coordMatch) {
            const lat = parseFloat(coordMatch[1]), lng = parseFloat(coordMatch[2]);
            map.setView([lat, lng], 19);
            searchMarker = L.marker([lat, lng]).addTo(map).bindPopup("Koordinaten: " + query).openPopup();
            showToast(`✓ Koordinaten gefunden: ${lat}, ${lng}`);
            return;
        }
        try {
            document.getElementById('loadingOverlay').style.display = 'flex';
            const resp = await fetch(`https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(query)}`);
            const results = await resp.json();
            document.getElementById('loadingOverlay').style.display = 'none';
            if(results.length > 0) {
                const res = results[0];
                map.setView([res.lat, res.lon], 19);
                searchMarker = L.marker([res.lat, res.lon]).addTo(map).bindPopup(res.display_name).openPopup();
                showToast(`✓ Standort gefunden: ${res.display_name.split(',')[0]}`);
            } else { showToast("⚠️ Standort nicht gefunden!"); }
        } catch(e) { document.getElementById('loadingOverlay').style.display = 'none'; showToast("❌ Fehler: " + e.message); }
    }

    safeOn('btnLocateMe', 'click', () => {
        if(!map) return;
        AppState.liveFollow = true; 
        AppState.firstLocationFound = false; // Reset to trigger zoom-in effect
        startLiveTracking();
    });

    function initCompass() {
        const compass = document.getElementById('mapCompass');
        if (!compass) return;

        const handleOrientation = (event) => {
            let heading = event.webkitCompassHeading || (360 - event.alpha);
            if (heading !== undefined && heading !== null) {
                // Rotate the whole compass or just the needle? 
                // Usually rotating the whole disc while needle stays N is a "Navigation" feel.
                // But rotating the needle to point N while you turn is more common for field work.
                compass.style.transform = `rotate(${-heading}deg)`;
            }
        };

        if (typeof DeviceOrientationEvent !== 'undefined' && typeof DeviceOrientationEvent.requestPermission === 'function') {
            // iOS 13+
            DeviceOrientationEvent.requestPermission()
                .then(response => {
                    if (response === 'granted') {
                        window.addEventListener('deviceorientation', handleOrientation, true);
                    } else {
                        showToast("⚠️ Kompas-Berechtigung verweigert.");
                    }
                })
                .catch(console.error);
        } else {
            // Android / Others
            window.addEventListener('deviceorientationabsolute', handleOrientation, true);
            window.addEventListener('deviceorientation', handleOrientation, true);
        }
    }
    safeOn('btnSetStartPoint', 'click', setStartPoint);
    safeOn('btnStopAudit', 'click', stopDrawing);

    // Audit Buttons (Pins & Lines) - SYNC WITH PICKER & REPEAT MODE
    const syncPicker = (c) => {
        const p = document.getElementById('drawColorPicker');
        if(p) { p.value = c; p.dispatchEvent(new Event('input')); }
        AppState.activeColor = c;
    };

    safeOn('btnMarker08', 'click', () => { syncPicker('#FF00FF'); setMarkerByDepth('#FF00FF'); });
    safeOn('btnMarker16', 'click', () => { syncPicker('#FFFF00'); setMarkerByDepth('#FFFF00'); });
    safeOn('btnMarker32', 'click', () => { syncPicker('#00FFFF'); setMarkerByDepth('#00FFFF'); });
    safeOn('btnLine08', 'click', () => { syncPicker('#FF00FF'); setLineByDepth('#FF00FF', 6); });
    safeOn('btnLine16', 'click', () => { syncPicker('#FFFF00'); setLineByDepth('#FFFF00', 6); });
    safeOn('btnLine32', 'click', () => { syncPicker('#00FFFF'); setLineByDepth('#00FFFF', 6); });

    // Numbered Markers (1-10) - TETAP AKTIF SETELAH DIPAKAI
    document.querySelectorAll('.num-btn').forEach(btn => {
        btn.onclick = () => {
            const num = btn.dataset.num;
            AppState.activeDrawToolName = `num-${num}`;
            showToast(`Sistem: Platziere Nummerierung ${num} (Mehrfachplatzierung aktiv)`);
            activateNumberedPlacement(num);
        };
    });

    safeOn('btnSnipRect', 'click', () => { 
        stopDrawing(true);
        captureMode=true; 
        AppState.activeDrawTool = new L.Draw.Rectangle(map, { repeatMode: false, showRadius: false });
        AppState.activeDrawTool.enable();
    });
    safeOn('btnSnipCircle', 'click', () => { 
        stopDrawing(true);
        captureMode=true; 
        AppState.activeDrawTool = new L.Draw.Circle(map, { repeatMode: true });
        AppState.activeDrawTool.enable();
    });
    safeOn('btnSnipPoly', 'click', () => { 
        stopDrawing(true);
        captureMode=true; 
        AppState.activeDrawTool = new L.Draw.Polygon(map, { repeatMode: true });
        AppState.activeDrawTool.enable();
    });
    
    safeOn('btnClearShapes', 'click', () => { 
        if (!window.drawControl) return;
        const deleter = new L.EditToolbar.Delete(map, {
            featureGroup: layers.drawItems
        });
        deleter.enable();
        showToast("🧹 LÖSCH-MODUS: Objek zum Löschen anklicken & 'Save' drücken.");
    });
    safeOn('btnDeleteArea', 'click', () => {
        stopDrawing(true);
        window.isAreaDeleteMode = true;
        AppState.activeDrawTool = new L.Draw.Rectangle(map, { 
            shapeOptions: { color: '#ef4444', fillOpacity: 0.2, weight: 2, dashArray: '5, 5' } 
        });
        AppState.activeDrawTool.enable();
        showToast("🚮 BEREICH LÖSCHEN: Rechteck über zu löschende Objekte ziehen.");
    });
    safeOn('btnSavePNG', 'click', () => { triggerAusschnitt(); });
    // Removed duplicate listeners
    safeOn('btnClearSave', 'click', () => { 
        if(confirm("Möchten Sie ALLES LÖSCHEN? (Alle Daten werden unwiderruflich gelöscht)")) { 
            localStorage.clear(); 
            location.reload(); 
        } 
    });

    safeOn('btnSnipText', 'click', addTextToSnip);
    safeOn('snipTargetType', 'change', (e) => {
        const val = e.target.value;
        const indicator = document.getElementById('snipColorIndicator');
        if(!indicator) return;
        if(val === 'Anhang_0.8') indicator.style.background = '#ff00ff';
        else if(val === 'Anhang_1.6') indicator.style.background = '#ffff00';
        else if(val === 'Anhang_3.2') indicator.style.background = '#00ffff';
        else if(val === 'Anhang_Global') indicator.style.background = '#ffffff';
    });
    safeOn('btnSnipRetry', 'click', () => { document.getElementById('snipConfirmModal').style.display='none'; if(AppState.drawItems) AppState.drawItems.clearLayers(); });
    safeOn('btnDiscardSnip', 'click', () => document.getElementById('snipConfirmModal').style.display='none');
    safeOn('btnSaveSnip', 'click', saveFinalSnip);
    safeOn('btnSnipDownload', 'click', () => {
        const canvas = document.getElementById('snipEditorCanvas');
        const link = document.createElement('a');
        link.download = `Audit_Pro_${new Date().getTime()}.png`;
        link.href = canvas.toDataURL("image/png");
        link.click();
        showToast("✓ Anhang heruntergeladen!");
    });

    safeOn('btnUndo', 'click', undo);
    safeOn('btnRedo', 'click', redo);
    safeOn('btnColManager', 'click', openColManager);
    safeOn('btnCloseColManager', 'click', () => document.getElementById('colManagerModal').style.display = 'none');
    safeOn('btnRefreshTable', 'click', () => { renderTable(); showToast("✓ Tabelle aktualisiert"); });



    // Ansicht / Zoom Management
    const applyZoom = () => {
        const val = AppState.zoomLevel;
        const display = document.getElementById('zoomLevelDisplay');
        if(display) display.innerText = val + '%';
        
        const table = document.getElementById('mainTable');
        if(table) {
            const scale = val / 100;
            // Use transform for smooth scaling on iPad
            table.style.transform = `scale(${scale})`;
            table.style.transformOrigin = 'top left';
            
            // Adjust container height to account for scaling
            const container = table.parentElement;
            if (container) {
                container.style.overflow = 'auto';
                // Force a re-layout
                table.style.display = 'none';
                table.offsetHeight; 
                table.style.display = 'table';
            }
            
            showToast(`Ansicht: ${val}%`, 500);
        }
    };

    safeOn('btnZoomIn', 'click', () => {
        AppState.zoomLevel = Math.min(200, AppState.zoomLevel + 10);
        applyZoom();
    });
    safeOn('btnZoomOut', 'click', () => {
        AppState.zoomLevel = Math.max(50, AppState.zoomLevel - 10);
        applyZoom();
    });
}

function switchTab(tab) {
    const btnTable = document.getElementById('btnTabTable');
    const btnPlot = document.getElementById('btnTabPlot');
    const tabTable = document.getElementById('tabTable');
    const tabPlot = document.getElementById('tabPlot');

    if (tab === 'table') {
        btnTable.classList.add('active');
        btnPlot.classList.remove('active');
        tabTable.classList.add('active');
        tabPlot.classList.remove('active');
        tabTable.style.display = 'block';
        tabPlot.style.display = 'none';
    } else {
        btnTable.classList.remove('active');
        btnPlot.classList.add('active');
        tabTable.classList.remove('active');
        tabPlot.classList.add('active');
        tabTable.style.display = 'none';
        tabPlot.style.display = 'block';
        setTimeout(renderAppPlot, 50); // Small delay to ensure container is visible for canvas scaling
    }
}

// Old setStartPoint removed to avoid conflict

function stopDrawing(silent = false) {
    if (AppState.activeDrawTool && AppState.activeDrawTool.disable) {
        AppState.activeDrawTool.disable();
    }
    AppState.activeDrawTool = null;
    AppState.activeDrawToolName = null;
    
    // Bersihkan SEMUA click listener manual
    map.off('click', handleStickyMarkerClick);
    map.off('click', handleNumberedClick);
    if (AppState._currentNumHandler) {
        map.off('click', AppState._currentNumHandler);
        AppState._currentNumHandler = null;
    }
    
    document.getElementById('map').style.cursor = '';
    if(!silent) showToast("✓ Audit-Modus beendet");
}

function setLineByDepth(color, weight = 6) {
    stopDrawing(true);
    if(!map) { showToast("⚠️ Bitte zuerst Karte öffnen!"); return; }
    
    AppState.activeDrawTool = new L.Draw.Polyline(map, { 
        repeatMode: true,
        shapeOptions: { 
            color: color, 
            weight: weight,
            lineCap: 'round',
            lineJoin: 'round'
        }
    });
    AppState.activeDrawTool.enable();
    document.getElementById('map').style.cursor = 'crosshair';
    showToast(`📏 ${color === '#FF00FF' ? 'PINK' : color === '#FFFF00' ? 'GELB' : 'BLAU'} Linie aktiv.`);
}

function setMarkerByDepth(color) {
    stopDrawing(true);
    if(!map) { showToast("⚠️ Bitte zuerst Karte öffnen!"); return; }

    AppState.activeDrawToolName = 'sticky-marker';
    AppState.activeColor = color;
    document.getElementById('map').style.cursor = 'crosshair';
    map.on('click', handleStickyMarkerClick);
    showToast(`📍 ${color === '#FF00FF' ? 'PINK' : color === '#FFFF00' ? 'GELB' : 'BLAU'} Pin aktiv.`);
}

function handleStickyMarkerClick(e) {
    const color = AppState.activeColor;
    if (!layers.drawItems) {
        layers.drawItems = new L.FeatureGroup().addTo(map);
    }
    
    // Kembali ke bentuk PIN SVG yang profesional
    const pinSvg = `
        <svg width="32" height="42" viewBox="0 0 32 42" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M16 0C7.16344 0 0 7.16344 0 16C0 28 16 42 16 42C16 42 32 28 32 16C32 7.16344 24.8366 0 16 0Z" fill="${color}"/>
            <circle cx="16" cy="16" r="6" fill="white"/>
        </svg>`;

    const marker = L.marker(e.latlng, {
        draggable: true,
        interactive: true,
        zIndexOffset: 1000,
        markerColor: color,   // <-- SAVED for visibility toggle
        icon: L.divIcon({
            html: `<div style="width:32px; height:42px; filter: drop-shadow(0 3px 5px rgba(0,0,0,0.5));">${pinSvg}</div>`,
            iconSize: [32, 42], iconAnchor: [16, 42], className: 'custom-marker-icon'
        })
    });
    
    setupInteractiveLayer(marker);
    marker.addTo(layers.drawItems);
    saveMapData(); // ✔ Auto-persist
}

function activateNumberedPlacement(num) {
    stopDrawing(true);
    if(!map) { showToast("⚠️ Bitte zuerst Karte öffnen!"); return; }
    AppState.activeDrawToolName = `num-${num}`;
    document.getElementById('map').style.cursor = 'crosshair';
    
    const handler = (e) => handleNumberedClick(e, num);
    map.on('click', handler);
    AppState._currentNumHandler = handler;
    
    showToast(`🔢 Nummerierung ${num} aktiv.`);
}

function handleNumberedClick(e, num) {
    const color = AppState.activeColor;
    if (!layers.drawItems) {
        layers.drawItems = new L.FeatureGroup().addTo(map);
    }
    
    const marker = L.marker(e.latlng, {
        draggable: true,
        interactive: true,
        zIndexOffset: 1000,
        markerColor: color,
        markerNumber: num,      // ✔ Saved for persistence
        isNumbered: true,       // ✔ Saved for persistence
        icon: L.divIcon({
            html: `<div style="background:${color} !important; color:black !important; width:30px; height:30px; border-radius:50%; display:flex; align-items:center; justify-content:center; border:3px solid #808080; font-weight:bold; font-size:16px; box-shadow: 0 0 10px rgba(0,0,0,0.5);">${num}</div>`,
            iconSize: [30, 30], iconAnchor: [15, 15], className: 'numbered-marker'
        })
    });
    
    setupInteractiveLayer(marker);
    marker.addTo(layers.drawItems);
    saveMapData(); // ✔ Auto-persist
}

function initMap() {
    if (map) return;
    
    const mapContainer = document.getElementById('map');
    if (!mapContainer) return;

    // Inisialisasi Dasar Leaflet
    map = L.map('map', {
        zoomControl: false,
        attributionControl: true,
        doubleClickZoom: false,
        preferCanvas: true
    }).setView([51.1657, 10.4515], 6);

    AppState.map = map;
    L.control.zoom({ position: 'bottomleft' }).addTo(map);

    // Matikan SEMUA tulisan instruksi di kursor (Tooltips)
    L.drawLocal.draw.handlers.rectangle.tooltip.start = '';
    L.drawLocal.draw.handlers.circle.tooltip.start = '';
    L.drawLocal.draw.handlers.polygon.tooltip.start = '';
    L.drawLocal.draw.handlers.polyline.tooltip.start = '';
    L.drawLocal.draw.handlers.simpleshape.tooltip.end = '';

    // Register global layers - RESTORED FULL FEATURES
    const googleRoadmap = L.tileLayer('https://{s}.google.com/vt/lyrs=m&x={x}&y={y}&z={z}', {
        maxZoom: 22,
        maxNativeZoom: 20,
        subdomains:['mt0','mt1','mt2','mt3'],
        attribution: '&copy; Google Maps HD'
    });
    
    const googleHybrid = L.tileLayer('https://{s}.google.com/vt/lyrs=y&x={x}&y={y}&z={z}', {
        maxZoom: 22,
        maxNativeZoom: 20,
        subdomains:['mt0','mt1','mt2','mt3'],
        attribution: '&copy; Google Maps HD'
    });

    layers.standard = googleRoadmap;
    layers.satellite = googleHybrid;

    // Add default directly
    googleRoadmap.addTo(map);

    layers.drawItems = new L.FeatureGroup();
    map.addLayer(layers.drawItems);
    
    window.drawControl = new L.Control.Draw({
        edit: {
            featureGroup: layers.drawItems,
            remove: true,
            edit: false
        },
        draw: { 
            polygon: { shapeOptions: { color: '#3b82f6' }, repeatMode: true },
            polyline: { shapeOptions: { color: '#3b82f6', weight: 4 }, repeatMode: true },
            rectangle: { shapeOptions: { color: '#3b82f6' }, repeatMode: true },
            circle: { shapeOptions: { color: '#3b82f6' }, repeatMode: true },
            marker: { repeatMode: true },
            circlemarker: false
        },
        position: 'topright'
    });

    map.on(L.Draw.Event.CREATED, (e) => {
        const layer = e.layer;
        const color = AppState.activeColor || document.getElementById('drawColorPicker').value;

        if (captureMode) {
            layer.setStyle({ color: 'transparent', fillOpacity: 0 });
            map.addLayer(layer); 
            triggerAusschnitt(layer);
            captureMode = false;
            return;
        }

        if (window.isAreaDeleteMode) {
            const bounds = layer.getBounds();
            let count = 0;
            if (layers.drawItems) {
                layers.drawItems.eachLayer(l => {
                    let shouldDelete = false;
                    if (l instanceof L.Marker) {
                        if (bounds.contains(l.getLatLng())) shouldDelete = true;
                    } else if (l.getBounds) {
                        if (bounds.intersects(l.getBounds())) shouldDelete = true;
                    }

                    if (shouldDelete) {
                        layers.drawItems.removeLayer(l);
                        count++;
                    }
                });
            }
            window.isAreaDeleteMode = false;
            saveMapData();
            showToast(`✓ ${count} Objekte im Bereich gelöscht.`);
            return;
        }

        if (layer instanceof L.Marker) {
            const latlng = layer.getLatLng();
            const coloredMarker = L.marker(latlng, {
                icon: L.divIcon({
                    html: `<i class="fas fa-map-marker-alt" style="color: ${color}; font-size: 30px; text-shadow: 0 0 3px white;"></i>`,
                    iconSize: [30, 30],
                    iconAnchor: [15, 30],
                    className: 'custom-marker-icon'
                }),
                draggable: true
            });
            // STORE COLOR FOR PERSISTENCE
            coloredMarker.options.markerColor = color; 
            coloredMarker.addTo(layers.drawItems);
            setupInteractiveLayer(coloredMarker);
        } else {
            // FOR LINES AND SHAPES
            layer.setStyle({ color: color, weight: 8 }); // Make lines THICKER for easier tapping
            layer.options.markerColor = color; // Save color DNA
            layers.drawItems.addLayer(layer);
            setupInteractiveLayer(layer);
        }

        // REPEAT MODE ERZWINGEN
        if (AppState.activeDrawTool) {
            setTimeout(() => {
                if (AppState.activeDrawTool) AppState.activeDrawTool.enable();
            }, 50);
        }
        saveMapData(); // ✔ Auto-persist
    });

    map.on(L.Draw.Event.DELETED, saveMapData);
    map.on(L.Draw.Event.EDITED, saveMapData);
    
    // RESTORE SAVED DRAWINGS
    restoreMapDrawings();

    // AUTO-START GPS WATCHING
    if (AppState.liveFollow) {
        startLiveTracking();
    }

    map.on('dragstart', () => {
        if(AppState.liveFollow) {
            AppState.liveFollow = false;
            showToast("Auto-Follow AUS (Manueller Modus)");
        }
    });

    window.addEventListener('deviceorientationabsolute', (e) => {
        let h = e.webkitCompassHeading || (360 - e.alpha);
        if (h) {
            if(document.getElementById('userHeadingBeam')) document.getElementById('userHeadingBeam').style.transform = `rotate(${h}deg)`;
            if(document.getElementById('windrose')) document.getElementById('windrose').style.transform = `rotate(${-h}deg)`;
        }
    }, true);
}

function setupInteractiveLayer(layer) {
    const deleteAction = (e) => {
        L.DomEvent.stopPropagation(e);
        if (confirm("Dieses Objekt (Marker/Linie/Nummer) löschen?")) {
            if (layers.drawItems) {
                layers.drawItems.removeLayer(layer);
                saveMapData(); // ✔ Auto-persist on delete too
                showToast("✓ Objekt gelöscht");
            }
        }
    };

    layer.on('click', deleteAction);
    layer.on('dblclick', deleteAction); // Support double-click/double-tap
}

// (saveMapDrawings and restoreMapDrawings are defined further below,
//  using messstellen_library for correct per-project persistence)

function switchLayer(key) {
    if (!map) return;
    
    // Alle alten Ebenen entfernen
    if (layers.standard) map.removeLayer(layers.standard);
    if (layers.satellite) map.removeLayer(layers.satellite);
    if (layers.offlinePlan) map.removeLayer(layers.offlinePlan);
    toggleMapUI(false); // Restore UI for Online Maps

    // Tambahkan yang baru
    if (key === 'satellite') {
        layers.satellite.addTo(map);
        showToast("✓ GOOGLE HD SATELIT");
    } else {
        layers.standard.addTo(map);
        showToast("✓ STANDARD KARTE");
    }

    // Karten-Redraw erzwingen (Force Redraw)
    setTimeout(() => {
        map.invalidateSize();
        const currentZoom = map.getZoom();
        map.setZoom(currentZoom + 0.00001);
    }, 100);

    // Update Visual untuk SEMUA Tombol (Sidebar & Top Overlay)
    const status = document.getElementById('statusIndicator');
    const btnOff = document.getElementById('btnMapOffline'), btnOn = document.getElementById('btnMapOnline');
    const btnTopStd = document.getElementById('btnMapStandard'), btnTopSat = document.getElementById('btnMapSatellite');

    if (key === 'standard') {
        btnOff.classList.add('active'); btnOn.classList.remove('active');
        btnTopStd.style.background = '#3b82f6'; btnTopSat.style.background = 'rgba(59, 130, 246, 0.1)';
    } else {
        btnOn.classList.add('active'); btnOff.classList.remove('active');
        btnTopSat.style.background = '#3b82f6'; btnTopStd.style.background = 'rgba(59, 130, 246, 0.1)';
    }
}

window.toggleLayerColor = function(color, triggerBtnId) {
    if (!AppState.hiddenMapColors) AppState.hiddenMapColors = new Set();
    const colorLower = color.toLowerCase();
    const isHidden = AppState.hiddenMapColors.has(colorLower);
    const depthKey = colorLower === '#ff00ff' ? '08' : (colorLower === '#ffff00' ? '16' : '32');

    // Sync map layers
    if (layers.drawItems) {
        layers.drawItems.eachLayer(layer => {
            const c = getLayerColor(layer);
            if (!c || c.toLowerCase() !== colorLower) return;
            if (isHidden) {
                const el = layer.getElement ? layer.getElement() : null;
                if (el) el.style.display = '';
                if (layer.setStyle) layer.setStyle({ opacity: 1, fillOpacity: 0.5 });
                if (layer.setOpacity) layer.setOpacity(1);
            } else {
                const el = layer.getElement ? layer.getElement() : null;
                if (el) el.style.display = 'none';
                if (layer.setStyle) layer.setStyle({ opacity: 0, fillOpacity: 0 });
                if (layer.setOpacity) layer.setOpacity(0);
            }
        });
    }

    // Identify all related buttons
    const btnIds = [`btnToggle${depthKey}`, `btnToolbarToggle${depthKey}`];
    
    if (isHidden) {
        AppState.hiddenMapColors.delete(colorLower);
        btnIds.forEach(id => {
            const btn = document.getElementById(id);
            if (btn) {
                btn.style.opacity = '1';
                const icon = btn.querySelector('i');
                if (icon) icon.className = 'fas fa-eye';
            }
        });
    } else {
        AppState.hiddenMapColors.add(colorLower);
        btnIds.forEach(id => {
            const btn = document.getElementById(id);
            if (btn) {
                btn.style.opacity = '0.4';
                const icon = btn.querySelector('i');
                if (icon) icon.className = 'fas fa-eye-slash';
            }
        });
    }

    renderTable();
    saveToStorage();
};

window.refreshMapVisibility = function() {
    if (!layers.drawItems || !AppState.hiddenMapColors) return;
    
    layers.drawItems.eachLayer(layer => {
        const color = getLayerColor(layer);
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

    // Sync button UI
    ['#ff00ff', '#ffff00', '#00ffff'].forEach(color => {
        const btnId = color === '#ff00ff' ? 'btnToggle08' : (color === '#ffff00' ? 'btnToggle16' : 'btnToggle32');
        const iconId = 'icon' + btnId.replace('btnToggle', 'Toggle');
        const isHidden = AppState.hiddenMapColors.has(color);
        const icon = document.getElementById(iconId);
        const btn = document.getElementById(btnId);
        if (icon) icon.className = isHidden ? 'fas fa-eye-slash' : 'fas fa-eye';
        if (btn) btn.style.opacity = isHidden ? '0.4' : '1';
    });
}

function getLayerColor(layer) {
    if (layer.options && layer.options.markerColor) return layer.options.markerColor;
    if (layer.feature && layer.feature.properties && layer.feature.properties.markerColor)
        return layer.feature.properties.markerColor;
    if (layer.options && layer.options.color) return layer.options.color;
    return null;
}

function colorLabel(color) {
    if (color.toLowerCase() === '#ff00ff') return 'Pink (0.8m)';
    if (color.toLowerCase() === '#ffff00') return 'Gelb (1.6m)';
    if (color.toLowerCase() === '#00ffff') return 'Cyan (3.2m)';
    return color;
}
// ───────────────────────────────────────────────────────────────────────────

function toggleLiveFollow() {
    AppState.liveFollow = !AppState.liveFollow;
    if(AppState.liveFollow) { 
        AppState.map.locate({
            watch: true, 
            enableHighAccuracy: true,
            maximumAge: 1000,
            timeout: 10000,
            onLocationError: (e) => {
                const status = document.getElementById('gpsStatusText');
                if (status) status.innerText = "GPS Fehler: " + e.message;
            }
        }); 
        showToast('📍 Live-Verfolgung: AN'); 
    }
    else { 
        AppState.map.stopLocate(); 
        showToast('📍 Live-Verfolgung: AUS'); 
    }
}

function setStartPoint() {
    if (!map) return;
    
    // Use userMarker position if available, otherwise map center
    const latlng = AppState.userMarker ? AppState.userMarker.getLatLng() : map.getCenter();
    AppState.startPoint = latlng;
    
    if (AppState.startMarker) map.removeLayer(AppState.startMarker);
    AppState.startMarker = L.marker(latlng, {
        icon: L.divIcon({ 
            className: 'start-point-wrapper',
            html: `
                <div style="display:flex; align-items:center; gap:8px; white-space:nowrap;">
                    <div style="display:flex; align-items:center; justify-content:center; width:32px; height:32px; background:#000; border:2px solid #fff; border-radius:50%; box-shadow:0 0 15px rgba(0,0,0,0.8);">
                        <i class="fas fa-flag" style="color: #fff; font-size: 18px; font-weight:900;"></i>
                    </div>
                    <span style="background:#000; color:#fff; border:1px solid #fff; padding:4px 10px; border-radius:6px; font-weight:900; font-size:12px; letter-spacing:1px; box-shadow:0 4px 15px rgba(0,0,0,0.5);">STARTPUNKT</span>
                </div>
            `, 
            iconSize: [150, 32], 
            iconAnchor: [16, 16]
        }),
        zIndexOffset: 3000
    }).addTo(map);

    // Show distance immediately
    const distanceBadge = document.getElementById('distanceBadge');
    const distanceValue = document.getElementById('distanceValue');
    if(distanceBadge && distanceValue) {
        const dist = AppState.map.distance(latlng, latlng); // Initially 0
        distanceValue.innerText = "0.0";
        distanceBadge.style.display = 'flex';
    }

    showToast("🏁 Startpunkt gesetzt!");
}

async function triggerAusschnitt(targetLayer) {
    console.log("Triggering Ausschnitt...");
    if (!targetLayer) {
        if(!layers.drawItems) {
            showToast("⚠️ Keine Zeichnungen vorhanden!");
            return;
        }
        const layersG = layers.drawItems.getLayers();
        targetLayer = layersG[layersG.length - 1];
    }
    
    if (!targetLayer) { 
        showToast("⚠️ Bitte zuerst einen Bereich (Rechteck/Kreis) markieren!"); 
        return; 
    }

    showToast("📸 Foto wird vorbereitet...", 2000);
    document.getElementById('loadingOverlay').style.display = 'flex';
    
    const rowSel = document.getElementById('snipTargetRow');
    const typeSel = document.getElementById('snipTargetType');
    rowSel.innerHTML = '<option value="">-- ZEILE HIER WÄHLEN --</option>';
    
    // RE-BUILD DROPDOWN TO BE 100% CERTAIN
    typeSel.innerHTML = `
        <option value="Anhang_0.8">0.8m</option>
        <option value="Anhang_1.6">1.6m</option>
        <option value="Anhang_3.2">3.2m</option>
        <option value="Anhang_Global">Anhang (Gesamt)</option>
    `;
    // AUTO-DETECT CATEGORY BASED ON VISIBILITY
    const visibleColors = [];
    if (!AppState.hiddenMapColors.has('#ff00ff')) visibleColors.push('0.8');
    if (!AppState.hiddenMapColors.has('#ffff00')) visibleColors.push('1.6');
    if (!AppState.hiddenMapColors.has('#00ffff')) visibleColors.push('3.2');

    if (visibleColors.length > 1) {
        typeSel.value = 'Anhang_Global';
        document.getElementById('snipColorIndicator').style.background = '#ffffff';
    } else if (visibleColors.length === 1) {
        typeSel.value = `Anhang_${visibleColors[0]}`;
        const c = visibleColors[0] === '0.8' ? '#ff00ff' : (visibleColors[0] === '1.6' ? '#ffff00' : '#00ffff');
        document.getElementById('snipColorIndicator').style.background = c;
    } else {
        typeSel.value = 'Anhang_Global';
        document.getElementById('snipColorIndicator').style.background = '#ffffff';
    }

    const lastActiveRow = AppState.selectedCell ? AppState.selectedCell.split('-')[0] : "";

    AppState.data.forEach((r, i) => {
        const opt = document.createElement('option'); opt.value = i;
        opt.textContent = `ZEILE ${i+1}: ${r.Kennzeichen || '---'}`;
        if (lastActiveRow !== "" && parseInt(lastActiveRow) === i) opt.selected = true;
        else if (lastActiveRow === "" && i === 0) opt.selected = true; // Default to first row
        rowSel.appendChild(opt);
    });

    try {
        const mapEl = document.getElementById('map');
        // Hide UI elements
        const controls = document.querySelectorAll('.leaflet-control-container, .gps-status-bar, .map-sidebar');
        controls.forEach(c => c.style.display = 'none');
        
        await new Promise(r => setTimeout(r, 800));

        const canvas = await html2canvas(mapEl, {
            useCORS: true,
            allowTaint: true,
            scale: 3, // High-Resolution Capture (300+ DPI equivalent)
            logging: false,
            backgroundColor: '#111'
        });

        // Restore UI elements
        controls.forEach(c => c.style.display = '');

        const cropped = cropToShape(canvas, targetLayer, mapEl);
        const editorCanvas = document.getElementById('snipEditorCanvas');
        editorCanvas.originalImage = cropped;
        editorCanvas.width = cropped.width; 
        editorCanvas.height = cropped.height;
        
        snipTexts = [];
        drawEditorCanvas();
        
        document.getElementById('loadingOverlay').style.display = 'none';
        document.getElementById('snipConfirmModal').style.display = 'block';
        
        setTimeout(() => {
            editorCanvas.style.width = '100%'; 
            editorCanvas.style.height = 'auto';
        }, 100);

    } catch(e) { 
        document.getElementById('loadingOverlay').style.display = 'none'; 
        const controls = document.querySelectorAll('.leaflet-control-container, .gps-status-bar, .map-sidebar');
        controls.forEach(c => c.style.display = '');
        showToast("❌ Foto-Fehler: " + e.message, 6000); 
        console.error(e);
    }
}



function cropToShape(fullCanvas, layer, mapEl) {
    const temp = document.createElement('canvas'), ctx = temp.getContext('2d');
    const sx = fullCanvas.width / mapEl.offsetWidth, sy = fullCanvas.height / mapEl.offsetHeight;
    let b = layer.getBounds ? layer.getBounds() : L.latLngBounds(layer.getLatLng(), layer.getLatLng());
    const nw = AppState.map.latLngToContainerPoint(b.getNorthWest()), se = AppState.map.latLngToContainerPoint(b.getSouthEast());
    
    // Potong PAS sesuai kotak (tanpa margin m)
    let w = Math.max(10, se.x - nw.x), h = Math.max(10, se.y - nw.y);
    temp.width = w * sx; temp.height = h * sy;
    ctx.drawImage(fullCanvas, nw.x * sx, nw.y * sy, temp.width, temp.height, 0, 0, temp.width, temp.height);
    return temp;
}

function addTextToSnip() {
    const t = prompt("Text:"); if(!t) return;
    snipTexts.push({ text: t, x: 100, y: 100, color: document.getElementById('snipTextColor').value, fontSize: 40 });
    drawEditorCanvas();
}

function initEditorEvents() {
    const canvas = document.getElementById('snipEditorCanvas');
    if(!canvas) return;

    const btnIn = document.getElementById('btnSnipZoomIn');
    const btnOut = document.getElementById('btnSnipZoomOut');
    if(btnIn) btnIn.style.cursor = 'pointer';
    if(btnOut) btnOut.style.cursor = 'pointer';

    safeOn('btnSnipZoomIn', 'click', () => {
        let w = canvas.offsetWidth;
        canvas.style.width = (w + 50) + 'px';
        console.log("Zoom In: " + canvas.style.width);
    });
    safeOn('btnSnipZoomOut', 'click', () => {
        let w = canvas.offsetWidth;
        canvas.style.width = Math.max(200, w - 50) + 'px';
        console.log("Zoom Out: " + canvas.style.width);
    });

    canvas.onmousedown = (e) => {
        const p = getMousePos(e, canvas);
        draggingTextIdx = snipTexts.findIndex(t => p.x > t.x && p.x < t.x + 200 && p.y > t.y - 40 && p.y < t.y);
    };
    canvas.onmousemove = (e) => {
        if(draggingTextIdx === -1) return;
        const p = getMousePos(e, canvas); snipTexts[draggingTextIdx].x = p.x; snipTexts[draggingTextIdx].y = p.y;
        drawEditorCanvas();
    };
    canvas.onmouseup = () => draggingTextIdx = -1;
    canvas.onwheel = (e) => {
        e.preventDefault();
        snipTexts.forEach(t => { t.fontSize += e.deltaY > 0 ? -2 : 2; t.fontSize = Math.max(10, t.fontSize); });
        drawEditorCanvas();
    };
}
window.addEventListener('load', initEditorEvents);

function getMousePos(e, canvas) {
    const r = canvas.getBoundingClientRect();
    return { x: (e.clientX - r.left)*(canvas.width/r.width), y: (e.clientY - r.top)*(canvas.height/r.height) };
}

function drawEditorCanvas() {
    const canvas = document.getElementById('snipEditorCanvas');
    const ctx = canvas.getContext('2d');
    if(!canvas.originalImage) return;
    ctx.clearRect(0,0,canvas.width,canvas.height);
    ctx.drawImage(canvas.originalImage, 0, 0);
    snipTexts.forEach(t => {
        ctx.font = `bold ${t.fontSize}px Inter`;
        ctx.shadowColor = "black"; ctx.shadowBlur = 5;
        ctx.strokeStyle = "black"; ctx.lineWidth = 3; ctx.strokeText(t.text, t.x, t.y);
        ctx.fillStyle = t.color; ctx.fillText(t.text, t.x, t.y);
    });
}

function saveFinalSnip() {
    const canvas = document.getElementById('snipEditorCanvas');
    const targetIdxRaw = document.getElementById('snipTargetRow').value;
    let targetCol = document.getElementById('snipTargetType').value;
    
    const targetIdx = parseInt(targetIdxRaw) || 0;
    if (!targetCol) targetCol = 'Anhang_Global';

    try {
        if (AppState.data && AppState.data[targetIdx]) {
            AppState.data[targetIdx][targetCol] = canvas.toDataURL();
            renderTable(); 
            saveToStorage();
        }
    } catch (e) {
        console.error("Speichern fehlgeschlagen:", e);
    } finally {
        const modal = document.getElementById('snipConfirmModal');
        if (modal) modal.style.display = 'none';
    }
}

function toggleMapUI(isOffline) {
    const gps = document.getElementById('gpsStatus');
    const compass = document.getElementById('mapCompass');
    const searchArea = document.getElementById('inpMapSearch')?.parentElement;
    const navArea = document.getElementById('btnLocateMe')?.parentElement;

    if (isOffline) {
        if(gps) gps.style.display = 'none';
        if(compass) compass.style.display = 'none';
        if(searchArea) searchArea.style.display = 'none';
        if(navArea) navArea.style.display = 'none';
    } else {
        if(gps) gps.style.display = 'flex';
        if(compass) compass.style.display = 'flex';
        if(searchArea) searchArea.style.display = 'block';
        if(navArea) navArea.style.display = 'grid';
    }
}

function handleFileUpload(file) {
    // RESET PREVIOUS STATE
    AppState.data = [];
    AppState.newCols.clear();
    if(map) {
        Object.values(layers).forEach(l => map.removeLayer(l));
    }

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
            let startIdx = -1;
            for (let i = 0; i < 15; i++) { 
                if (JSON.stringify(allRows[i]).toLowerCase().includes('kennzeichen')) { 
                    startIdx = i + 1; 
                    break; 
                } 
            }
            if (startIdx === -1) {
                showToast("⚠️ Ungültiges Datei-Format");
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
            
            // UI Updates for Welcome Screen
            const fileStatus = document.getElementById('fileStatus');
            if(fileStatus) fileStatus.innerText = `📂 DATEI: ${file.name.toUpperCase()}`;
            
            const btnStart = document.getElementById('btnFinalStart');
            if(btnStart) btnStart.disabled = false;

            const fileName = file.name;
            const firstWord = fileName.split(/[._\s-]/)[0].toUpperCase();
            const pInput = document.getElementById('projectName');
            if(pInput) pInput.value = firstWord;

            showToast(`✓ ${mapped.length} Messstellen bereit!`);
        } catch (err) { 
            showToast('Import-Fehler'); 
        }
    };
    reader.readAsArrayBuffer(file);
}

function updateLiveCalculations(row, col, tr) {
    if (!(col.startsWith('R') && col.includes('_'))) return;
    const dStr = col.split('_')[1], d = parseFloat(dStr), n = col.substring(1, 2);
    const rawVal = (row[col] || "").toString().trim().replace(',', '.');
    const rhoKey = `ρ${n} [Ωm]_${dStr}`;
    const mwKey = `MW [Ωm]_${dStr}`;
    const sdKey = `SD [Ωm]_${dStr}`;

    // 1. Update Rho for current input
    if (rawVal === "") {
        row[rhoKey] = "";
    } else {
        const rVal = parseFloat(rawVal);
        if (!isNaN(rVal) && !isNaN(d)) {
            row[rhoKey] = (2 * Math.PI * d * rVal).toFixed(2);
        }
    }

    // 2. Calculate MW and SD for the group
    let sum = 0, count = 0, values = [];
    for(let k=1; k<=3; k++) {
        const rkRaw = (row[`R${k} [Ω]_${dStr}`] || "").toString().trim().replace(',', '.');
        if (rkRaw !== "") {
            const rk = parseFloat(rkRaw);
            if(!isNaN(rk)) { 
                sum += rk; 
                count++; 
                values.push(rk);
            }
        }
    }

    if(count > 0) {
        const mwR = sum / count;
        row[mwKey] = (2 * Math.PI * d * mwR).toFixed(2);
        if(count > 1) {
            let sqSum = 0;
            values.forEach(v => sqSum += Math.pow(v - mwR, 2));
            row[sdKey] = (2 * Math.PI * d * Math.sqrt(sqSum / (count - 1))).toFixed(2);
        } else {
            row[sdKey] = ""; // Standard deviation requires at least 2 points
        }
    } else {
        row[mwKey] = "";
        row[sdKey] = "";
    }

    // 3. Update DOM for all affected inputs in this row
    tr.querySelectorAll('input').forEach(tInp => {
        const tc = tInp.dataset.col || '';
        if (tc !== col && (tc.startsWith('ρ') || tc.startsWith('MW') || tc.startsWith('SD'))) {
            tInp.value = row[tc] || "";
        }
    });
}

function renderTable(onlyBody = false) {
    const thead = document.getElementById('tableHead'), tbody = document.getElementById('tableBody');
    if (!thead || !tbody) return;

    if (!onlyBody || thead.innerHTML === "") {
        thead.innerHTML = '';
        const tr1 = document.createElement('tr'), tr2 = document.createElement('tr');
        tr1.innerHTML = '<th rowspan="2" style="border-right: 3px solid #3b82f6; background: #f8fafc; color: #1e293b; font-weight: 900;">#</th>';

        TABLE_STRUCTURE.forEach(g => {
            if(AppState.hiddenColumns.has(g.class)) return;
            const groupCols = g.columns.filter(c => !AppState.hiddenColumns.has(c));
            if (groupCols.length === 0) return;
            const cs = DEPTH_COLORS[g.class] || { bg: '#f3f4f6', border: '#3b82f6', text: '#1f2937' };
            
            const th = document.createElement('th');
            th.style.position = 'relative';
            th.style.borderLeft = `4px solid ${cs.border}`;
            if(cs) { th.style.backgroundColor = cs.bg; th.style.color = cs.text; th.style.borderBottom = `5px solid ${cs.border}`; }

            if (g.group === 'Anhang (Gesamt)') {
                th.rowSpan = 2;
                th.style.verticalAlign = 'middle';
                th.style.minWidth = AppState.columnWidths[g.columns[0]] || '80px';
                th.innerHTML = "ANHANG<br>(GESAMT)";
                tr1.appendChild(th);
                // DO NOT add to tr2
            } else {
                th.textContent = g.group; 
                th.colSpan = groupCols.length;
                tr1.appendChild(th);
                
                groupCols.forEach((c, cIdx) => {
                    const sh = document.createElement('th'); sh.style.position = 'relative';
                    sh.style.backgroundColor = cs.bg;
                    sh.style.color = cs.text;
                    
                    // Add colored underline for ALL headers to show categories clearly
                    sh.style.borderBottom = `3px solid ${cs.border}`;
                    
                    if (cIdx === 0) sh.style.borderLeft = `4px solid ${cs.border}`;
                    sh.style.minWidth = AppState.columnWidths[c] || '80px';
                    const label = document.createElement('span');
                    label.innerHTML = c.includes('_') ? c.split('_')[0].replace(' ', '<br>') : c;
                    sh.appendChild(label);

                    if (c === 'Kennzeichen') {
                        const fi = document.createElement('input');
                        fi.type = 'text'; fi.placeholder = 'Filter...'; fi.className = 'header-filter-input';
                        fi.value = AppState.filters['Kennzeichen'] || "";
                        fi.onclick = (e) => e.stopPropagation();
                        fi.oninput = (e) => {
                            AppState.filters['Kennzeichen'] = e.target.value;
                            renderTable(true);
                        };
                        sh.appendChild(fi);
                    }
                    tr2.appendChild(sh);
                });
            }
        });
        thead.append(tr1, tr2);
    }

    tbody.innerHTML = '';
    let filteredData = AppState.data.map((row, index) => ({ ...row, _originalIndex: index }));
    const kzFilter = (AppState.filters['Kennzeichen'] || "").trim().toLowerCase();
    if (kzFilter) {
        if (kzFilter.includes('-')) {
            const parts = kzFilter.split('-').map(p => p.trim());
            if (parts.length === 2) {
                const s = parts[0], e = parts[1];
                filteredData = filteredData.filter(row => {
                    const v = (row['Kennzeichen'] || "").toString().toLowerCase();
                    if(!isNaN(s) && !isNaN(e) && !isNaN(v)) return Number(v) >= Number(s) && Number(v) <= Number(e);
                    return v >= s && v <= e;
                });
            }
        } else {
            filteredData = filteredData.filter(row => (row['Kennzeichen'] || "").toString().toLowerCase().includes(kzFilter));
        }
    }

    filteredData.forEach((row) => {
        const tr = document.createElement('tr');
        const idx = row._originalIndex;
        const originalRow = AppState.data[idx];
        tr.innerHTML = `<td style="border-right: 3px solid #3b82f6; background: #f8fafc; color: #1e293b; font-weight: 900; text-align: center;">${idx + 1}</td>`;
        
        TABLE_STRUCTURE.forEach(g => {
            if (AppState.hiddenColumns.has(g.class)) return;
            const groupCols = g.columns.filter(c => !AppState.hiddenColumns.has(c));
            const cs = DEPTH_COLORS[g.class] || { bg: 'transparent', border: '#3b82f6' };
            
            groupCols.forEach((col, colIdx) => {
                const td = document.createElement('td');
                const cs = DEPTH_COLORS[g.class] || { bg: 'transparent', border: '#3b82f6' };
                
                td.style.backgroundColor = 'transparent';
                // Professional white-ish separator for data rows
                td.style.borderBottom = '1px solid rgba(255,255,255,0.15)';
                
                if (colIdx === 0) td.style.borderLeft = `4px solid ${cs.border}`;
                
                if(col.toLowerCase().includes('anhang')) {
                    td.style.padding = '0';
                    td.style.borderRight = `1px solid rgba(255,255,255,0.05)`;
                    
                    const wrapper = document.createElement('div');
                    wrapper.style.display = 'flex';
                    wrapper.style.gap = '6px';
                    wrapper.style.justifyContent = 'center';
                    wrapper.style.alignItems = 'center';
                    wrapper.style.height = '32px';
                    wrapper.style.width = '100%';
                    wrapper.style.whiteSpace = 'nowrap';
                    wrapper.style.backgroundColor = 'rgba(255,255,255,0.02)';

                    if(originalRow[col]) {
                        const b = document.createElement('button'); 
                        b.className='btn-modern'; 
                        b.innerHTML='👁️';
                        b.style.padding = '2px 8px'; 
                        b.style.height = '24px';
                        b.onclick = () => openImagePreview(originalRow[col], "Vorschau");
                        wrapper.appendChild(b);

                        const delBtn = document.createElement('button'); 
                        delBtn.className='btn-modern btn-danger'; 
                        delBtn.innerHTML='✕';
                        delBtn.style.padding = '2px 8px';
                        delBtn.style.height = '24px';
                        delBtn.onclick = () => {
                            if(confirm("Möchten Sie dieses Foto wirklich löschen?")) {
                                delete originalRow[col]; renderTable(true); saveToStorage();
                                showToast("✓ Foto gelöscht");
                            }
                        };
                        wrapper.appendChild(delBtn);
                    } else {
                        wrapper.innerHTML = '-';
                        wrapper.style.color = 'var(--text-dim)';
                    }
                    td.appendChild(wrapper);
                } else {
                    const inp = document.createElement('input');
                    inp.type = 'text'; 
                    inp.value = originalRow[col] || ""; 
                    inp.className = 'cell-input';
                    inp.style.color = '#ffffff';
                    inp.style.backgroundColor = 'transparent';
                    
                    if (originalRow._isNew || AppState.newCols.has(col)) {
                        inp.style.backgroundColor = 'rgba(210, 180, 140, 0.3)';
                    }

                    inp.dataset.col = col;
                    inp.dataset.row = idx;
                    if (typeof cellFormatting !== 'undefined' && cellFormatting[`${idx}-${col}`]) {
                        if (typeof applyFormattingToCell === 'function') applyFormattingToCell(inp, cellFormatting[`${idx}-${col}`]);
                    }
                    if (col.startsWith('ρ') || col.startsWith('MW') || col.startsWith('SD')) {
                        inp.readOnly = true; 
                        inp.style.background = 'rgba(255,255,255,0.03)'; 
                        inp.style.fontWeight = '700';
                    }
                    inp.onfocus = () => AppState.selectedCell = `${idx}-${col}`;
                    inp.onchange = (e) => {
                        pushState();
                        originalRow[col] = e.target.value;
                        saveToStorage();
                    };
                    inp.oninput = (e) => {
                        originalRow[col] = e.target.value;
                        updateLiveCalculations(originalRow, col, tr);
                        saveToStorage();
                    };
                    td.appendChild(inp);
                }
                tr.appendChild(td);
            });
        });
        tbody.appendChild(tr);
    });

    // TRIGGER PLOT UPDATE (SAFE)
    try { 
        if (typeof renderAppPlot === 'function') renderAppPlot(); 
    } catch(e) { 
        console.warn("Plotting error:", e); 
    }
}

function openImagePreview(url, t) {
    const v = document.createElement('div');
    v.style = "position:fixed;inset:0;background:rgba(0,0,0,0.9);z-index:99999;display:flex;align-items:center;justify-content:center;";
    v.onclick = () => document.body.removeChild(v);
    const img = new Image(); img.src = url; img.style = "max-width:90%;max-height:90%;border:5px solid white;border-radius:10px;";
    v.appendChild(img); document.body.appendChild(v);
}

// ── AUTO-SAVE MAP DATA ─────────────────────────────────────────────────────
// Saves drawItems GeoJSON directly into messstellen_library so that
// restoreMapDrawings() can find them after a page refresh — no Speichern needed.
function saveMapData() {
    if (!layers.drawItems) return;

    // Stamp markerColor and metadata into feature.properties for each layer
    layers.drawItems.eachLayer(l => {
        if (!l.feature) l.feature = { type: 'Feature', properties: {} };
        
        const c = l.options && l.options.markerColor;
        if (c) {
            l.feature.properties.markerColor = c;
        }

        // Capture Line Color if markerColor is missing
        if (!c && l.options && l.options.color) {
            l.feature.properties.markerColor = l.options.color;
        }

        // Capture Numbering Data
        if (l.options && l.options.isNumbered) {
            l.feature.properties.isNumbered = true;
            l.feature.properties.markerNumber = l.options.markerNumber;
        }
    });

    const pName = document.getElementById('projectName')?.value || 'Unbenanntes Projekt';
    let library = JSON.parse(localStorage.getItem('messstellen_library') || '{}');
    if (!library[pName]) {
        library[pName] = { name: pName, data: AppState.data || [], timestamp: Date.now() };
    }
    library[pName].mapData = layers.drawItems.toGeoJSON();
    library[pName].timestamp = Date.now();
    localStorage.setItem('messstellen_library', JSON.stringify(library));
    localStorage.setItem('current_project_id', pName);
}
// ─────────────────────────────────────────────────────────────────────────────

function showToast(m, duration = 4000) { 
    const t = document.getElementById('toast');
    const msg = document.getElementById('toastMessage');
    if(!t || !msg) return;
    msg.innerText = m;
    t.classList.add('show');
    setTimeout(() => t.classList.remove('show'), duration);
}

function saveToStorage() { 
    try {
        const pName = document.getElementById('projectName').value || 'Unbenanntes Projekt';
        
        if (layers.drawItems) {
            layers.drawItems.eachLayer(l => {
                if (l.options && l.options.markerColor) {
                    if(!l.feature) l.feature = { type: 'Feature', properties: {} };
                    l.feature.properties.markerColor = l.options.markerColor;
                }
            });
        }

        const project = {
            name: pName,
            data: AppState.data,
            mapData: layers.drawItems ? layers.drawItems.toGeoJSON() : null,
            timestamp: Date.now(),
            hiddenMapColors: Array.from(AppState.hiddenMapColors || []),
            newCols: Array.from(AppState.newCols || [])
        };
        
        let library = JSON.parse(localStorage.getItem('messstellen_library') || '{}');
        library[pName] = project;
        localStorage.setItem('messstellen_library', JSON.stringify(library));
        localStorage.setItem('current_project_id', pName);
        return true;
    } catch (e) {
        console.error("Save error:", e);
        showToast("❌ Fehler beim Speichern!");
        return false;
    }
}

function loadFromStorage() { 
    const pName = localStorage.getItem('current_project_id');
    if(!pName) return;

    let library = JSON.parse(localStorage.getItem('messstellen_library') || '{}');
    const project = library[pName];
    
    if(project) {
        AppState.data = project.data || [];
        document.getElementById('projectName').value = project.name;
        
        // Restore hidden states with strict defaults
        if (!project.hiddenMapColors || !project.hiddenMapColors.includes('#00ffff')) {
            AppState.hiddenMapColors = new Set(project.hiddenMapColors || []);
            AppState.hiddenMapColors.add('#00ffff'); // FORCE HIDE 3.2M
        } else {
            AppState.hiddenMapColors = new Set(project.hiddenMapColors);
        }
        
        // Sync map colors to table columns
        if (AppState.hiddenMapColors.has('#00ffff')) AppState.hiddenColumns.add('32');
        else AppState.hiddenColumns.delete('32');
        
        if (AppState.hiddenMapColors.has('#ffff00')) AppState.hiddenColumns.add('16');
        else AppState.hiddenColumns.delete('16');
        
        if (AppState.hiddenMapColors.has('#ff00ff')) AppState.hiddenColumns.add('08');
        else AppState.hiddenColumns.delete('08');
        
        // Initial UI Sync for buttons
        ['08', '16', '32'].forEach(depth => {
            const btn = document.getElementById(`btnToggle${depth}`);
            if (btn) {
                const color = depth === '08' ? '#ff00ff' : (depth === '16' ? '#ffff00' : '#00ffff');
                const isHidden = AppState.hiddenMapColors.has(color);
                btn.style.opacity = isHidden ? '0.4' : '1';
                const icon = btn.querySelector('i');
                if (icon) icon.className = isHidden ? 'fas fa-eye-slash' : 'fas fa-eye';
            }
        });
        
        AppState.newCols = new Set(project.newCols || []);
        if (AppState.newCols.size > 0) {
            let zusatzGroup = TABLE_STRUCTURE.find(g => g.group === 'Zusatz');
            if (!zusatzGroup) {
                zusatzGroup = { group: 'Zusatz', class: 'zusatz', columns: [] };
                TABLE_STRUCTURE.push(zusatzGroup);
            }
            AppState.newCols.forEach(colName => {
                if (!zusatzGroup.columns.includes(colName)) {
                    zusatzGroup.columns.push(colName);
                }
            });
        }
    }
}

function restoreMapDrawings() {
    const pName = document.getElementById('projectName').value;
    let library = JSON.parse(localStorage.getItem('messstellen_library') || '{}');
    const project = library[pName];
    
    if (layers.drawItems && map) {
        layers.drawItems.clearLayers();
        if (project && project.mapData) {
            try {
                L.geoJSON(project.mapData, {
                    pointToLayer: (feature, latlng) => {
                        const color = feature.properties.markerColor || '#3b82f6';
                        const isNum = feature.properties.isNumbered;
                        const num   = feature.properties.markerNumber;

                        if (isNum) {
                            return L.marker(latlng, {
                                draggable: true,
                                markerColor: color,
                                markerNumber: num,
                                isNumbered: true,
                                icon: L.divIcon({
                                    html: `<div style="background:${color} !important; color:black !important; width:30px; height:30px; border-radius:50%; display:flex; align-items:center; justify-content:center; border:3px solid #808080; font-weight:bold; font-size:16px; box-shadow: 0 0 10px rgba(0,0,0,0.5);">${num}</div>`,
                                    iconSize: [30, 30], iconAnchor: [15, 15], className: 'numbered-marker'
                                })
                            });
                        } else {
                            const pinSvg = `<svg width="32" height="42" viewBox="0 0 32 42" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M16 0C7.16344 0 0 7.16344 0 16C0 28 16 42 16 42C16 42 32 28 32 16C32 7.16344 24.8366 0 16 0Z" fill="${color}"/><circle cx="16" cy="16" r="6" fill="white"/></svg>`;
                            return L.marker(latlng, {
                                draggable: true,
                                markerColor: color,
                                icon: L.divIcon({
                                    html: `<div style="width:32px;height:42px;filter:drop-shadow(0 3px 5px rgba(0,0,0,0.5));">${pinSvg}</div>`,
                                    iconSize: [32, 42], iconAnchor: [16, 42], className: 'custom-marker-icon'
                                })
                            });
                        }
                    },
                    style: (feature) => {
                        return { color: feature.properties.markerColor || '#3b82f6', weight: 8 };
                    },
                    onEachFeature: (feature, layer) => {
                        layer.options.markerColor = feature.properties.markerColor;
                        if (feature.properties.isNumbered) {
                            layer.options.isNumbered = true;
                            layer.options.markerNumber = feature.properties.markerNumber;
                        }
                        setupInteractiveLayer(layer);
                        layers.drawItems.addLayer(layer);
                    }
                }).addTo(map);
            } catch (e) { console.error("Error restoring map", e); }
        }
    }
}

function renderProjectList() {
    const grid = document.getElementById('standortGrid');
    const library = JSON.parse(localStorage.getItem('messstellen_library') || '{}');
    if (!grid) return;
    if (Object.keys(library).length === 0) {
        grid.innerHTML = '<p style="color:#888; text-align:center; padding:20px;">Keine Projekte vorhanden</p>';
        return;
    }
    grid.innerHTML = '';
    Object.values(library).sort((a,b) => b.timestamp - a.timestamp).forEach(p => {
        const item = document.createElement('div');
        item.className = 'standort-item';
        item.style = 'background:rgba(255,255,255,0.05); padding:15px; border-radius:12px; margin-bottom:10px; display:flex; justify-content:space-between; align-items:center; border:1px solid rgba(255,255,255,0.1); cursor:pointer;';
        item.innerHTML = `
            <div>
                <div style="font-weight:bold; color:#3b82f6;">${p.name}</div>
                <div style="font-size:11px; color:#888;">${p.data.length} Zeilen • ${new Date(p.timestamp).toLocaleString()}</div>
            </div>
            <button class="btn-modern btn-danger btn-delete-project" style="padding:8px;"><i class="fas fa-trash"></i></button>
        `;
        item.onclick = (e) => {
            if (e.target.closest('.btn-delete-project')) {
                if(confirm(`Projekt "${p.name}" wirklich löschen?`)) {
                    delete library[p.name];
                    localStorage.setItem('messstellen_library', JSON.stringify(library));
                    renderProjectList();
                }
                return;
            }
            loadProject(p.name);
            document.getElementById('standortModal').style.display = 'none';
        };
        grid.appendChild(item);
    });
}

function loadProject(name) {
    let library = JSON.parse(localStorage.getItem('messstellen_library') || '{}');
    const p = library[name];
    if (p) {
        localStorage.setItem('current_project_id', name);
        AppState.data = p.data;
        document.getElementById('projectName').value = p.name;
        renderTable();
        if (map) restoreMapDrawings();
    }
}

function createEmptyRows(n) { for(let i=0; i<n; i++) AppState.data.push({}); }
function openExportModal() {
    if (AppState.data.length === 0) return;
    // Pre-tick checkboxes based on current visible depths
    const chk08 = document.getElementById('exportChk08');
    const chk16 = document.getElementById('exportChk16');
    const chk32 = document.getElementById('exportChk32');
    if (chk08) chk08.checked = !AppState.hiddenColumns.has('08');
    if (chk16) chk16.checked = !AppState.hiddenColumns.has('16');
    if (chk32) chk32.checked = !AppState.hiddenColumns.has('32');

    const modal = document.getElementById('exportModal');
    modal.style.display = 'flex';
}

// Called when user clicks "Exportieren" in the modal
function confirmExportDepths() {
    const chk08 = document.getElementById('exportChk08')?.checked;
    const chk16 = document.getElementById('exportChk16')?.checked;
    const chk32 = document.getElementById('exportChk32')?.checked;
    if (!chk08 && !chk16 && !chk32) {
        showToast("⚠️ Bitte mindestens eine Tiefe auswählen!", 3000);
        return;
    }
    document.getElementById('exportModal').style.display = 'none';
    exportExcel({ export08: chk08, export16: chk16, export32: chk32 });
}

async function exportExcel(opts = null) { 
    if (AppState.data.length === 0) {
        showToast("⚠️ Keine Daten zum Exportieren!", 3000);
        return;
    }
    
    document.getElementById('loadingOverlay').style.display = 'flex';
    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Bodenwiderstand');

        // Styles
        const headerFont = { name: 'Inter', bold: true, size: 13 };
        const dataFont = { name: 'Inter', size: 12 };
        const centerAlignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        const borderStyle = {
            top: { style: 'thin' }, left: { style: 'thin' },
            bottom: { style: 'thin' }, right: { style: 'thin' }
        };

        // Header Rows (Starts at Row 6)
        const row6 = worksheet.getRow(6);
        const row7 = worksheet.getRow(7);

        // Identify ALL unique columns in the data
        const standardCols = ["Kennzeichen", "Alt-Kz.", "Typ", "Örtlichkeit", "Meter [m]", "Datum"];
        const customCols = Array.from(AppState.newCols);
        
        const h1 = ["Kennzeichen", "Alt-Kz.", "Typ", "Örtlichkeit", "Meter\n[m]", "Datum", ...customCols];
        const h2 = Array(h1.length).fill("");
        
        const depthsToExport = [];
        if (opts) {
            if (opts.export08) depthsToExport.push('08');
            if (opts.export16) depthsToExport.push('16');
            if (opts.export32) depthsToExport.push('32');
        } else {
            // Fallback: use hidden columns logic
            ['08','16','32'].forEach(d => { if (!AppState.hiddenColumns.has(d)) depthsToExport.push(d); });
        }
        const depthLabels = { '08': '0.8 m', '16': '1.6 m', '32': '3.2 m' };
        
        depthsToExport.forEach(d => {
            h1.push(depthLabels[d], "", "", "", "", "", "", "", "");
            h2.push("R1\n[Ω]", "R2\n[Ω]", "R3\n[Ω]", "ρ1\n[Ωm]", "ρ2\n[Ωm]", "ρ3\n[Ωm]", "Mittelwert\n[Ωm]", "Standard\nAbweichung", "Anhang");
        });
        
        // Add Anhang Global Header
        const gesAnhangColIdx = h1.length;
        h1.push("Anhang (Gesamt)");
        h2.push("");

        row6.values = h1;
        row7.values = h2;
        row7.height = 45; 

        // Apply Header Styles
        [row6, row7].forEach(row => {
            row.eachCell((cell, colNumber) => {
                cell.font = headerFont;
                cell.alignment = centerAlignment;
                cell.border = borderStyle;
                
                let color = 'FFD9D9D9'; // Default Basis Color
                
                // Group colors
                let currentCol = h1.length - (depthsToExport.length * 9) - 1 + 1; // Adjust for Ges. Anhang
                depthsToExport.forEach(d => {
                    if (colNumber >= currentCol && colNumber < currentCol + 9) {
                        if (d === '08') color = 'FFF2DCDB';
                        if (d === '16') color = 'FFFFFFCC';
                        if (d === '32') color = 'FFD9E1F2';
                    }
                    currentCol += 9;
                });
                
                // Color for Gesamte Anhang
                if (colNumber === gesAnhangColIdx + 1) color = 'FFFFFFFF';

                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
                cell.font = { ...headerFont, color: { argb: 'FF000000' } };
            });
        });

        // Merges for dynamic headers
        for(let i=1; i<=h1.length - (depthsToExport.length * 9) - 1; i++) {
            worksheet.mergeCells(6, i, 7, i);
        }
        worksheet.mergeCells(6, gesAnhangColIdx + 1, 7, gesAnhangColIdx + 1); // Merge Ges. Anhang
        
        let mergeCol = h1.length - (depthsToExport.length * 9);
        depthsToExport.forEach(() => {
            worksheet.mergeCells(6, mergeCol, 6, mergeCol + 8);
            mergeCol += 9;
        });

        // Add Data
        for (let i = 0; i < AppState.data.length; i++) {
            const dataRow = AppState.data[i];
            const excelRow = worksheet.getRow(8 + i);
            excelRow.height = 200; 

            const values = [
                dataRow['Kennzeichen'] || "", dataRow['Alt-Kz.'] || "", dataRow['Typ'] || "", 
                dataRow['Örtlichkeit'] || "", dataRow['Meter [m]'] || "", dataRow['Datum'] || ""
            ];
            
            // Add custom column values
            customCols.forEach(c => values.push(dataRow[c] || ""));

            depthsToExport.forEach(d => {
                const sfx = d === '08' ? '0.8' : (d === '16' ? '1.6' : '3.2');
                const findData = (prefix) => {
                    return dataRow[`${prefix} [Ω]_${sfx}`] || 
                           dataRow[`${prefix} [Ωm]_${sfx}`] || 
                           dataRow[`${prefix}_${sfx}`] || 
                           dataRow[`${prefix} [V]_${sfx}`] || "";
                };
                values.push(
                    findData('R1'), findData('R2'), findData('R3'),
                    findData('ρ1'), findData('ρ2'), findData('ρ3'),
                    findData('MW'), findData('SD'), ""
                );
            });
            
            values.push(""); // Placeholder for Gesamte Anhang
            
            excelRow.values = values;

            excelRow.eachCell((cell, colNumber) => {
                cell.font = dataFont;
                cell.alignment = centerAlignment;
                cell.border = borderStyle;
                if (dataRow._isNew) {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD2B48C' } };
                }
            });

            // Anhang images
            const baseCols = 6 + customCols.length;
            depthsToExport.forEach((d, depthIndex) => {
                const suffix = d === '08' ? '0.8' : (d === '16' ? '1.6' : '3.2');
                const imgData = dataRow[`Anhang_${suffix}`];
                const imgColIndex = baseCols + depthIndex * 9 + 8;
                if (imgData && imgData.startsWith('data:image')) {
                    const imgId = workbook.addImage({ base64: imgData, extension: 'png' });
                    worksheet.addImage(imgId, {
                        tl: { col: imgColIndex, row: 7 + i, nativeColOff: 5, nativeRowOff: 5 },
                        ext: { width: 300, height: 230 }
                    });
                }
            });

            // Anhang Global image
            const gesAnhangImgData = dataRow['Anhang_Global'];
            if (gesAnhangImgData && gesAnhangImgData.startsWith('data:image')) {
                const imgId = workbook.addImage({ base64: gesAnhangImgData, extension: 'png' });
                worksheet.addImage(imgId, {
                    tl: { col: gesAnhangColIdx, row: 7 + i, nativeColOff: 5, nativeRowOff: 5 },
                    ext: { width: 300, height: 230 }
                });
            }
        }

        // Column Widths
        const colWidths = [
            { width: 20 }, { width: 15 }, { width: 12 }, { width: 30 }, { width: 12 }, { width: 15 }
        ];
        
        // Widths for custom columns
        customCols.forEach(() => colWidths.push({ width: 20 }));

        depthsToExport.forEach(() => {
            colWidths.push(
                { width: 10 }, { width: 10 }, { width: 10 }, { width: 15 }, { width: 15 }, { width: 15 }, { width: 20 }, { width: 25 }, { width: 50 }
            );
        });

        // Width for Gesamte Anhang (must match the other Anhang columns)
        colWidths.push({ width: 50 });

        worksheet.columns = colWidths;

        // --- STATISTIK-WEDAL SHEET (CLEAN REPORT LOOK) ---
        const statSheet = workbook.addWorksheet('Statistik-WEDAL');
        
        // Header title for the sheet
        statSheet.mergeCells('A1:L1');
        const titleCell = statSheet.getCell('A1');
        titleCell.value = 'BODENWIDERSTAND - STATISTIK - WEDAL';
        titleCell.font = { name: 'Inter', bold: true, size: 24, color: { argb: 'FF0F172A' } };
        titleCell.alignment = { horizontal: 'center' };

        // --- CHART GENERATION (Numerical Scientific Charts) ---
        const renderExportChart = (type, showSD) => {
            const validData = AppState.data.filter(r => {
                const kmVal = (r['Kennzeichen'] || "").toString().trim();
                if (kmVal === "" || isNaN(parseFloat(kmVal.replace(',', '.')))) return false;
                return Object.keys(r).some(k => (k.includes('MW') || k.includes('R1') || k.includes('ρ')) && r[k] !== "" && r[k] !== null);
            });

            if (validData.length === 0) return null;

            const dataWithKm = validData.map((r, idx) => {
                let kmStr = r['Kennzeichen'].toString().replace(',', '.');
                let km = parseFloat(kmStr);
                return { ...r, _km: km, _idx: idx };
            }).sort((a, b) => a._km - b._km);

            const minKm = dataWithKm[0]._km;
            const maxKm = dataWithKm[dataWithKm.length - 1]._km;
            const kmRange = Math.max(0.001, maxKm - minKm);

            const dynamicWidth = 2400; 
            const canv = document.createElement('canvas');
            canv.width = dynamicWidth; canv.height = 1400; // Significantly taller
            const c = canv.getContext('2d');
            const cp = { top: 120, bottom: 120, left: 140, right: 100 };
            const cw = canv.width - cp.left - cp.right;
            const ch = canv.height - cp.top - cp.bottom;

            c.fillStyle = '#ffffff'; c.fillRect(0, 0, canv.width, canv.height);
            const dCols = { '08': '#ff00ff', '16': '#eab308', '32': '#06b6d4' };
            
            // LEGEND DRAWING
            const lx = cp.left + 50;
            const ly = 95; // Just below title
            c.font = 'bold 22px Inter, sans-serif'; c.textAlign = 'left';
            depthsToExport.forEach((d, i) => {
                const label = depthLabels[d];
                const curX = lx + (i * 350);
                c.fillStyle = dCols[d];
                if (type === 'scatter') {
                    const sz = 12;
                    if (d === '08') { c.beginPath(); c.arc(curX, ly - 8, sz, 0, Math.PI*2); c.fill(); }
                    else if (d === '16') { c.fillRect(curX - sz, ly - 8 - sz, sz*2, sz*2); }
                    else { c.beginPath(); c.moveTo(curX, ly - 8 - sz*1.4); c.lineTo(curX + sz*1.4, ly - 8); c.lineTo(curX, ly - 8 + sz*1.4); c.lineTo(curX - sz*1.4, ly - 8); c.closePath(); c.fill(); }
                } else {
                    c.fillRect(curX - 15, ly - 23, 30, 30);
                }
                c.fillStyle = '#334155';
                c.fillText(label, curX + 35, ly);
            });

            const isLinear = type === 'scatter';
            const px = 50; // Horizontal inner padding to prevent clipping of markers at edges
            
            if (isLinear) {
                // X-Axis Ticks (Linear KM Axis - Aligned with Comparative intervals)
                const tickStep = kmRange > 30 ? 10 : (kmRange > 15 ? 5 : (kmRange > 5 ? 2 : 1));
                c.strokeStyle = '#cbd5e1'; c.lineWidth = 1;
                c.textAlign = 'center'; c.font = 'bold 18px Inter, sans-serif'; c.fillStyle = '#475569';
                
                let lastLabelX = -150;
                for (let val = Math.ceil(minKm / tickStep) * tickStep; val <= maxKm; val += tickStep) {
                    const x = cp.left + px + ((cw - 2*px) * (val - minKm) / kmRange);
                    c.beginPath(); c.moveTo(x, cp.top); c.lineTo(x, cp.top + ch); c.stroke();
                    // Avoid overlapping labels & format numbers cleanly
                    if (x - lastLabelX > 110) {
                        const labelText = Number(val.toFixed(2)).toString() + " km";
                        c.fillText(labelText, x, cp.top + ch + 50);
                        lastLabelX = x;
                    }
                }
            } else {
                // X-Axis Ticks (Discrete Category)
                const colWidth = (cw - 2*px) / dataWithKm.length;
                c.strokeStyle = '#cbd5e1'; c.lineWidth = 1;
                c.textAlign = 'center'; c.font = 'bold 16px Inter, sans-serif'; c.fillStyle = '#475569';
                dataWithKm.forEach((row, i) => {
                    const x = cp.left + px + (i * colWidth) + (colWidth / 2);
                    c.beginPath(); c.moveTo(x, cp.top); c.lineTo(x, cp.top + ch); c.stroke();
                    if (dataWithKm.length < 25 || i % Math.ceil(dataWithKm.length/20) === 0) {
                        c.fillText(row['Kennzeichen'], x, cp.top + ch + 50);
                    }
                });
            }

            // Grid Y
            for(let i=0; i<=10; i++) {
                const y = cp.top + ch - (ch * (i/10));
                c.strokeStyle = '#f1f5f9';
                c.beginPath(); c.moveTo(cp.left, y); c.lineTo(cp.left + cw, y); c.stroke();
            }

            // Axes Box
            c.strokeStyle = '#0f172a'; c.lineWidth = 4;
            c.strokeRect(cp.left, cp.top, cw, ch);

            let mv = 10;
            dataWithKm.forEach(r => {
                depthsToExport.forEach(d => {
                    const sfx = d === '08' ? '0.8' : (d === '16' ? '1.6' : '3.2');
                    const v = parseFloat(r[`MW [Ωm]_${sfx}`]) || 0;
                    const s = showSD ? (parseFloat(r[`SD [Ωm]_${sfx}`]) || 0) : 0;
                    mv = Math.max(mv, v + s);
                });
            });
            mv *= 1.15;

            // Y Labels
            c.textAlign = 'right'; c.font = 'bold 22px Inter, sans-serif'; c.fillStyle = '#1e293b';
            for(let i=0; i<=10; i++) {
                const y = cp.top + ch - (ch * (i/10));
                c.fillText((mv * (i/10)).toFixed(0), cp.left - 30, y + 8);
            }

            const barGroupW = (cw - 2*px) / dataWithKm.length;
            dataWithKm.forEach((row, i) => {
                let x;
                if (isLinear) x = cp.left + px + ((cw - 2*px) * (row._km - minKm) / kmRange);
                else x = cp.left + px + (i * barGroupW) + (barGroupW / 2);

                depthsToExport.forEach((d, di) => {
                    const sfx = d === '08' ? '0.8' : (d === '16' ? '1.6' : '3.2');
                    const val = parseFloat(row[`MW [Ωm]_${sfx}`]) || 0;
                    if (val <= 0) return;
                    
                    const y = cp.top + ch - (ch * (val / mv));
                    const xOff = !isLinear ? (di - (depthsToExport.length-1)/2) * (barGroupW * 0.22) : 0;
                    const bw = !isLinear ? (barGroupW * 0.65) / depthsToExport.length : 18;

                    if (type === 'balken') {
                        c.fillStyle = dCols[d];
                        c.fillRect(x + xOff - bw/2, y, bw, (cp.top + ch) - y);
                        c.strokeStyle = 'rgba(0,0,0,0.4)'; c.lineWidth = 1.5;
                        c.strokeRect(x + xOff - bw/2, y, bw, (cp.top + ch) - y);
                    } else {
                        c.fillStyle = dCols[d]; c.strokeStyle = '#000'; c.lineWidth = 2.5;
                        const sz = 14;
                        if (d === '08') { c.beginPath(); c.arc(x, y, sz, 0, Math.PI*2); c.fill(); c.stroke(); }
                        else if (d === '16') { c.fillRect(x - sz, y - sz, sz*2, sz*2); c.strokeRect(x - sz, y - sz, sz*2, sz*2); }
                        else { c.beginPath(); c.moveTo(x, y - sz*1.4); c.lineTo(x + sz*1.4, y); c.lineTo(x, y + sz*1.4); c.lineTo(x - sz*1.4, y); c.closePath(); c.fill(); c.stroke(); }
                    }

                    if (showSD) {
                        const sd = parseFloat(row[`SD [Ωm]_${sfx}`]) || 0;
                        if (sd > 0) {
                            const sdy = ch * (sd / mv);
                            c.strokeStyle = '#000'; c.lineWidth = 3;
                            c.beginPath(); c.moveTo(x + xOff, y - sdy); c.lineTo(x + xOff, y + sdy); c.stroke();
                            c.beginPath(); c.moveTo(x + xOff - 8, y - sdy); c.lineTo(x + xOff + 8, y - sdy); c.stroke();
                            c.beginPath(); c.moveTo(x + xOff - 8, y + sdy); c.lineTo(x + xOff + 8, y + sdy); c.stroke();
                        }
                    }
                });
            });

            c.textAlign = 'center'; c.font = 'bold 38px Inter, sans-serif'; c.fillStyle = '#0f172a';
            const chartTitle = type === 'balken' ? 'BALKEN - BODENWIDERSTAND - MITTELWERT - WEDAL' : 'PLOT - BODENWIDERSTAND - MITTELWERT - WEDAL';
            c.fillText(`${chartTitle}${showSD ? ' mit Standard Abweichung' : ''}`, canv.width/2, 50);
            
            return { data: canv.toDataURL('image/jpeg', 0.85), w: dynamicWidth };
        };

        const chartPositions = [
            { type: 'balken', sd: false, r: 3, c: 1 },
            { type: 'balken', sd: true, r: 48, c: 1 }, // Spaced out
            { type: 'scatter', sd: false, r: 93, c: 1 },
            { type: 'scatter', sd: true, r: 138, c: 1 }
        ];

        chartPositions.forEach(cfg => {
            const chartObj = renderExportChart(cfg.type, cfg.sd);
            if (!chartObj) return;
            const imgId = workbook.addImage({ base64: chartObj.data, extension: 'jpeg' });
            statSheet.addImage(imgId, {
                tl: { col: cfg.c, row: cfg.r },
                ext: { width: 1800, height: 850 } // Taller in Excel
            });
        });


        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const standort = document.getElementById('projectName').value.trim() || 'WEDAL';
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Bodenwiderstand_${standort}_${new Date().toISOString().split('T')[0]}.xlsx`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.getElementById('loadingOverlay').style.display = 'none';
        showToast("✓ Excel-Statistik erstellt!");
    } catch (err) {
        console.error(err);
        showToast("❌ Fehler beim Excel-Export!", 5000);
    }
    document.getElementById('loadingOverlay').style.display = 'none';
}
function adjustColWidth(col, delta) {
    let current = parseInt(AppState.columnWidths[col]) || 100;
    let next = Math.max(50, current + delta);
    AppState.columnWidths[col] = next + 'px';
    renderTable();
    saveToStorage();
}

// --- SCIENTIFIC DRAWING HELPERS ---
function drawScientificBar(ctx, x, y, w, h, color) {
    ctx.save();
    // 1. Clipping for Hatching
    ctx.beginPath(); ctx.rect(x, y, w, h); ctx.clip();
    
    // 2. Diagonal Hatching
    ctx.strokeStyle = color; ctx.lineWidth = 1; ctx.globalAlpha = 0.4;
    for(let i = -h; i < w; i += 6) {
        ctx.beginPath(); ctx.moveTo(x + i, y); ctx.lineTo(x + i + h, y + h); ctx.stroke();
    }
    ctx.restore();

    // 3. Outline
    ctx.strokeStyle = color; ctx.lineWidth = 2; ctx.globalAlpha = 1;
    ctx.strokeRect(x, y, w, h);

    // 4. Cylinder Top (Ellipse)
    ctx.fillStyle = color; ctx.globalAlpha = 0.2;
    ctx.beginPath(); ctx.ellipse(x + w/2, y, w/2, 4, 0, 0, Math.PI * 2); ctx.fill();
    ctx.globalAlpha = 1; ctx.stroke();
}

// --- INTERACTIVE APP PLOT ---
function renderAppPlot() {
    const container = document.getElementById('appPlotContainer');
    if (!container) return;
    
    const activeDepths = ['0.8', '1.6', '3.2'].filter(d => !AppState.hiddenColumns.has(d === '0.8' ? '08' : (d === '1.6' ? '16' : '32')));
    const plotData = AppState.data.filter(row => {
        return activeDepths.some(d => {
            const mwKey = Object.keys(row).find(k => k.includes('MW') && k.includes(d));
            return mwKey && row[mwKey] && parseFloat(row[mwKey]) > 0;
        });
    });
    
    if (plotData.length === 0 || activeDepths.length === 0) {
        container.innerHTML = `
            <div style="display:flex; flex-direction:column; align-items:center; justify-content:center; height:400px; color:#64748b; font-family:var(--font-display);">
                <i class="fas fa-chart-area" style="font-size:80px; margin-bottom:20px; opacity:0.3;"></i>
                <div style="font-weight:700; font-size:24px; letter-spacing:1px;">KEINE DATEN ZUR ANALYSE</div>
                <div style="font-size:16px; margin-top:5px; opacity:0.7;">Bitte Messwerte (MW) in die Tabelle eingeben.</div>
            </div>`;
        return;
    }

    const minW = Math.max(container.clientWidth, plotData.length * 120);
    const minH = 650;

    container.style.overflow = 'auto';
    container.style.maxHeight = '80vh';
    container.innerHTML = `<canvas id="liveChartCanvas" style="width:${minW}px; height:${minH}px; cursor:crosshair;"></canvas>`;
    
    const canvas = document.getElementById('liveChartCanvas');
    const ctx = canvas.getContext('2d');
    
    const dpr = window.devicePixelRatio || 1;
    canvas.width = minW * dpr;
    canvas.height = minH * dpr;
    ctx.scale(dpr, dpr);

    const w = minW, h = minH;
    
    // Standard Dashboard Sizing
    const fsTitle = 24;
    const fsLabel = 14;
    const fsTick = 12;
    
    const pad = { 
        top: 70, 
        bottom: 90, 
        left: 100, 
        right: 40 
    }; 
    const chartW = w - pad.left - pad.right;
    const chartH = h - pad.top - pad.bottom;

    let maxV = 10;
    plotData.forEach(row => {
        activeDepths.forEach(d => {
            const mwKey = Object.keys(row).find(k => k.includes('MW') && k.includes(d));
            const val = mwKey ? parseFloat(row[mwKey]) || 0 : 0;
            const sdKey = Object.keys(row).find(k => k.includes('SD') && k.includes(d));
            const sdVal = (AppState.showStd && sdKey) ? parseFloat(row[sdKey]) || 0 : 0;
            maxV = Math.max(maxV, val + sdVal);
        });
    });
    maxV *= 1.15;

    // Background
    ctx.fillStyle = 'rgba(15, 23, 42, 0.4)';
    ctx.fillRect(pad.left, pad.top, chartW, chartH);

    // Ticks & Axis (Neon Style)
    ctx.strokeStyle = '#64748b'; ctx.lineWidth = 3;
    ctx.beginPath(); 
    ctx.moveTo(pad.left, pad.top); ctx.lineTo(pad.left, pad.top + chartH); 
    ctx.lineTo(pad.left + chartW, pad.top + chartH); ctx.stroke();

    // LEGEND AT TOP (Horizontal Layout)
    const colors = { '0.8': '#ff00ff', '1.6': '#ffff00', '3.2': '#00ffff' };
    const legW = 120;
    const legStartX = pad.left + (chartW - (activeDepths.length * legW)) / 2;
    activeDepths.forEach((d, di) => {
        const lx = legStartX + (di * legW);
        const ly = 25;
        const mSize = 14;
        ctx.strokeStyle = colors[d]; ctx.lineWidth = 2; ctx.beginPath();
        
        // Marker Logic: Balken = Squares, Scatter = Unique Markers
        if (AppState.chartType === 'balken') {
            ctx.fillStyle = colors[d] + '44';
            ctx.fillRect(lx, ly, mSize, mSize);
            ctx.rect(lx, ly, mSize, mSize);
        } else {
            if (d === '0.8') ctx.arc(lx + mSize/2, ly + mSize/2, mSize/2, 0, Math.PI*2);
            else if (d === '1.6') ctx.rect(lx, ly, mSize, mSize);
            else {
                ctx.moveTo(lx + mSize/2, ly); ctx.lineTo(lx + mSize, ly + mSize/2); ctx.lineTo(lx + mSize/2, ly + mSize); ctx.lineTo(lx, ly + mSize/2); ctx.closePath();
            }
        }
        ctx.stroke();
        
        ctx.fillStyle = '#fff'; ctx.font = `bold 14px 'Outfit', sans-serif`; ctx.textAlign = 'left';
        ctx.fillText(`${d}m`, lx + mSize * 1.5, ly + mSize * 0.85);
    });

    // Axis Titles
    ctx.fillStyle = '#fff'; ctx.font = `900 ${fsTitle}px 'Outfit', sans-serif`; ctx.textAlign = 'center';
    ctx.fillText('KILOMETRIERUNG [km]', pad.left + chartW/2, h - 15);
    
    ctx.save(); ctx.translate(40, pad.top + chartH/2); ctx.rotate(-Math.PI/2);
    ctx.fillText('ρ [Ωm]', 0, 0); ctx.restore();

    // Y-Axis Major & Minor Ticks
    ctx.textAlign = 'right'; ctx.font = `bold ${fsTick}px 'Inter', sans-serif`;
    for(let i=0; i<=10; i++) {
        const val = (maxV * (i/10));
        const y = pad.top + chartH - (chartH * (i/10));
        const isMajor = i % 2 === 0;
        
        if(isMajor) {
            ctx.strokeStyle = 'rgba(255,255,255,0.15)'; ctx.lineWidth = 1;
            ctx.beginPath(); ctx.moveTo(pad.left, y); ctx.lineTo(pad.left + chartW, y); ctx.stroke();
            ctx.strokeStyle = '#fff'; ctx.lineWidth = 3;
            ctx.beginPath(); ctx.moveTo(pad.left, y); ctx.lineTo(pad.left - 10, y); ctx.stroke();
            ctx.fillStyle = '#fff';
            ctx.fillText(val.toFixed(1), pad.left - 15, y + 5);
        } else {
            ctx.strokeStyle = '#94a3b8'; ctx.lineWidth = 1;
            ctx.beginPath(); ctx.moveTo(pad.left, y); ctx.lineTo(pad.left - 6, y); ctx.stroke();
        }
    }

    const bG = chartW / Math.max(1, plotData.length);
    const barAreas = [];

    // X-Axis Major & Minor Ticks
    ctx.textAlign = 'center'; ctx.font = `bold ${fsTick}px 'Inter', sans-serif`;
    plotData.forEach((row, i) => {
        const x = pad.left + (i * bG) + bG/2;
        ctx.strokeStyle = '#fff'; ctx.lineWidth = 3;
        ctx.beginPath(); ctx.moveTo(x, pad.top + chartH); ctx.lineTo(x, pad.top + chartH + 10); ctx.stroke();
        
        ctx.fillStyle = '#fff';
        let rawKz = row['Kennzeichen'] || (i+1).toString();
        let displayKz = rawKz;
        if (!isNaN(parseFloat(rawKz)) && isFinite(rawKz)) {
            displayKz = (parseFloat(rawKz) / 1000).toFixed(3);
        }
        ctx.fillText(displayKz, x, pad.top + chartH + 30);

        if (i < plotData.length - 1) {
            const mx = pad.left + (i * bG) + bG;
            ctx.strokeStyle = '#94a3b8'; ctx.lineWidth = 1;
            ctx.beginPath(); ctx.moveTo(mx, pad.top + chartH); ctx.lineTo(mx, pad.top + chartH + 6); ctx.stroke();
        }
    });

    if (AppState.chartType === 'balken') {
        const groupW = bG * 0.8;
        const barW = groupW / (activeDepths.length || 1);
        plotData.forEach((row, i) => {
            const startX = pad.left + (i * bG) + (bG - groupW) / 2;
            activeDepths.forEach((d, di) => {
                const mwKey = Object.keys(row).find(k => k.includes('MW') && k.includes(d));
                const val = mwKey ? parseFloat(row[mwKey]) || 0 : 0;
                if (val <= 0) return;

                const bx = startX + (di * barW);
                const by = pad.top + chartH - (val / maxV) * chartH;
                const bw = Math.max(4, barW - 6);
                const bh = (val / maxV) * chartH;

                drawScientificBar(ctx, bx, by, bw, bh, colors[d]);
                
                if (AppState.showStd) {
                    const sdKey = Object.keys(row).find(k => k.includes('SD') && k.includes(d));
                    const sdVal = sdKey ? parseFloat(row[sdKey]) || 0 : 0;
                    if (sdVal > 0) {
                        const sdH = (sdVal / maxV) * chartH;
                        ctx.strokeStyle = '#fff'; ctx.lineWidth = 3;
                        ctx.beginPath();
                        ctx.moveTo(bx + bw/2, by - sdH); ctx.lineTo(bx + bw/2, by + sdH);
                        ctx.moveTo(bx + bw/2 - 10, by - sdH); ctx.lineTo(bx + bw/2 + 10, by - sdH);
                        ctx.moveTo(bx + bw/2 - 10, by + sdH); ctx.lineTo(bx + bw/2 + 10, by + sdH);
                        ctx.stroke();
                    }
                }
                barAreas.push({ x: bx, y: by - 10, w: bw, h: bh + 20, idx: AppState.data.indexOf(row) });
            });
        });
    } else {
        activeDepths.forEach(d => {
            const pts = [];
            plotData.forEach((row, i) => {
                const mwKey = Object.keys(row).find(k => k.includes('MW') && k.includes(d));
                const val = mwKey ? parseFloat(row[mwKey]) || 0 : 0;
                if (val <= 0) return;
                pts.push({
                    x: pad.left + (i * bG) + bG/2,
                    y: pad.top + chartH - (val / maxV) * chartH,
                    val: val,
                    row: row
                });
            });

            if (pts.length === 0) return;

            ctx.strokeStyle = colors[d] + '88'; ctx.lineWidth = 4; ctx.lineJoin = 'round';
            ctx.beginPath();
            pts.forEach((p, pi) => { if(pi === 0) ctx.moveTo(p.x, p.y); else ctx.lineTo(p.x, p.y); });
            ctx.stroke();

            pts.forEach(p => {
                if (AppState.showStd) {
                    const sdKey = Object.keys(p.row).find(k => k.includes('SD') && k.includes(d));
                    const sdVal = sdKey ? parseFloat(p.row[sdKey]) || 0 : 0;
                    if (sdVal > 0) {
                        const sdH = (sdVal / maxV) * chartH;
                        ctx.strokeStyle = '#fff'; ctx.lineWidth = 3;
                        ctx.beginPath();
                        ctx.moveTo(p.x, p.y - sdH); ctx.lineTo(p.x, p.y + sdH);
                        ctx.moveTo(p.x - 10, p.y - sdH); ctx.lineTo(p.x + 10, p.y - sdH);
                        ctx.moveTo(p.x - 10, p.y + sdH); ctx.lineTo(p.x + 10, p.y + sdH);
                        ctx.stroke();
                    }
                }

                ctx.strokeStyle = colors[d]; ctx.lineWidth = 5; ctx.beginPath();
                if (d === '0.8') ctx.arc(p.x, p.y, 12, 0, Math.PI*2);
                else if (d === '1.6') ctx.rect(p.x - 12, p.y - 12, 24, 24);
                else {
                    ctx.moveTo(p.x, p.y - 15); ctx.lineTo(p.x + 15, p.y); ctx.lineTo(p.x, p.y + 15); ctx.lineTo(p.x - 15, p.y); ctx.closePath();
                }
                ctx.stroke();
                barAreas.push({ x: p.x - 20, y: p.y - 20, w: 40, h: 40, idx: AppState.data.indexOf(p.row) });
            });
        });
    }


    canvas.onclick = (e) => {
        const rect = canvas.getBoundingClientRect();
        const mx = e.clientX - rect.left, my = e.clientY - rect.top;
        const hit = barAreas.find(b => mx >= b.x && mx <= b.x + b.w && my >= b.y && my <= b.y + b.h);
        if (hit) {
            switchTab('table');
            setTimeout(() => {
                const trs = document.querySelectorAll('#tableBody tr');
                if(trs[hit.idx]) {
                    trs[hit.idx].scrollIntoView({ behavior: 'smooth', block: 'center' });
                    trs[hit.idx].style.boxShadow = 'inset 0 0 20px rgba(0, 242, 255, 0.5)';
                    setTimeout(() => trs[hit.idx].style.boxShadow = '', 2000);
                }
            }, 100);
        }
    };
}



function updateChartUI() {
    const bBalken = document.getElementById('btnChartBalken');
    const bScatter = document.getElementById('btnChartScatter');
    const bStdOn = document.getElementById('btnStdOn');
    const bStdOff = document.getElementById('btnStdOff');

    if (bBalken) bBalken.classList.toggle('active', AppState.chartType === 'balken');
    if (bScatter) bScatter.classList.toggle('active', AppState.chartType === 'scatter');
    if (bStdOn) bStdOn.classList.toggle('active', AppState.showStd);
    if (bStdOff) bStdOff.classList.toggle('active', !AppState.showStd);
}

// ─────────────────────────────────────────────────────────────────────────────
// COLUMN MANAGER LOGIC
// ─────────────────────────────────────────────────────────────────────────────

function openColManager() {
    const modal = document.getElementById('colManagerModal');
    const grid = document.getElementById('colManagerGrid');
    if (!modal || !grid) return;

    grid.innerHTML = '';
    TABLE_STRUCTURE.forEach(g => {
        const isHidden = AppState.hiddenColumns.has(g.class);
        const item = document.createElement('div');
        item.className = 'col-toggle-item';
        item.style.display = 'flex';
        item.style.alignItems = 'center';
        item.style.justifyContent = 'space-between';
        item.style.padding = '12px';
        item.style.background = 'rgba(255,255,255,0.03)';
        item.style.border = `1px solid ${isHidden ? 'var(--border-main)' : (DEPTH_COLORS[g.class]?.border || 'var(--accent-primary)')}`;
        item.style.borderRadius = '4px';
        item.style.cursor = 'pointer';
        item.style.transition = '0.3s';
        
        // Checkbox UI für Klarheit
        const iconClass = isHidden ? 'far fa-square' : 'fas fa-check-square';
        const statusColor = isHidden ? '#555' : (DEPTH_COLORS[g.class]?.border || '#fff');

        item.innerHTML = `
            <div style="display:flex; align-items:center; gap:10px;">
                <i class="${iconClass}" style="color: ${statusColor}; font-size: 16px;"></i>
                <span style="font-weight:600; font-size:12px; color: ${isHidden ? '#888' : '#fff'}">${g.group}</span>
            </div>
            <div class="custom-toggle ${isHidden ? '' : 'active'}" style="width:30px; height:15px; border:1px solid #555; border-radius:10px; position:relative;">
                <div style="width:10px; height:10px; background:${isHidden ? '#555' : 'var(--accent-primary)'}; border-radius:50%; position:absolute; top:2px; left:${isHidden ? '3px' : '15px'}; transition:0.3s;"></div>
            </div>
        `;
        
        item.onclick = () => toggleColumnGroup(g.class, g.group);
        grid.appendChild(item);
    });

    modal.style.display = 'block';
}

function toggleColumnGroup(cls, groupName) {
    if (AppState.hiddenColumns.has(cls)) {
        AppState.hiddenColumns.delete(cls);
    } else {
        AppState.hiddenColumns.add(cls);
    }
    
    renderTable();
    saveToStorage();
    openColManager(); 
}
function startLiveTracking() {
    if (!navigator.geolocation) return;
    if (AppState.watchId) navigator.geolocation.clearWatch(AppState.watchId);
    
    AppState.watchId = navigator.geolocation.watchPosition(
        (pos) => {
            const { latitude, longitude, accuracy } = pos.coords;
            updateUserMarker(latitude, longitude, accuracy);
        },
        (err) => {
            console.warn("GPS Error:", err);
            const statusText = document.getElementById('gpsStatusText');
            if(statusText) statusText.innerText = "❌ GPS gesperrt / Fehler";
        },
        { enableHighAccuracy: true, maximumAge: 0, timeout: 15000 }
    );
}

function updateUserMarker(lat, lng, accuracy) {
    if (!map) return;
    const latlng = [lat, lng];

    if (!layers.userLocation) {
        layers.userLocation = L.layerGroup().addTo(map);
    }
    layers.userLocation.clearLayers();

    // Blue Accuracy Circle
    L.circle(latlng, { radius: accuracy, color: '#3b82f6', fillOpacity: 0.1, weight: 1 }).addTo(layers.userLocation);
    
    // Core User Dot
    AppState.userMarker = L.circleMarker(latlng, {
        radius: 8,
        fillColor: '#3b82f6',
        color: '#fff',
        weight: 2,
        fillOpacity: 1
    }).addTo(layers.userLocation);

    // Update Dashboard UI
    const statusText = document.getElementById('gpsStatusText');
    const dot = document.getElementById('gpsAccuracyDot');
    if (statusText) statusText.innerText = `GPS Aktiv (${accuracy.toFixed(1)}m)`;
    if (dot) {
        dot.style.background = accuracy < 10 ? '#39ff14' : (accuracy < 30 ? '#ff8c00' : '#ef4444');
    }

    // Auto-Follow Logic: Center and Zoom In for precision
    if (AppState.liveFollow) {
        if (!AppState.firstLocationFound) {
            map.setView(latlng, 19, { animate: true });
            AppState.firstLocationFound = true;
        } else {
            map.panTo(latlng, { animate: true });
        }
    }
}
