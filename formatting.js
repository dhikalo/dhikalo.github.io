// === FORMATTING FUNCTIONS ===
let selectedCells = new Set();
let cellFormatting = {};

// Track selected cell
document.addEventListener('DOMContentLoaded', () => {
    setTimeout(() => {
        const tableBody = document.getElementById('tableBody');
        if (tableBody) {
            tableBody.addEventListener('click', (e) => {
                if (e.target.classList.contains('cell-input')) {
                    const row = e.target.dataset.row;
                    const col = e.target.dataset.col;
                    const key = ${row}-;
                    selectedCells.clear();
                    selectedCells.add(key);
                }
            });
        }
    }, 1000);
});

function toggleFormat(type) {
    if (selectedCells.size === 0) {
        showToast('Bitte wählen Sie eine Zelle!', 'warning');
        return;
    }
    
    selectedCells.forEach(key => {
        if (!cellFormatting[key]) cellFormatting[key] = {};
        cellFormatting[key][type] = !cellFormatting[key][type];
        
        const [row, col] = key.split('-');
        const input = document.querySelector(input[data-row=""][data-col=""]);
        if (input) {
            applyFormattingToCell(input, cellFormatting[key]);
        }
    });
    
    const btn = document.getElementById(tn);
    if (btn) btn.classList.toggle('active');
}

function setAlignment(align) {
    if (selectedCells.size === 0) {
        showToast('Bitte wählen Sie eine Zelle!', 'warning');
        return;
    }
    
    selectedCells.forEach(key => {
        if (!cellFormatting[key]) cellFormatting[key] = {};
        cellFormatting[key].align = align;
        
        const [row, col] = key.split('-');
        const input = document.querySelector(input[data-row=""][data-col=""]);
        if (input) {
            applyFormattingToCell(input, cellFormatting[key]);
        }
    });
}

function applyFormattingToCell(input, format) {
    if (format.bold) input.style.fontWeight = 'bold';
    else input.style.fontWeight = '';
    
    if (format.italic) input.style.fontStyle = 'italic';
    else input.style.fontStyle = '';
    
    if (format.underline) input.style.textDecoration = 'underline';
    else input.style.textDecoration = '';
    
    if (format.bgColor) input.style.backgroundColor = format.bgColor;
    if (format.textColor) input.style.color = format.textColor;
    if (format.align) input.style.textAlign = format.align;
}

function clearFormatting() {
    if (selectedCells.size === 0) {
        showToast('Bitte wählen Sie eine Zelle!', 'warning');
        return;
    }
    
    selectedCells.forEach(key => {
        delete cellFormatting[key];
        
        const [row, col] = key.split('-');
        const input = document.querySelector(input[data-row=""][data-col=""]);
        if (input) {
            input.style.fontWeight = '';
            input.style.fontStyle = '';
            input.style.textDecoration = '';
            input.style.backgroundColor = '';
            input.style.color = '';
            input.style.textAlign = '';
        }
    });
    
    showToast('Formatierung gelöscht', 'success');
}

// === COLOR PICKERS ===
let bgColorPicker, textColorPicker;

function initColorPickers() {
    console.log('Initializing color pickers...');
    
    const btnBgColor = document.getElementById('btnBgColor');
    const btnTextColor = document.getElementById('btnTextColor');
    
    if (btnBgColor && typeof Pickr !== 'undefined') {
        bgColorPicker = Pickr.create({
            el: '#btnBgColor',
            theme: 'nano',
            default: '#FFFFFF',
            swatches: [
                '#FEF9C3', '#FEE2E2', '#DBEAFE', '#D1FAE5', 
                '#E0E7FF', '#FCE7F3', '#F3F4F6', '#FFFFFF'
            ],
            components: {
                preview: true,
                opacity: true,
                hue: true,
                interaction: {
                    hex: true,
                    input: true,
                    save: true
                }
            }
        });
        
        bgColorPicker.on('save', (color) => {
            if (!color) return;
            const hex = color.toHEXA().toString();
            
            selectedCells.forEach(key => {
                if (!cellFormatting[key]) cellFormatting[key] = {};
                cellFormatting[key].bgColor = hex;
                
                const [row, col] = key.split('-');
                const input = document.querySelector(input[data-row=""][data-col=""]);
                if (input) {
                    input.style.backgroundColor = hex;
                }
            });
            
            const preview = document.getElementById('bgColorPreview');
            if (preview) preview.style.background = hex;
            
            bgColorPicker.hide();
        });
    }
    
    if (btnTextColor && typeof Pickr !== 'undefined') {
        textColorPicker = Pickr.create({
            el: '#btnTextColor',
            theme: 'nano',
            default: '#111827',
            swatches: [
                '#111827', '#EF4444', '#2563EB', '#10B981', 
                '#F59E0B', '#8B5CF6', '#EC4899', '#06B6D4'
            ],
            components: {
                preview: true,
                opacity: true,
                hue: true,
                interaction: {
                    hex: true,
                    input: true,
                    save: true
                }
            }
        });
        
        textColorPicker.on('save', (color) => {
            if (!color) return;
            const hex = color.toHEXA().toString();
            
            selectedCells.forEach(key => {
                if (!cellFormatting[key]) cellFormatting[key] = {};
                cellFormatting[key].textColor = hex;
                
                const [row, col] = key.split('-');
                const input = document.querySelector(input[data-row=""][data-col=""]);
                if (input) {
                    input.style.color = hex;
                }
            });
            
            const preview = document.getElementById('textColorPreview');
            if (preview) preview.style.background = hex;
            
            textColorPicker.hide();
        });
    }
    
    console.log('Color pickers initialized');
}

// === COLUMN FILTER ===
let hiddenColumns = new Set();

function initColumnFilter() {
    const filterInput = document.getElementById('columnFilter');
    const btnResetFilter = document.getElementById('btnResetFilter');
    
    if (filterInput) {
        filterInput.addEventListener('input', (e) => {
            const query = e.target.value.toLowerCase().trim();
            filterColumns(query);
        });
    }
    
    if (btnResetFilter) {
        btnResetFilter.addEventListener('click', () => {
            if (filterInput) filterInput.value = '';
            hiddenColumns.clear();
            renderTable();
            updateFilterInfo();
        });
    }
}

function filterColumns(query) {
    if (!query) {
        hiddenColumns.clear();
        renderTable();
        updateFilterInfo();
        return;
    }
    
    hiddenColumns.clear();
    ALL_COLUMNS.forEach((col, idx) => {
        if (!col.toLowerCase().includes(query)) {
            hiddenColumns.add(idx);
        }
    });
    
    renderTable();
    updateFilterInfo();
}

function updateFilterInfo() {
    const visibleCount = ALL_COLUMNS.length - hiddenColumns.size;
    const info = document.getElementById('visibleColumns');
    if (info) {
        if (hiddenColumns.size === 0) {
            info.textContent = 'Alle Spalten sichtbar';
        } else {
            info.textContent = ${visibleCount} von  Spalten sichtbar;
        }
    }
}

