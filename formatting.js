/* ============================================
   FORMATTING MODULE - Cell Styling & Colors
   ============================================ */

let selectedCells = new Set();
let cellFormatting = {};

document.addEventListener('DOMContentLoaded', () => {
    setTimeout(() => {
        const tableBody = document.getElementById('tableBody');
        if (tableBody) {
            tableBody.addEventListener('click', (e) => {
                if (e.target.classList.contains('cell-input')) {
                    const row = e.target.dataset.row;
                    const col = e.target.dataset.col;
                    const key = `${row}-${col}`;
                    selectedCells.clear();
                    selectedCells.add(key);
                }
            });
        }

        // Formatting buttons
        const btnBold = document.getElementById('btnBold');
        const btnItalic = document.getElementById('btnItalic');
        const btnUnderline = document.getElementById('btnUnderline');
        const btnBgColor = document.getElementById('btnBgColor');

        if (btnBold) btnBold.addEventListener('click', () => toggleFormat('bold'));
        if (btnItalic) btnItalic.addEventListener('click', () => toggleFormat('italic'));
        if (btnUnderline) btnUnderline.addEventListener('click', () => toggleFormat('underline'));
        if (btnBgColor) {
            btnBgColor.addEventListener('click', () => {
                const color = prompt('Hintergrundfarbe (Hex, z.B. #FEF9C3):');
                if (color) applyBgColor(color);
            });
        }
    }, 500);
});

function toggleFormat(type) {
    if (selectedCells.size === 0) {
        if (typeof showToast === 'function') showToast('Bitte zuerst eine Zelle auswählen');
        return;
    }

    selectedCells.forEach(key => {
        if (!cellFormatting[key]) cellFormatting[key] = {};
        cellFormatting[key][type] = !cellFormatting[key][type];

        const [row, col] = key.split('-');
        const input = document.querySelector(`input[data-row="${row}"][data-col="${col}"]`);
        if (input) {
            applyFormattingToCell(input, cellFormatting[key]);
        }
    });
}

function applyBgColor(color) {
    if (selectedCells.size === 0) return;

    selectedCells.forEach(key => {
        if (!cellFormatting[key]) cellFormatting[key] = {};
        cellFormatting[key].bgColor = color;

        const [row, col] = key.split('-');
        const input = document.querySelector(`input[data-row="${row}"][data-col="${col}"]`);
        if (input) {
            input.style.backgroundColor = color;
        }
    });
}

function applyFormattingToCell(input, format) {
    if (!input || !format) return;

    input.style.fontWeight = format.bold ? 'bold' : '';
    input.style.fontStyle = format.italic ? 'italic' : '';
    input.style.textDecoration = format.underline ? 'underline' : '';
    if (format.bgColor) input.style.backgroundColor = format.bgColor;
    if (format.textColor) input.style.color = format.textColor;
    if (format.align) input.style.textAlign = format.align;
}

function clearFormatting() {
    if (selectedCells.size === 0) return;

    selectedCells.forEach(key => {
        delete cellFormatting[key];

        const [row, col] = key.split('-');
        const input = document.querySelector(`input[data-row="${row}"][data-col="${col}"]`);
        if (input) {
            input.style.fontWeight = '';
            input.style.fontStyle = '';
            input.style.textDecoration = '';
            input.style.backgroundColor = '';
            input.style.color = '';
            input.style.textAlign = '';
        }
    });

    if (typeof showToast === 'function') showToast('Formatierung gelöscht');
}
