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

const AppState = {
    hiddenColumns: new Set(['32', 'potential', 'spannung', 'strom', 'widerstand', 'audit', 'Kennzeichen', 'Alt-Kz.', 'Datum']),
    newCols: new Set()
};

async function test() {
    const h1 = [];
    const h2 = [];
    const colDefinitions = [];
    const depthsToExport = ['08', '16'];
    const customCols = [];

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

    console.log('H1 Headers:', h1);
    console.log('H2 Headers:', h2);
    console.log('ColDefinitions Keys:', colDefinitions.map(c => c.key));
}

test();
