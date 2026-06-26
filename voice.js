/* ============================================
   SPRACHEINGABE — Voice Input Module
   Messstellen Manager Pro

   Allows field workers to speak measurement
   values directly into table cells.
   No internet needed — uses device speech API.
   ============================================ */

(function () {
    'use strict';

    // ── Support check ──────────────────────────────────────────────────────────
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRecognition) {
        console.warn('[Voice] Web Speech API not supported — Spracheingabe deaktiviert');
        // Still inject the toolbar button but show a "not supported" toast on click
        injectUI(false);
        return;
    }

    // ── State ──────────────────────────────────────────────────────────────────
    let recognition = null;
    let isListening = false;
    let activeInput = null;     // currently focused cell <input>
    let activeRow = null;       // row index of activeInput
    let rowDictationMode = false;

    // ── Styles ────────────────────────────────────────────────────────────────
    const styleEl = document.createElement('style');
    styleEl.textContent = `
        /* ── Voice FAB button ── */
        #voiceMicFab {
            position: fixed;
            bottom: 80px;
            right: 18px;
            width: 60px;
            height: 60px;
            border-radius: 50%;
            background: #1e293b;
            border: 2px solid #334155;
            color: #64748b;
            font-size: 22px;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 9999;
            box-shadow: 0 4px 16px rgba(0,0,0,0.5);
            transition: background 0.2s, color 0.2s, border-color 0.2s, transform 0.1s;
            touch-action: manipulation;
            -webkit-tap-highlight-color: transparent;
        }
        #voiceMicFab.active {
            background: #0f172a;
            border-color: #3b82f6;
            color: #3b82f6;
        }
        #voiceMicFab.listening {
            background: #ef4444;
            border-color: #fca5a5;
            color: white;
            animation: voicePulse 0.9s ease-in-out infinite;
        }
        #voiceMicFab:disabled {
            opacity: 0.35;
            cursor: not-allowed;
        }
        @keyframes voicePulse {
            0%, 100% { box-shadow: 0 0 0 0 rgba(239,68,68,0.6); }
            50%       { box-shadow: 0 0 0 14px rgba(239,68,68,0); }
        }

        /* ── Voice preview bar (shows interim transcript) ── */
        #voicePreviewBar {
            position: fixed;
            bottom: 150px;
            right: 18px;
            max-width: 260px;
            background: #1e293b;
            border: 1px solid #3b82f6;
            border-radius: 10px;
            padding: 8px 14px;
            color: #93c5fd;
            font-size: 14px;
            font-family: 'Inter', sans-serif;
            z-index: 9998;
            display: none;
            word-break: break-word;
            box-shadow: 0 4px 16px rgba(0,0,0,0.5);
        }
        #voicePreviewBar.show { display: block; }
        #voicePreviewBar .voice-label {
            font-size: 10px;
            color: #64748b;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-bottom: 3px;
        }

        /* ── Row dictation modal ── */
        #voiceDictationModal {
            position: fixed;
            inset: 0;
            background: rgba(0,0,0,0.7);
            z-index: 10000;
            display: none;
            align-items: flex-end;
            justify-content: center;
        }
        #voiceDictationModal.show { display: flex; }
        #voiceDictationPanel {
            background: #0f172a;
            border-top: 2px solid #3b82f6;
            border-radius: 16px 16px 0 0;
            padding: 24px 20px 32px;
            width: 100%;
            max-width: 500px;
        }
        #voiceDictationPanel h3 {
            color: #e2e8f0;
            font-size: 16px;
            margin: 0 0 6px;
        }
        #voiceDictationPanel p {
            color: #64748b;
            font-size: 12px;
            margin: 0 0 16px;
        }
        #dictationMicBtn {
            width: 80px;
            height: 80px;
            border-radius: 50%;
            background: #1e293b;
            border: 2px solid #334155;
            color: #64748b;
            font-size: 28px;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 0 auto 12px;
            transition: all 0.2s;
        }
        #dictationMicBtn.listening {
            background: #ef4444;
            border-color: #fca5a5;
            color: white;
            animation: voicePulse 0.9s ease-in-out infinite;
        }
        #dictationTranscript {
            min-height: 40px;
            background: #1e293b;
            border-radius: 8px;
            padding: 10px 14px;
            color: #93c5fd;
            font-size: 14px;
            margin-bottom: 12px;
            text-align: center;
        }
        #dictationResult {
            font-size: 12px;
            color: #10b981;
            text-align: center;
            min-height: 18px;
            margin-bottom: 12px;
        }
        #btnCloseDictation {
            display: block;
            width: 100%;
            padding: 10px;
            background: #1e293b;
            border: 1px solid #334155;
            border-radius: 8px;
            color: #94a3b8;
            font-size: 14px;
            cursor: pointer;
        }
    `;
    document.head.appendChild(styleEl);

    // ── DOM ────────────────────────────────────────────────────────────────────
    function injectUI(supported) {
        // Floating mic FAB
        const fab = document.createElement('button');
        fab.id = 'voiceMicFab';
        fab.title = 'Zelle auswählen für Spracheingabe';
        fab.innerHTML = '<i class="fas fa-microphone"></i>';
        if (!supported) fab.disabled = true;
        document.body.appendChild(fab);

        // Interim preview bar
        const preview = document.createElement('div');
        preview.id = 'voicePreviewBar';
        preview.innerHTML = '<div class="voice-label">Erkannt</div><div id="voicePreviewText">...</div>';
        document.body.appendChild(preview);

        // Row dictation modal
        const modal = document.createElement('div');
        modal.id = 'voiceDictationModal';
        modal.innerHTML = `
            <div id="voiceDictationPanel">
                <h3><i class="fas fa-microphone-alt"></i> Spracheingabe</h3>
                <p>Zelle auswählen, Mikrofon drücken, Zahl sprechen</p>
                <div id="dictationTranscript">Tippen um zu starten...</div>
                <div id="dictationResult"></div>
                <button id="dictationMicBtn"><i class="fas fa-microphone"></i></button>
                <button id="btnCloseDictation"><i class="fas fa-times"></i> Schließen</button>
            </div>
        `;
        document.body.appendChild(modal);

        if (supported) initEvents(fab, modal);
        else {
            fab.addEventListener('click', () => {
                if (typeof showToast === 'function') showToast('Spracheingabe wird von diesem Browser nicht unterstützt. Bitte Chrome verwenden.');
            });
        }
    }

    // ── Number parsing ─────────────────────────────────────────────────────────
    const ONES = {
        'null': 0, 'ein': 1, 'eins': 1, 'eine': 1,
        'zwei': 2, 'zwo': 2, 'drei': 3, 'vier': 4,
        'fünf': 5, 'sechs': 6, 'sieben': 7, 'acht': 8, 'neun': 9,
        'zehn': 10, 'elf': 11, 'zwölf': 12,
        'dreizehn': 13, 'vierzehn': 14, 'fünfzehn': 15,
        'sechzehn': 16, 'siebzehn': 17, 'achtzehn': 18, 'neunzehn': 19
    };
    const TENS = {
        'zwanzig': 20, 'dreißig': 30, 'vierzig': 40,
        'fünfzig': 50, 'sechzig': 60, 'siebzig': 70,
        'achtzig': 80, 'neunzig': 90
    };

    function parseCompound(word) {
        // "fünfundvierzig" → 45, "dreiundzwanzig" → 23
        const m = word.match(/^(.+?)und(.+)$/);
        if (m) {
            const ones = ONES[m[1]];
            const tens = TENS[m[2]];
            if (ones !== undefined && tens !== undefined) return tens + ones;
        }
        return null;
    }

    function convertGermanWords(text) {
        let parts = text.trim().split(/\s+/);
        let negative = false;
        let result = 0;
        let decimalStr = '';
        let afterDecimal = false;

        for (const part of parts) {
            if (part === 'minus' || part === '-') { negative = true; continue; }
            if (part === '.' || part === 'punkt' || part === 'komma') { afterDecimal = true; continue; }

            if (afterDecimal) {
                const n = ONES[part] !== undefined ? ONES[part] : parseFloat(part);
                if (!isNaN(n)) decimalStr += String(Math.round(n));
                continue;
            }

            const compound = parseCompound(part);
            if (compound !== null) { result += compound; continue; }
            if (ONES[part] !== undefined) { result += ONES[part]; continue; }
            if (TENS[part] !== undefined) { result += TENS[part]; continue; }
            if (part === 'hundert') { result = result === 0 ? 100 : result * 100; continue; }
            if (part === 'tausend') { result = result === 0 ? 1000 : result * 1000; continue; }
        }

        if (result === 0 && decimalStr === '') return null;
        const final = (negative ? '-' : '') + result + (decimalStr ? '.' + decimalStr : '');
        return final;
    }

    function parseSpokenNumber(raw) {
        let text = raw.trim().toLowerCase();

        // Strip filler words before number
        text = text.replace(/\bist\b/g, '').replace(/\bgleich\b/g, '').replace(/\bergibt\b/g, '').replace(/=/g, '').trim();

        // Normalise German decimal forms
        text = text.replace(/\bkomma\b/g, ' . ');
        text = text.replace(/\bpunkt\b/g, ' . ');
        text = text.replace(/\bminus\b/g, '-');
        text = text.replace(/\bnegativ\b/g, '-');

        // Strip common units
        text = text.replace(/\bohm\b/g, '').replace(/\bvolt\b/g, '')
                   .replace(/\bampere\b/g, '').replace(/\bmeter\b/g, '')
                   .replace(/\bkiloohm\b/g, '').replace(/\bΩ/g, '');

        // German comma as decimal separator: "45,3" → "45.3"
        text = text.replace(/(\d),(\d)/g, '$1.$2');

        text = text.trim();

        // Direct numeric parse (most speech engines output digits already)
        const direct = parseFloat(text.replace(',', '.'));
        if (!isNaN(direct) && String(direct).length > 0) return String(direct);

        // Fallback: convert German word numbers
        return convertGermanWords(text);
    }

    // ── Fill a single cell ─────────────────────────────────────────────────────
    function fillCell(inp, value) {
        inp.value = value;
        inp.style.opacity = '1';
        // Trigger the app's oninput handler (recalculates ρ/MW/SD, saves, etc.)
        inp.dispatchEvent(new Event('input', { bubbles: true }));
        inp.dispatchEvent(new Event('change', { bubbles: true }));
    }

    // ── Single-cell recognition ────────────────────────────────────────────────
    function createRecognition(lang) {
        const rec = new SpeechRecognition();
        rec.lang = lang || 'de-DE';
        rec.continuous = false;
        rec.interimResults = true;
        rec.maxAlternatives = 3;
        return rec;
    }

    function startListening(fab) {
        if (!activeInput) {
            if (typeof showToast === 'function') showToast('Zelle antippen, dann Mikrofon drücken');
            return;
        }
        if (isListening) { stopListening(fab); return; }

        recognition = createRecognition('de-DE');
        isListening = true;
        fab.classList.add('listening');
        fab.innerHTML = '<i class="fas fa-stop"></i>';
        fab.title = 'Aufnahme stoppen';

        const preview = document.getElementById('voicePreviewBar');
        const previewText = document.getElementById('voicePreviewText');
        if (preview) preview.classList.add('show');

        recognition.onresult = (event) => {
            let interim = '';
            let final = '';
            for (let i = event.resultIndex; i < event.results.length; i++) {
                if (event.results[i].isFinal) {
                    final += event.results[i][0].transcript;
                } else {
                    interim += event.results[i][0].transcript;
                }
            }

            if (previewText) previewText.textContent = (interim || final || '...');

            if (final) {
                const parsed = parseSpokenNumber(final);
                if (parsed !== null && parsed !== '' && !isNaN(parseFloat(parsed))) {
                    fillCell(activeInput, parsed);
                    if (typeof showToast === 'function') showToast(`✓ "${final.trim()}" → ${parsed}`);
                } else {
                    // Not a number — still fill (e.g. Kommentar field)
                    fillCell(activeInput, final.trim());
                    if (typeof showToast === 'function') showToast(`✓ "${final.trim()}"`);
                }
                stopListening(fab);
            }
        };

        recognition.onerror = (event) => {
            if (event.error === 'no-speech') {
                if (typeof showToast === 'function') showToast('Kein Sprachsignal — bitte noch einmal sprechen');
            } else if (event.error === 'not-allowed') {
                if (typeof showToast === 'function') showToast('Mikrofon-Zugriff verweigert — bitte im Browser erlauben');
            } else if (event.error !== 'aborted') {
                if (typeof showToast === 'function') showToast('Sprache nicht erkannt');
            }
            stopListening(fab);
        };

        recognition.onend = () => { if (isListening) stopListening(fab); };

        try {
            recognition.start();
        } catch (err) {
            console.warn('[Voice] Recognition start error:', err);
            stopListening(fab);
        }
    }

    function stopListening(fab) {
        isListening = false;
        if (fab) {
            fab.classList.remove('listening');
            fab.innerHTML = '<i class="fas fa-microphone"></i>';
            fab.title = 'Spracheingabe starten';
        }
        if (activeInput) activeInput.style.opacity = '1';
        const preview = document.getElementById('voicePreviewBar');
        if (preview) {
            setTimeout(() => preview.classList.remove('show'), 800);
        }
        try { recognition && recognition.abort(); } catch (e) { /* ignore */ }
        recognition = null;
    }

    // ── Row dictation mode ─────────────────────────────────────────────────────
    // User says: "R1 45, R2 47 Komma 5, R3 50" → fills all R fields of the selected row

    // Keep dictation state outside so it survives modal re-opens
    let dictRecognition = null;
    let dictListening = false;

    function dictStop() {
        dictListening = false;
        const mb = document.getElementById('dictationMicBtn');
        if (mb) { mb.classList.remove('listening'); mb.innerHTML = '<i class="fas fa-microphone"></i>'; }
        try { dictRecognition && dictRecognition.abort(); } catch (e) {}
        dictRecognition = null;
    }

    function startRowDictation() {
        const modal = document.getElementById('voiceDictationModal');
        if (!modal) return;

        // Determine which cell is selected
        let cellLabel = '—';
        let rowIdx = -1;
        let colName = '';
        if (typeof AppState !== 'undefined' && AppState.selectedCell) {
            const parts = AppState.selectedCell.split('-');
            rowIdx = parseInt(parts[0]);
            colName = parts.slice(1).join('-') || '';
            // Format nice label: "Zeile 1: R1/0,8"
            const shortCol = colName.replace(' [Ω]', '').replace('_', '/').replace('.', ',');
            cellLabel = `Zeile ${rowIdx + 1}: ${shortCol}`;
        }

        // Reset UI
        const transcript = document.getElementById('dictationTranscript');
        const resultEl   = document.getElementById('dictationResult');
        const micBtn     = document.getElementById('dictationMicBtn');
        if (transcript) transcript.textContent = 'Mikrofon drücken, Zahl sprechen';
        if (resultEl)   resultEl.textContent   = cellLabel;
        if (micBtn)     { micBtn.classList.remove('listening'); micBtn.innerHTML = '<i class="fas fa-microphone"></i>'; }

        modal.classList.add('show');

        // Wire up buttons only once
        if (!modal._voiceReady) {
            modal._voiceReady = true;

            document.getElementById('dictationMicBtn').addEventListener('click', () => {
                const btn = document.getElementById('dictationMicBtn');
                const tx  = document.getElementById('dictationTranscript');
                const res = document.getElementById('dictationResult');

                if (dictListening) { dictStop(); return; }

                // Re-read selected cell
                let rIdx = -1, cName = '';
                if (typeof AppState !== 'undefined' && AppState.selectedCell) {
                    const p = AppState.selectedCell.split('-');
                    rIdx = parseInt(p[0]);
                    cName = p.slice(1).join('-') || '';
                }
                if (rIdx < 0 || !cName) {
                    if (typeof showToast === 'function') showToast('Bitte zuerst eine Zelle in der Tabelle auswählen');
                    return;
                }

                dictRecognition = createRecognition('de-DE');
                dictRecognition.continuous = false;
                dictRecognition.interimResults = true;
                dictListening = true;
                btn.classList.add('listening');
                btn.innerHTML = '<i class="fas fa-stop"></i>';
                if (tx) tx.textContent = '🎙 Zahl sprechen...';

                dictRecognition.onresult = (event) => {
                    let interim = '', final = '';
                    for (let i = event.resultIndex; i < event.results.length; i++) {
                        if (event.results[i].isFinal) final  += event.results[i][0].transcript;
                        else                          interim += event.results[i][0].transcript;
                    }
                    if (tx) tx.textContent = final.trim() || interim || '...';
                    if (final.trim()) {
                        const parsed = parseSpokenNumber(final.trim());
                        const val = (parsed !== null && !isNaN(parseFloat(parsed))) ? parsed : null;
                        if (val !== null) {
                            const row = AppState.data[rIdx];
                            if (row) {
                                row[cName] = val;
                                if (typeof renderTable === 'function') renderTable();
                                if (typeof debouncedSave === 'function') debouncedSave(800);
                            }
                            const shortCol = cName.replace(' [Ω]', '').replace('_', '/').replace('.', ',');
                            if (res) res.textContent = `✓ Zeile ${rIdx + 1} ${shortCol} = ${val}`;
                            if (typeof showToast === 'function') showToast(`✓ "${final.trim()}" → ${val}`);
                        } else {
                            if (res) res.textContent = `Nicht erkannt: "${final.trim()}"`;
                        }
                        dictStop();
                    }
                };

                dictRecognition.onerror = (e) => {
                    if (e.error !== 'no-speech' && e.error !== 'aborted') {
                        if (tx) tx.textContent = 'Fehler: ' + e.error;
                    }
                    dictStop();
                };
                dictRecognition.onend = () => { if (dictListening) dictStop(); };
                try { dictRecognition.start(); } catch (e) { dictStop(); }
            });

            document.getElementById('btnCloseDictation').addEventListener('click', () => {
                dictStop();
                document.getElementById('voiceDictationModal').classList.remove('show');
            });
        }
    }

    // Parse "R1 45 R2 47 R3 50" or "R1 fünfzig, R2 ..." into table fields
    function parseRowDictation(text, rowIdx) {
        if (typeof AppState === 'undefined') return 0;
        const row = AppState.data[rowIdx];
        if (!row) return 0;

        let filled = 0;

        // Strategy: split by known field labels (R1, R2, R3, Pot, Spannung, Strom)
        // Pattern: label followed by number
        const fieldPatterns = [
            { re: /\bR\s*1\b/i, col: (depth) => `R1 [Ω]_${depth}` },
            { re: /\bR\s*2\b/i, col: (depth) => `R2 [Ω]_${depth}` },
            { re: /\bR\s*3\b/i, col: (depth) => `R3 [Ω]_${depth}` },
        ];

        // Detect depth from visible columns or selected cell context
        const depths = ['0.8', '1.6', '3.2'];
        let targetDepth = '0.8'; // default
        if (typeof AppState !== 'undefined' && AppState.selectedCell) {
            const colName = AppState.selectedCell.split('-')[1] || '';
            for (const d of depths) {
                if (colName.endsWith('_' + d)) { targetDepth = d; break; }
            }
        }

        // Tokenise: "R1 fünfzig R2 siebenundvierzig" → [{label:'R1', raw:'fünfzig'}, ...]
        // Simple approach: find label positions and extract text between them
        const labelRe = /\b(R\s*1|R\s*2|R\s*3|pot(?:ential)?|spannung|strom|kommentar)\b/gi;
        const tokens = [];
        let lastMatch = null;
        let m;
        labelRe.lastIndex = 0;
        while ((m = labelRe.exec(text)) !== null) {
            if (lastMatch) {
                const raw = text.slice(lastMatch.index + lastMatch[0].length, m.index).trim();
                tokens.push({ label: lastMatch[1].toLowerCase().replace(/\s/g, ''), raw });
            }
            lastMatch = m;
        }
        if (lastMatch) {
            const raw = text.slice(lastMatch.index + lastMatch[0].length).trim();
            tokens.push({ label: lastMatch[1].toLowerCase().replace(/\s/g, ''), raw });
        }

        if (tokens.length === 0) return 0;

        // Map labels to column names and fill
        const labelToCol = {
            'r1': `R1 [Ω]_${targetDepth}`,
            'r2': `R2 [Ω]_${targetDepth}`,
            'r3': `R3 [Ω]_${targetDepth}`,
        };

        tokens.forEach(({ label, raw }) => {
            const colName = labelToCol[label];
            if (!colName) return;

            // Strip trailing punctuation/comma from raw
            const cleanRaw = raw.replace(/[,;.]$/, '').trim();
            const parsed = parseSpokenNumber(cleanRaw);
            const val = (parsed !== null && !isNaN(parseFloat(parsed))) ? parsed : cleanRaw;

            if (val) {
                row[colName] = val;
                filled++;
            }
        });

        if (filled > 0 && typeof renderTable === 'function') {
            renderTable();
            if (typeof debouncedSave === 'function') debouncedSave(800);
        }

        return filled;
    }

    // ── Events ─────────────────────────────────────────────────────────────────
    function initEvents(fab) {
        // Track focused cell input
        document.addEventListener('focusin', (e) => {
            const inp = e.target;
            if (inp.tagName === 'INPUT' && inp.classList.contains('cell-input') && !inp.readOnly) {
                activeInput = inp;
                activeRow = inp.dataset ? parseInt(inp.dataset.row) : -1;
                fab.classList.add('active');
                fab.title = 'Spracheingabe starten';
            }
        });

        document.addEventListener('focusout', () => {
            setTimeout(() => {
                const focused = document.activeElement;
                if (!focused || !focused.classList || !focused.classList.contains('cell-input')) {
                    if (!isListening) {
                        // Only deactivate FAB if not listening (don't lose state mid-listen)
                        activeInput = null;
                        fab.classList.remove('active');
                        fab.title = 'Zelle antippen für Spracheingabe';
                    }
                }
            }, 250);
        });

        // FAB click → toggle listening
        fab.addEventListener('click', () => startListening(fab));

        // Keyboard shortcut: hold Space while a numeric cell is focused (desktop use)
        document.addEventListener('keydown', (e) => {
            if (e.code === 'Space' && e.shiftKey && activeInput && !isListening) {
                e.preventDefault();
                startListening(fab);
            }
        });

        // Expose row dictation globally (called from toolbar button)
        window.openVoiceDictation = startRowDictation;
    }

    // ── Boot ───────────────────────────────────────────────────────────────────
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => injectUI(true));
    } else {
        injectUI(true);
    }

    console.log('[Voice] Spracheingabe bereit — Zelle antippen, dann Mikrofon drücken (oder Shift+Space)');

})();
