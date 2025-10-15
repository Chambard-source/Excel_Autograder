window.addEventListener('DOMContentLoaded', function () {
            // --- elements (support either id your page uses) ---
            const keyInput = document.getElementById('keyFile') || document.getElementById('key');
            const includeArtifacts = document.getElementById('includeArtifacts');
            const rangeTotalPoints = document.getElementById('rangeTotalPoints');
            const btnBuildFromRanges = document.getElementById('btnBuildFromRanges');

            // Single-sheet panel elements (legacy)
            const rangeSheet = document.getElementById('rangeSheet');
            const rangeRows = document.getElementById('rangeRows');
            const btnAddRangeRow = document.getElementById('btnAddRangeRow');

            // Multi-sheet panel elements (new)
            const groupsEl = document.getElementById('rangeGroups');
            const btnAddRangeSheet = document.getElementById('btnAddRangeSheet');

            // --- common helpers ---
            let sheetNames = [];
            async function loadSheetNames() {
                try {
                    const file = keyInput?.files?.[0];
                    if (!file) return [];
                    const fd = new FormData();
                    fd.append('key', file);
                    const resp = await fetch('/api/key/sheets', { method: 'POST', body: fd });
                    if (!resp.ok) return [];
                    sheetNames = await resp.json();
                    return sheetNames;
                } catch { return []; }
            }

            function a1Row(name = '', ranges = '') {
                const row = document.createElement('div');
                row.className = 'hstack';
                row.style.gap = '8px';
                row.innerHTML = `
                                              <input class="sectionName" type="text" placeholder="Section name (e.g., Totals)" style="min-width:220px" value="${name}">
                                              <input class="sectionRanges" type="text" placeholder="A1 ranges (comma-separated, e.g., A2:B20, E8:E18)" style="min-width:360px" value="${ranges}">
                                              <button class="ghost removeRowBtn" type="button">Remove</button>
                                            `;
                row.querySelector('.removeRowBtn').onclick = () => row.remove();
                return row;
            }

            // --- SINGLE-SHEET MODE (if #rangeSheet exists) ---
            if (rangeSheet && rangeRows) {
                // seed one row
                if (rangeRows.children.length === 0) rangeRows.appendChild(a1Row());
                btnAddRangeRow?.addEventListener('click', () => rangeRows.appendChild(a1Row()));

                keyInput?.addEventListener('change', async () => {
                    rangeSheet.innerHTML = `<option value="">Select a sheet…</option>`;
                    const names = await loadSheetNames();
                    for (const n of names) rangeSheet.appendChild(new Option(n, n));
                });
            }

            // --- MULTI-SHEET MODE (if #rangeGroups exists) ---
            function addGroup(defaultSheet) {
                const g = document.createElement('div');
                g.className = 'card stack';
                g.style.padding = '12px';
                g.innerHTML = `
                                          <div class="hstack" style="gap:8px;align-items:center;flex-wrap:wrap;">
                                            <select class="sheetSel" style="min-width:220px"></select>
                                            <button class="ghost addRowBtn" type="button">+ Add Section + Ranges</button>
                                            <button class="danger removeGroupBtn" type="button" style="margin-left:auto">Remove sheet</button>
                                          </div>
                                          <div class="stack rows" style="gap:8px;margin-top:8px;"></div>
                                        `;
                const sheetSel = g.querySelector('.sheetSel');
                const rowsEl = g.querySelector('.rows');

                for (const name of sheetNames) sheetSel.appendChild(new Option(name, name));
                if (defaultSheet && sheetNames.includes(defaultSheet)) sheetSel.value = defaultSheet;

                function addRow(n = '', r = '') { rowsEl.appendChild(a1Row(n, r)); }
                g.querySelector('.addRowBtn').onclick = () => addRow();
                g.querySelector('.removeGroupBtn').onclick = () => g.remove();
                addRow();

                groupsEl.appendChild(g);
            }

            if (groupsEl && btnAddRangeSheet) {
                btnAddRangeSheet.addEventListener('click', async () => {
                    if (sheetNames.length === 0) await loadSheetNames();
                    if (sheetNames.length === 0) return alert('Choose a key workbook first.');
                    addGroup(sheetNames[0]);
                });

                keyInput?.addEventListener('change', async () => { await loadSheetNames(); });
            }

            // --- Build button (works for either mode) ---
            btnBuildFromRanges?.addEventListener('click', async () => {
                const file = keyInput?.files?.[0];
                if (!file) return alert('Choose a key workbook first.');

                // Compose sections_json:
                //  - multi-sheet: from groups
                //  - single-sheet: from drop-down + rows
                const payload = {};

                if (groupsEl && groupsEl.children.length > 0) {
                    // multi-sheet
                    for (const g of groupsEl.children) {
                        const sheet = (g.querySelector('.sheetSel')?.value || '').trim();
                        if (!sheet) continue;
                        const rows = Array.from(g.querySelectorAll('.rows .hstack'));
                        const sections = [];
                        for (const r of rows) {
                            const name = (r.querySelector('.sectionName')?.value || '').trim() || 'Section';
                            const rangesText = (r.querySelector('.sectionRanges')?.value || '').trim();
                            if (!rangesText) continue;
                            const ranges = rangesText.split(',').map(s => s.trim()).filter(Boolean);
                            if (ranges.length) sections.push({ section: name, ranges });
                        }
                        if (sections.length) payload[sheet] = sections;
                    }
                } else if (rangeSheet && rangeRows) {
                    // single-sheet
                    const sheet = (rangeSheet.value || '').trim();
                    if (!sheet) return alert('Pick a sheet.');
                    const rows = Array.from(rangeRows.querySelectorAll('.hstack'));
                    const sections = [];
                    for (const r of rows) {
                        const name = (r.querySelector('.sectionName')?.value || '').trim() || 'Section';
                        const rangesText = (r.querySelector('.sectionRanges')?.value || '').trim();
                        if (!rangesText) continue;
                        const ranges = rangesText.split(',').map(s => s.trim()).filter(Boolean);
                        if (ranges.length) sections.push({ section: name, ranges });
                    }
                    if (!sections.length) return alert('Add at least one row with ranges.');
                    payload[sheet] = sections;
                } else {
                    return alert('Add at least one sheet/section.');
                }

                const fd = new FormData();
                fd.append('key', file);
                fd.append('sections_json', JSON.stringify(payload));
                if (includeArtifacts?.checked) fd.append('include_artifacts', 'true');
                if (rangeTotalPoints?.value) fd.append('total', rangeTotalPoints.value);

                btnBuildFromRanges.disabled = true;
                const orig = btnBuildFromRanges.textContent;
                btnBuildFromRanges.textContent = 'Building…';
                try {
                    const resp = await fetch('/api/rubric/from-ranges', { method: 'POST', body: fd });
                    if (!resp.ok) {
                        const err = await resp.json().catch(() => ({}));
                        throw new Error(err.error || resp.statusText);
                    }
                    const rub = await resp.json();
                    if (window.renderBuilderFromRubric) {
                        window.renderBuilderFromRubric(rub);
                    } else {
                        rubric = rub;
                        if (window.render) render();
                        if (window.refreshJsonPreview) refreshJsonPreview();
                    }
                } catch (e) {
                    alert('Failed to build from sections: ' + e.message);
                    console.error(e);
                } finally {
                    btnBuildFromRanges.disabled = false;
                    btnBuildFromRanges.textContent = orig;
                }
            });
        });
    