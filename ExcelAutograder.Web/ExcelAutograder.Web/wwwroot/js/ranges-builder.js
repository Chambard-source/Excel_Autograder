/**
 * Bootstraps the "Build From Ranges" UI once the DOM is ready.
 * Wires up inputs, loads sheet names from the uploaded key workbook,
 * and builds the sections_json payload (single-sheet or multi-sheet modes)
 * to request a rubric from /api/rubric/from-ranges.
 */
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

    /**
     * Fetches available sheet names from the uploaded key workbook.
     *
     * REQUEST: POST /api/key/sheets with form-data key=<File>
     * @returns {Promise<string[]>} Resolves to an array of sheet names; [] if not available or on failure.
     *
     * Notes:
     * - Reads the selected file from `keyInput`.
     * - Gracefully returns [] on any error (network, no file, non-OK response).
     */
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
        } catch {
            return [];
        }
    }

    /**
     * Creates a single "Section + Ranges" row element for either mode.
     *
     * @param {string} [name='']    Initial value for the section name input.
     * @param {string} [ranges='']  Initial value for the ranges input (comma-separated A1 ranges).
     * @returns {HTMLDivElement}    A <div> row containing inputs and a "Remove" button.
     *
     * Structure:
     * <div.hstack>
     *   <input.sectionName />
     *   <input.sectionRanges />
     *   <button.removeRowBtn />
     * </div>
     */
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

        /**
         * Adds a blank "Section + Ranges" row to the single-sheet panel.
         * Bound to the "+ Add Range" button when present.
         */
        btnAddRangeRow?.addEventListener('click', () => rangeRows.appendChild(a1Row()));

        /**
         * When the user selects a key workbook, load its sheet names and populate the dropdown.
         */
        keyInput?.addEventListener('change', async () => {
            rangeSheet.innerHTML = `<option value="">Select a sheet…</option>`;
            const names = await loadSheetNames();
            for (const n of names) rangeSheet.appendChild(new Option(n, n));
        });
    }

    /**
     * Creates and appends a new multi-sheet "group" card, which lets the user:
     *  - choose a sheet (from `sheetNames`)
     *  - add one or more "Section + Ranges" rows for that sheet
     *  - remove the entire sheet group
     *
     * @param {string} [defaultSheet]  Optional sheet name to pre-select if it exists in `sheetNames`.
     * @returns {void}
     *
     * DOM inserted:
     * <div.card.stack>
     *   <div.hstack>
     *     <select.sheetSel>...sheetNames...</select>
     *     <button.addRowBtn />
     *     <button.removeGroupBtn />
     *   </div>
     *   <div.rows> <!-- holds a1Row() rows --> </div>
     * </div>
     */
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

        // Populate sheet list
        for (const name of sheetNames) sheetSel.appendChild(new Option(name, name));
        if (defaultSheet && sheetNames.includes(defaultSheet)) sheetSel.value = defaultSheet;

        /**
         * Adds a single "Section + Ranges" row inside this group's rows container.
         * @param {string} [n=''] Pre-filled section name.
         * @param {string} [r=''] Pre-filled ranges string.
         */
        function addRow(n = '', r = '') { rowsEl.appendChild(a1Row(n, r)); }

        // Wire up buttons for this group
        g.querySelector('.addRowBtn').onclick = () => addRow();
        g.querySelector('.removeGroupBtn').onclick = () => g.remove();

        // Start with one empty row to guide the user
        addRow();

        groupsEl.appendChild(g);
    }

    // --- MULTI-SHEET MODE (if #rangeGroups exists) ---
    if (groupsEl && btnAddRangeSheet) {
        /**
         * Adds a new sheet group. If sheet names have not been loaded yet,
         * they are fetched from the selected key workbook first.
         */
        btnAddRangeSheet.addEventListener('click', async () => {
            if (sheetNames.length === 0) await loadSheetNames();
            if (sheetNames.length === 0) return alert('Choose a key workbook first.');
            addGroup(sheetNames[0]); // default to the first available sheet
        });

        /**
         * Refresh available sheet names when user changes the key workbook.
         */
        keyInput?.addEventListener('change', async () => { await loadSheetNames(); });
    }

    /**
     * Builds the sections_json payload (single- or multi-sheet), posts it to
     * /api/rubric/from-ranges with the uploaded key, and renders the resulting rubric.
     *
     * Side effects:
     * - Disables and updates the "Build" button text while in-flight.
     * - On success, calls `window.renderBuilderFromRubric(rub)` if present,
     *   otherwise assigns `rubric = rub` and triggers `render()` / `refreshJsonPreview()` if available.
     * - On failure, shows an alert with the error message and logs to console.
     */
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

        // Build request
        const fd = new FormData();
        fd.append('key', file);
        fd.append('sections_json', JSON.stringify(payload));
        if (includeArtifacts?.checked) fd.append('include_artifacts', 'true');
        if (rangeTotalPoints?.value) fd.append('total', rangeTotalPoints.value);

        // UX: disable button while building
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

            // Prefer dedicated renderer if provided by the page
            if (window.renderBuilderFromRubric) {
                window.renderBuilderFromRubric(rub);
            } else {
                // Fallback: assign and try generic renderers if present
                // eslint-disable-next-line no-undef
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
