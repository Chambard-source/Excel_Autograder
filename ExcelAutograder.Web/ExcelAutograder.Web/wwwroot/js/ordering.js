/**
 * Section Ordering + Rendering Hooks
 * ----------------------------------
 * Extends the rubric UI to support a global section order, keep sheet keys
 * aligned with that order, and surface simple controls (↑/↓/Remove/Add/Clear)
 * to manage the ordering. Also wraps `render()` and `sortKey()` so existing
 * UI lists follow the section order automatically.
 */
(function () {
    // Ensure a global rubric object exists
    window.rubric = window.rubric || { meta: {}, sheets: {} };

    // === Keep sheets object key order aligned with the grading section order ===

    /**
     * Reorders rubric.sheets keys to follow the first occurrence of any of each sheet's
     * sections within the global rubric.meta.sectionOrder. Ties fall back to name sort.
     *
     * Heuristic:
     *  - For each sheet, find the earliest index of any of its `section_order` entries
     *    inside the global `meta.sectionOrder`. Use that as the sheet's rank.
     *  - Sort sheets by rank, then by sheet name.
     *
     * Side Effect:
     *  - Mutates `rubric.sheets` by replacing it with a newly constructed ordered object.
     */
    function reorderSheetsBySectionOrder() {
        try {
            if (!rubric || !rubric.sheets) return;

            const global = Array.isArray(rubric.meta?.sectionOrder) ? rubric.meta.sectionOrder : [];

            const arr = Object.entries(rubric.sheets).map(([name, spec]) => {
                const list = Array.isArray(spec.section_order) ? spec.section_order : [];

                // Find the earliest global index matched by any section in this sheet
                let idx = Number.POSITIVE_INFINITY;
                for (let i = 0; i < global.length; i++) {
                    if (list.includes(global[i])) { idx = i; break; }
                }
                return { name, spec, idx };
            });

            arr.sort((a, b) => (a.idx - b.idx) || a.name.localeCompare(b.name));

            const ordered = {};
            for (const x of arr) ordered[x.name] = x.spec;
            rubric.sheets = ordered;
        } catch (e) { /* no-op: keep existing order on failure */ }
    }
    // Expose for external callers that may want to force a reorder
    window.reorderSheetsBySectionOrder = reorderSheetsBySectionOrder;

    /**
     * Returns a copy of the current global section order.
     * @returns {string[]} The section names in order.
     */
    function getSectionOrder() {
        const m = rubric.meta || {};
        return Array.isArray(m.sectionOrder) ? m.sectionOrder.slice() : [];
    }

    /**
     * Sets the global section order, marks it as strict, and triggers
     * dependent UI/ordering refreshes (sheet order, order UI, renderers).
     *
     * @param {string[]} arr - New ordered list of section names.
     */
    function setSectionOrder(arr) {
        rubric.meta = rubric.meta || {};
        rubric.meta.sectionOrder = arr.slice();
        rubric.meta.strictSectionOrder = true;

        // Keep sheet keys aligned with the new global order
        try { reorderSheetsBySectionOrder(); } catch (e) { }

        // Update the controls UI and re-render lists/JSON preview, if available
        try { refreshOrderUI(); } catch (e) { }
        try { if (typeof render === 'function') render(); } catch (e) { }
        try { if (typeof refreshJsonPreview === 'function') refreshJsonPreview(); } catch (e) { }
    }

    /**
     * Scans all rules across all sheets and returns a sorted list of unique section names.
     * @returns {string[]} Unique section names (localeCompare sorted).
     */
    function allSectionNames() {
        const set = new Set();
        const sheets = rubric.sheets || {};
        Object.values(sheets).forEach(s => (s.checks || []).forEach(r => {
            const nm = (r.section || '').trim();
            if (nm) set.add(nm);
        }));
        return Array.from(set).sort((a, b) => a.localeCompare(b));
    }

    /**
     * Rebuilds the Section Order management UI:
     *  - Renders current ordered sections with ↑/↓/Remove controls.
     *  - Populates a picker with remaining (unordered) sections.
     *  - Wires "Add" and "Clear" buttons to update the global order.
     *
     * Expected DOM:
     *  - #sectionOrder (container for ordered rows)
     *  - #sectionOrderPicker (select for remaining sections)
     *  - #btnOrderAdd (button) and #btnOrderClear (button)
     */
    function refreshOrderUI() {
        const host = document.getElementById('sectionOrder');
        const pick = document.getElementById('sectionOrderPicker');
        if (!host || !pick) return;

        const order = getSectionOrder();
        const all = allSectionNames();
        const remaining = all.filter(n => !order.includes(n));

        // Render ordered list
        host.innerHTML = '';
        order.forEach((name, i) => {
            const row = document.createElement('div'); row.className = 'order-row';
            const up = document.createElement('button'); up.className = 'ghost'; up.textContent = '↑';
            const down = document.createElement('button'); down.className = 'ghost'; down.textContent = '↓';
            const del = document.createElement('button'); del.className = 'ghost danger'; del.textContent = 'Remove';
            const nm = document.createElement('span'); nm.className = 'name'; nm.textContent = name;
            const idx = document.createElement('span'); idx.className = 'tiny muted'; idx.textContent = '#' + (i + 1);
            const spacer = document.createElement('div'); spacer.className = 'spacer';

            up.onclick = () => {
                if (i === 0) return;
                const arr = getSectionOrder();
                [arr[i - 1], arr[i]] = [arr[i], arr[i - 1]];
                setSectionOrder(arr);
            };
            down.onclick = () => {
                const arr = getSectionOrder();
                if (i >= arr.length - 1) return;
                [arr[i + 1], arr[i]] = [arr[i], arr[i + 1]];
                setSectionOrder(arr);
            };
            del.onclick = () => {
                const arr = getSectionOrder().filter(s => s !== name);
                setSectionOrder(arr);
            };

            row.append(up, down, nm, spacer, idx, del);
            host.append(row);
        });

        // Populate remaining picker
        pick.innerHTML = '';
        if (remaining.length === 0) {
            pick.append(new Option('(All sections ordered)', '', true, true));
        } else {
            remaining.forEach(n => pick.append(new Option(n, n)));
        }

        // Wire add/clear buttons
        const btnAdd = document.getElementById('btnOrderAdd');
        const btnClear = document.getElementById('btnOrderClear');

        if (btnAdd) btnAdd.onclick = () => {
            const v = pick.value;
            if (!v) return;
            const arr = getSectionOrder();
            arr.push(v);
            setSectionOrder(arr);
        };
        if (btnClear) btnClear.onclick = () => setSectionOrder([]);
    }

    // --- Sorting hook for rule lists ----------------------------------------

    /**
     * Wraps/overrides a global `sortKey(rule, index)` to incorporate section order.
     * Fallbacks to prior sortKey (if it existed) or index when errors occur.
     *
     * Key format: "<secRank>|<type>|<cell>|<note>|<index>"
     *  - secRank: zero-padded index of section in global order, or '999' if not present.
     */
    const oldSortKey = (typeof sortKey === 'function') ? sortKey : null;
    window.sortKey = function (rule, index) {
        try {
            const sec = (rule.section || '').trim();
            const order = (rubric.meta && Array.isArray(rubric.meta.sectionOrder)) ? rubric.meta.sectionOrder : [];
            const idx = order.indexOf(sec);
            const secRank = (idx >= 0) ? String(idx).padStart(3, '0') : '999';

            const t = (rule.type || '').toString();
            const c = (rule.cell || '').toString();
            const n = (rule.note || '').toString();

            return secRank + '|' + t + '|' + c + '|' + n + '|' + String(index).padStart(6, '0');
        } catch (e) {
            return oldSortKey ? oldSortKey(rule, index) : (String(index).padStart(6, '0'));
        }
    };

    // --- Render hook to refresh the order UI after each render ----------------

    /**
     * If a global `render()` exists, wrap it so that after every render
     * we rebuild the section-order controls to reflect any changes.
     */
    if (typeof render === 'function') {
        const _orig = render;
        window.render = function () {
            const result = _orig.apply(this, arguments);
            try { refreshOrderUI(); } catch (e) { }
            return result;
        };
    }

    /**
     * Initial boot-time refresh of the order UI once the DOM is ready.
     */
    document.addEventListener('DOMContentLoaded', () => {
        try { refreshOrderUI(); } catch (e) { }
    });
})();

/**
 * Boot-Time Sheet Reorder (idempotent)
 * ------------------------------------
 * On load, if a global sectionOrder is present, attempt to normalize the
 * sheet key ordering to match. This runs independently of the main IIFE
 * and is safe if run before/after it due to guards.
 */
; (() => {
    try {
        if (window.rubric && window.rubric.meta && Array.isArray(window.rubric.meta.sectionOrder)) {
            // ensure a consistent sheet order based on global section ordering
            if (typeof window.reorderSheetsBySectionOrder === 'function') {
                window.reorderSheetsBySectionOrder(window.rubric);
            }
        }
    } catch (e) { console.warn('reorderSheetsBySectionOrder on boot failed', e); }
})();
