
        (function () {
            window.rubric = window.rubric || { meta: {}, sheets: {} };
            
// === Keep sheets object key order aligned with the grading section order ===
function reorderSheetsBySectionOrder() {
    try {
        if (!rubric || !rubric.sheets) return;
        const global = Array.isArray(rubric.meta?.sectionOrder) ? rubric.meta.sectionOrder : [];
        const arr = Object.entries(rubric.sheets).map(([name, spec]) => {
            const list = Array.isArray(spec.section_order) ? spec.section_order : [];
            // find first index of any section from this sheet in the global order
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
    } catch (e) { /* no-op */ }
}

            function getSectionOrder() {
                const m = rubric.meta || {};
                return Array.isArray(m.sectionOrder) ? m.sectionOrder.slice() : [];
            }
            function setSectionOrder(arr) {
                rubric.meta = rubric.meta || {};
                rubric.meta.sectionOrder = arr.slice();
                    rubric.meta.strictSectionOrder = true;
    try { reorderSheetsBySectionOrder(); } catch (e) { }
    try { refreshOrderUI(); } catch (e) { }
                try { if (typeof render === 'function') render(); } catch (e) { }
                try { if (typeof refreshJsonPreview === 'function') refreshJsonPreview(); } catch (e) { }
            }
            function allSectionNames() {
                const set = new Set();
                const sheets = rubric.sheets || {};
                Object.values(sheets).forEach(s => (s.checks || []).forEach(r => {
                    const nm = (r.section || '').trim();
                    if (nm) set.add(nm);
                }));
                return Array.from(set).sort((a, b) => a.localeCompare(b));
            }
            function refreshOrderUI() {
                const host = document.getElementById('sectionOrder');
                const pick = document.getElementById('sectionOrderPicker');
                if (!host || !pick) return;
                const order = getSectionOrder();
                const all = allSectionNames();
                const remaining = all.filter(n => !order.includes(n));

                host.innerHTML = '';
                order.forEach((name, i) => {
                    const row = document.createElement('div'); row.className = 'order-row';
                    const up = document.createElement('button'); up.className = 'ghost'; up.textContent = '↑';
                    const down = document.createElement('button'); down.className = 'ghost'; down.textContent = '↓';
                    const del = document.createElement('button'); del.className = 'ghost danger'; del.textContent = 'Remove';
                    const nm = document.createElement('span'); nm.className = 'name'; nm.textContent = name;
                    const idx = document.createElement('span'); idx.className = 'tiny muted'; idx.textContent = '#' + (i + 1);
                    const spacer = document.createElement('div'); spacer.className = 'spacer';

                    up.onclick = () => { if (i === 0) return; const arr = getSectionOrder();[arr[i - 1], arr[i]] = [arr[i], arr[i - 1]]; setSectionOrder(arr); };
                    down.onclick = () => { const arr = getSectionOrder(); if (i >= arr.length - 1) return;[arr[i + 1], arr[i]] = [arr[i], arr[i + 1]]; setSectionOrder(arr); };
                    del.onclick = () => { const arr = getSectionOrder().filter(s => s !== name); setSectionOrder(arr); };

                    row.append(up, down, nm, spacer, idx, del);
                    host.append(row);
                });

                pick.innerHTML = '';
                if (remaining.length === 0) {
                    pick.append(new Option('(All sections ordered)', '', true, true));
                } else {
                    remaining.forEach(n => pick.append(new Option(n, n)));
                }
                const btnAdd = document.getElementById('btnOrderAdd');
                const btnClear = document.getElementById('btnOrderClear');
                if (btnAdd) btnAdd.onclick = () => {
                    const v = pick.value;
                    if (!v) return;
                    const arr = getSectionOrder(); arr.push(v); setSectionOrder(arr);
                };
                if (btnClear) btnClear.onclick = () => setSectionOrder([]);
            }

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

            if (typeof render === 'function') {
                const _orig = render;
                window.render = function () {
                    const result = _orig.apply(this, arguments);
                    try { refreshOrderUI(); } catch (e) { }
                    return result;
                };
            }

            document.addEventListener('DOMContentLoaded', () => {
                try { refreshOrderUI(); } catch (e) { }
            });
        })();

;(() => {
  try {
    if (window.rubric && window.rubric.meta && Array.isArray(window.rubric.meta.sectionOrder)) {
      // ensure a consistent sheet order based on global section ordering
      if (typeof window.reorderSheetsBySectionOrder === 'function') {
         window.reorderSheetsBySectionOrder(window.rubric);
      }
    }
  } catch (e) { console.warn('reorderSheetsBySectionOrder on boot failed', e); }
})();
