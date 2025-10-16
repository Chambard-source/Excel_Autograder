

// ---------- helpers
const $ = s => document.querySelector(s);
const el = (t, c, h) => { const e = document.createElement(t); if (c) e.className = c; if (h != null) e.innerHTML = h; return e; };
const download = (name, text) => { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([text], { type: 'application/json' })); a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 1000); };

// track folded sections per sheet
window.UI = window.UI || {};
UI.selected = UI.selected instanceof Set ? UI.selected : new Set();
UI.sortBy = UI.sortBy || 'section-cell';
UI.groupBySection = UI.groupBySection !== false;   // default ON
UI.fold = UI.fold || {};      // { [sheetName]: Set<sectionName> }
UI.activeSheet = UI.activeSheet || null; // last hovered sheet
UI.railOpen = UI.railOpen || new Set(); // right-rail open sheets

/** Expand or collapse all section groups for a given sheet and re-render the UI.
* @param {string} sheetName - The sheet whose sections to toggle.
* @param {boolean} expand - If true, expand all; if false, collapse all.
*/
function expandAllSections(sheetName, expand) {
    if (!sheetName) return;
    const folded = UI.fold[sheetName] instanceof Set
        ? UI.fold[sheetName]
        : (UI.fold[sheetName] = new Set());

    folded.clear();
    if (!expand) {
        // mark every section as folded
        const spec = (rubric.sheets && rubric.sheets[sheetName]) || {};
        const secs = new Set((spec.checks || []).map(r => r.section || '(No section)'));
        secs.forEach(s => folded.add(s));
    }
    render(); // re-render the current view
}

// Helper: fall back to the first sheet if we don't know which is active

/** Return the current active sheet name, falling back to the first sheet in the rubric.
 * @returns {string}
 */
function getActiveSheetName() {
    if (UI.activeSheet && (rubric.sheets || {})[UI.activeSheet]) return UI.activeSheet;
    const names = Object.keys(rubric.sheets || {});
    return names[0] || '';
}



// ensure every rule has a stable id we can use across reorders/moves

/** Ensure every rule in the rubric has a stable unique id for selection and reordering. */
function ensureRuleIds() {
    UI.ruleSeq = UI.ruleSeq || 1;
    const sheets = rubric.sheets || {};
    for (const spec of Object.values(sheets)) {
        const checks = spec.checks || [];
        for (const r of checks) {
            if (!r.__id) r.__id = `r${UI.ruleSeq++}`;
        }
    }
}

// --- PIVOT HELPERS (UI) ---
/** Normalize and attach a pivot spec to a rule (UI-friendly shape).
 * @param {object} r - Rule object (mutated).
 * @returns {{sheet:(string|null), tableNameLike:(string|null), rows:string[], columns:string[], filters:string[], values:Array<{field:string, agg:string}>}}
 */
function ensurePivot(r) {
    const p = r.pivot || {};
    const arr = v => Array.isArray(v) ? v : (v ? [v] : []);
    const pv = {
        sheet: p.sheet ?? p.Sheet ?? null,
        tableNameLike: p.tableNameLike ?? p.TableNameLike ?? null,
        rows: arr(p.rows ?? p.Rows).filter(Boolean),
        columns: arr(p.columns ?? p.Columns).filter(Boolean),
        filters: arr(p.filters ?? p.Filters).filter(Boolean),
        values: (p.values ?? p.Values ?? []).map(v => ({
            field: v.field ?? v.Field ?? "",
            agg: (v.agg ?? v.Agg ?? "sum").toLowerCase()
        }))
    };
    r.pivot = pv; // write back normalized shape so preview is clean
    return pv;
}

const splitCSV = s => (s || '').split(',').map(x => x.trim()).filter(Boolean);
const joinCSV = arr => (arr || []).join(', ');

/** Build the UI block that edits a rule's Pivot “Values” (field + aggregation) list.
* @param {object} rule - The rule whose pivot values are edited.
* @param {Function} onChange - Callback to invoke after edits (usually refreshJsonPreview).
* @returns {HTMLElement}
*/
function pivotValuesBlock(rule, onChange) {
    const pv = ensurePivot(rule);
    const wrap = el('div', 'stack');

    const list = el('div', 'stack');
    wrap.append(list);

    function oneRow(item, idx) {
        const row = el('div', 'hstack');
        const f = el('input'); f.placeholder = 'Field (e.g., Sales)'; f.style.width = '220px'; f.value = item.field || '';
        const agg = el('select');['sum', 'count', 'average', 'min', 'max'].forEach(a => { const o = el('option'); o.value = o.textContent = a; if ((item.agg || 'sum') === a) o.selected = true; agg.append(o); });
        const del = el('button', 'ghost danger', 'Remove');
        row.append(labelled(f, 'Field'), labelled(agg, 'Aggregation'), del);

        f.onchange = () => { pv.values[idx].field = f.value.trim(); onChange(); };
        agg.onchange = () => { pv.values[idx].agg = agg.value; onChange(); };
        del.onclick = () => { pv.values.splice(idx, 1); render(); };

        return row;
    }

    function refreshList() {
        list.innerHTML = '';
        pv.values.forEach((v, i) => list.append(oneRow(v, i)));
    }
    refreshList();

    const addBtn = el('button', 'ghost', '+ Add Value');
    addBtn.onclick = () => { pv.values.push({ field: '', agg: 'sum' }); refreshList(); onChange(); };
    wrap.append(addBtn);

    return wrap;
}

/** Normalize and attach a chart spec to a rule (UI-friendly shape).
* @param {object} r - Rule object (mutated).
* @returns {{sheet:(string|null),name_like:(string|null),type:(string|null),title:(string|null),title_ref:(string|null),legend_pos:(string|null),data_labels:(boolean|null),x_title:(string|null),y_title:(string|null),series:Array<{name:(string|null),name_ref:(string|null),cat_ref:(string|null),val_ref:(string|null)}>}}
*/
function ensureChart(r) {
    const c = r.chart || {};
    const norm = {
        sheet: c.sheet ?? null,
        name_like: c.name_like ?? null,
        type: c.type ?? null,
        title: c.title ?? null,
        title_ref: c.title_ref ?? null,
        legend_pos: c.legend_pos ?? null,
        data_labels: (typeof c.data_labels === 'boolean') ? c.data_labels : null,
        x_title: c.x_title ?? null,
        y_title: c.y_title ?? null,
        series: Array.isArray(c.series) ? c.series.map(s => ({
            name: s?.name ?? null,
            name_ref: s?.name_ref ?? null,
            cat_ref: s?.cat_ref ?? null,
            val_ref: s?.val_ref ?? null
        })) : []
    };
    r.chart = norm;
    return norm;
}

// ---------- builder state
let rubric = { points: 5, report: { include_pass_fail_column: true, include_comments: true }, sheets: {} };
const RULE_TYPES = ['value', 'formula', 'format', 'range_value', 'range_formula', 'range_format', 'custom_note', 'range_sequence', 'range_numeric', 'pivot_layout', 'conditional_format', 'chart', 'table'];
const emptyRule = (type = 'formula') => ({ type, points: 1, cell: '', range: '', note: '', tolerance: null, any_of: null, expected_from_key: null, expected: null, expected_regex: null, expected_formula: null, allow_regex: null, expected_formula_regex: null, format: null, require: null, start: null, step: null, require_absolute: null });

/** Compute a tuple sort key for a rule based on UI.sortBy setting.
* @param {object} rule
* @param {number} idx - Original index (stable tiebreaker).
* @returns {Array<string|number>}
*/
function sortKey(rule, idx) {
    const sec = (rule.section || '').toLowerCase();
    const cell = (rule.cell || '').toUpperCase();
    const t = (rule.type || '').toLowerCase();
    switch (UI.sortBy) {
        case 'cell': return [cell, t, idx];
        case 'type': return [t, cell, idx];
        case 'section-type-cell': return [sec, t, cell, idx];
        case 'section-cell':
        default: return [sec, cell, idx];
    }
}

/** HTML-escape a string for safe innerHTML/text use.
* @param {unknown} s
* @returns {string}
*/
function esc(s) { return (s ?? '').toString().replace(/&/g, '&amp;').replace(/</g, '&lt;'); }

// -------- Compact rules list used both in sheets and in the right rail

/** Render the compact list of rules for a sheet, optionally grouped by section.
 * @param {string} sheetName
 * @param {{checks?:object[]}} spec - Sheet spec from rubric.
 * @param {HTMLElement} host - Container to render into (overwritten).
 */
function renderRulesList(sheetName, spec, host) {
    host.innerHTML = '';

    const rows = (spec.checks || [])
        .map((r, i) => [r, i])
        .sort((a, b) => {
            const ka = sortKey(a[0], a[1]), kb = sortKey(b[0], b[1]);
            return ka < kb ? -1 : ka > kb ? 1 : 0;
        });

    const folded = UI.fold[sheetName] instanceof Set ? UI.fold[sheetName] : (UI.fold[sheetName] = new Set());

    if (UI.groupBySection) {
        const groups = new Map();
        for (const [r, i] of rows) {
            const sec = r.section || '(No section)';
            if (!groups.has(sec)) groups.set(sec, []);
            groups.get(sec).push([r, i]);
        }

        for (const [sec, arr] of groups) {
            const group = el('div', 'section-group');
            const head = el('div', 'section-header', `
                                                <span class="chev">▾</span>
                                                <strong>${esc(sec)}</strong>
                                                <span class="count">${arr.length} ${arr.length === 1 ? 'rule' : 'rules'}</span>
                                                `);
            const isCollapsed = folded.has(sec);
            if (isCollapsed) { group.classList.add('collapsed'); head.querySelector('.chev').style.transform = 'rotate(-90deg)'; }

            head.onclick = () => {
                const collapsed = group.classList.toggle('collapsed');
                head.querySelector('.chev').style.transform = collapsed ? 'rotate(-90deg)' : '';
                if (collapsed) folded.add(sec); else folded.delete(sec);
            };

            group.append(head);

            for (const [r] of arr) {
                const id = r.__id; // set by ensureRuleIds()
                const row = el('div', 'rule-row', `
                                                <input type="checkbox" class="rule-select" data-id="${id}" ${UI.selected.has(id) ? 'checked' : ''}>
                                                <div class="mono">${esc(r.cell || '(range)')}</div>
                                                <div class="mono">${esc(r.type || '')}</div>
                                                <div>${esc(r.note || '')}</div>
                                            `);
                row.querySelector('.rule-select').onchange = (ev) => {
                    if (ev.target.checked) UI.selected.add(id); else UI.selected.delete(id);
                };
                group.append(row);
            }

            host.append(group);
        }
        return;
    }

    // flat fallback
    for (const [r] of rows) {
        const id = r.__id;
        const row = el('div', 'rule-row', `
                                        <input type="checkbox" class="rule-select" data-id="${id}" ${UI.selected.has(id) ? 'checked' : ''}>
                                        <div class="mono">${esc(r.cell || '(range)')}</div>
                                        <div class="mono">${esc(r.type || '')}</div>
                                        <div>${esc(r.note || '')}</div>
                                    `);
        row.querySelector('.rule-select').onchange = (ev) => {
            if (ev.target.checked) UI.selected.add(id); else UI.selected.delete(id);
        };
        host.append(row);
    }
}



// -------- Right rail: per-sheet expandable lists of rules

/** Render the expandable right-rail listing of all sheets and their rules. */
function renderRailRules() {
    const host = document.getElementById('railRules');
    if (!host) return;

    const openBefore = new Set(UI.railOpen instanceof Set ? UI.railOpen : (UI.railOpen = new Set()));
    host.innerHTML = '';

    for (const [name, spec] of Object.entries(rubric.sheets || {})) {
        const box = el('div', 'rail-sheet');
        const head = el('div', 'rail-head', `
                                            <span class="chev">${openBefore.has(name) ? '▾' : '▸'}</span>
                                            <strong>${esc(name)}</strong>
                                            <span class="count">${(spec.checks || []).length} ${((spec.checks || []).length === 1 ? 'rule' : 'rules')}</span>
                                        `);
        const body = el('div', 'rail-body');

        if (!openBefore.has(name)) box.classList.add('collapsed');

        head.onclick = () => {
            const wasOpen = UI.railOpen.has(name);
            if (wasOpen) UI.railOpen.delete(name); else UI.railOpen.add(name);
            box.classList.toggle('collapsed', wasOpen);
            head.querySelector('.chev').textContent = wasOpen ? '▸' : '▾';
        };

        renderRulesList(name, spec, body);
        box.append(head, body);
        host.append(box);
    }
}


// ---------- render
/** Rebuild the main builder UI: sheets, compact lists, JSON preview, right rail, etc. */
function render() {
    UI.selected = UI.selected instanceof Set ? UI.selected : new Set();

    ensureRuleIds();

    $('#optPassFail').checked = !(rubric.report && rubric.report.include_pass_fail_column === false);
    $('#optComments').checked = !(rubric.report && rubric.report.include_comments === false);
    $('#totalPoints').value = rubric.points ?? '';
    const host = $('#sheets'); host.innerHTML = '';
    Object.entries(rubric.sheets).forEach(([name, spec]) => host.appendChild(sheetBlock(name, spec)));
    refreshJsonPreview();
    refreshMovePicker();
    renderRailRules();
}

// Toolbar events
const elGroup = document.getElementById('groupBySection');
if (elGroup) elGroup.addEventListener('change', (e) => {
    UI.groupBySection = !!e.target.checked;
    render();
});

const elSort = document.getElementById('sortBy');
if (elSort) {
    elSort.innerHTML = '';
    [
        ['section-cell', 'Section ▸ Cell'],
        ['cell', 'Cell (A→Z)'],
        ['type', 'Type'],
        ['section-type-cell', 'Section ▸ Type ▸ Cell']
    ].forEach(([val, label]) => {
        const opt = document.createElement('option');
        opt.value = val;
        opt.textContent = label;
        elSort.appendChild(opt);
    });

    // show the current setting
    elSort.value = UI.sortBy || 'section-cell';

    elSort.addEventListener('change', (e) => {
        UI.sortBy = e.target.value;
        render();
    });
}


const elMove = document.getElementById('bulkMoveBtn');
if (elMove) elMove.addEventListener('click', () => {
    if (UI.selected.size === 0) { alert('Select at least one rule.'); return; }
    const sec = prompt('Move selected rules to section:', '');
    if (sec === null) return;

    for (const key of UI.selected) {
        const [sheet, idxStr] = key.split(':');
        const spec = rubric.sheets[sheet];
        const i = parseInt(idxStr, 10);
        if (spec && spec.checks[i]) spec.checks[i].section = sec || null;
    }
    render();
});

// Expand/Collapse all sections for the currently hovered sheet (UI.activeSheet)
document.getElementById('btnExpandSections')?.addEventListener('click', () => {
    expandAllSections(getActiveSheetName(), true);
});
document.getElementById('btnCollapseSections')?.addEventListener('click', () => {
    expandAllSections(getActiveSheetName(), false);
});

// ---- collect unique section names across all rules

/** Collect a unique, sorted list of all section names across all sheets.
 * @returns {string[]}
 */
function allSectionNames() {
    const set = new Set();
    for (const sheet of Object.values(rubric.sheets || {})) {
        for (const r of (sheet.checks || [])) {
            const s = (r.section || '').trim();
            if (s) set.add(s);
        }
    }
    return [...set].sort((a, b) => a.localeCompare(b));
}

// ---- (re)populate the section picker

/** Re-populate the "move to section" <select> with known sections (and keep prior selection if possible). */
function refreshMovePicker() {
    const sel = document.getElementById('moveToSection');
    if (!sel) return;
    const keep = sel.value;
    sel.innerHTML = '';
    const opt0 = new Option('Move selected to section…', '', true, false);
    const optNew = new Option('+ New section…', '__new__', false, false);
    sel.append(opt0, optNew);
    for (const s of allSectionNames()) sel.append(new Option(s, s, false, false));
    // restore previous choice if still present
    if ([...sel.options].some(o => o.value === keep)) sel.value = keep;
}

// ---- locate the selected rules (checkboxes live next to each rule row)
/** Find all currently checked rule checkboxes that include a data-ref like "Sheet|Index".
 * @returns {Array<{sheetName:string, idx:number}>}
 */
function getSelectedRuleRefs() {
    // Works with any checkbox that has data-ref like "SheetName|12"
    // If your checkboxes don’t have data-ref yet, see comment below.
    const boxes = document.querySelectorAll('input[type="checkbox"][data-ref].rule-select:checked');
    const refs = [];
    boxes.forEach(b => {
        const [sheetName, idxStr] = String(b.dataset.ref || '').split('|');
        const idx = Number(idxStr);
        if (sheetName && Number.isInteger(idx)) refs.push({ sheetName, idx });
    });
    return refs;
}

// ---- perform the move
async function moveSelectedToSection() {
    const sel = document.getElementById('moveToSection');
    if (!sel) return;

    let target = sel.value;
    if (!target) return; // no-op
    if (target === '__new__') {
        const name = prompt('New section name:', '');
        if (!name) return;
        target = name.trim();
    }

    const refs = getSelectedRuleRefs();
    if (refs.length === 0) { alert('No rules selected.'); return; }

    for (const { sheetName, idx } of refs) {
        const sheet = rubric.sheets?.[sheetName];
        if (!sheet || !sheet.checks || !sheet.checks[idx]) continue;
        sheet.checks[idx].section = target;    // <- move
    }

    render();                 // redraw builder
    refreshMovePicker();      // keep picker updated
    clearSelection();         // optional
}

/** Uncheck all .rule-select checkboxes in the DOM. */

function clearSelection() {
    document.querySelectorAll('input[type="checkbox"].rule-select:checked')
        .forEach(b => (b.checked = false));
}

/** Ensure obj[key] is an array (create if missing) and return it.
* @template T
* @param {Object} obj
* @param {string} key
* @returns {T[]}
*/
function ensureArray(obj, key) {
    obj[key] = Array.isArray(obj[key]) ? obj[key] : [];
    return obj[key];
}

/** Register a section name in global meta.sectionOrder and sheet.section_order if missing.
* @param {string} name - Section label to register.
* @param {string} sheetName - Sheet to register per-sheet order on.
*/
function registerSection(name, sheetName) {
    // global order
    rubric.meta = rubric.meta || {};
    const globalOrder = ensureArray(rubric.meta, 'sectionOrder');
    if (!globalOrder.includes(name)) globalOrder.push(name);

    // per-sheet order
    const sheet = rubric.sheets?.[sheetName];
    if (sheet) {
        sheet.section_order = Array.isArray(sheet.section_order) ? sheet.section_order : [];
        if (!sheet.section_order.includes(name)) sheet.section_order.push(name);
    }
}

// ---- hook buttons
document.getElementById('btnMoveToSection').onclick = () => {
    const sel = Array.from(UI.selected || []);
    if (!sel.length) return;

    let dest = document.getElementById('moveToSection').value;
    if (!dest) return;
    if (dest === '__new__') {
        const name = prompt('Name for new section:');
        if (!name) return;
        dest = name.trim();
    }

    const touchedSheets = new Set();

    for (const [sheetName, spec] of Object.entries(rubric.sheets || {})) {
        const checks = spec.checks || [];
        for (const r of checks) {
            if (UI.selected.has(r.__id)) {
                r.section = dest;
                touchedSheets.add(sheetName);
            }
        }
    }

    // ensure the section is listed for ordering
    for (const sName of touchedSheets) registerSection(dest, sName);

    UI.selected.clear();
    render();              // refresh main + rail
    refreshMovePicker();   // update the section picker
    refreshJsonPreview();  // preview JSON reflects new orders
};


document.getElementById('btnClearSelection').onclick = () => {
    UI.selected.clear();
    render();
};


document.getElementById('rubricFile').addEventListener('change', async (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    try {
        const text = await f.text();
        const obj = JSON.parse(text);
        rubric = obj;
        render();
        refreshJsonPreview();
    } catch (err) {
        alert('Invalid rubric JSON: ' + (err?.message || err));
    }
});

/** Download the current rubric JSON using the optional filename input as a base. */

function downloadRubricJSON() {
    const base = (document.getElementById('jsonFilename')?.value.trim() || 'rubric')
        .replace(/\.(json)?$/i, '');
    const name = base + '.json';
    const blob = new Blob([JSON.stringify(collectRubric(), null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(a.href);
}

/** Build a full UI block for a sheet: header, compact list, rule editors, add/dup/remove handlers.
* @param {string} name - Sheet name.
* @param {{checks:object[]}} spec - Sheet spec from rubric (mutated by UI).
* @returns {HTMLElement}
*/
function sheetBlock(name, spec) {
    const wrap = el('div', 'sheet');

    // mark the DOM node and remember which sheet the mouse is over
    wrap.dataset.name = name;
    wrap.addEventListener('mouseenter', () => { UI.activeSheet = name; });

    const header = el('div', 'hstack');
    const nameInput = el('input'); nameInput.type = 'text'; nameInput.value = name; nameInput.style.width = '260px';
    const btnDup = el('button', 'ghost', 'Duplicate'); const btnDel = el('button', 'ghost danger', 'Remove');
    header.append(el('div', '', '<span class="tiny">Sheet name</span><br/>'), nameInput, btnDup, btnDel);
    wrap.append(header);

    // COMPACT (grouped) LIST just above the full editors
    const compactHost = el('div', 'stack');
    wrap.append(compactHost);
    renderRulesList(name, spec, compactHost);

    const list = el('div', 'stack');
    (spec.checks || []).forEach((r, i) => list.append(ruleBlock(name, spec, r, i)));
    wrap.append(list);

    const btnAdd = el('button', 'ghost', '+ Add Rule'); btnAdd.onclick = () => { spec.checks.push(emptyRule('formula')); render(); }; wrap.append(btnAdd);
    nameInput.onchange = () => { const newName = (nameInput.value || 'Sheet').trim(); if (newName !== name) { rubric.sheets[newName] = spec; delete rubric.sheets[name]; render(); } };
    btnDel.onclick = () => { delete rubric.sheets[name]; render(); };
    btnDup.onclick = () => { let base = name + ' Copy', i = 1, c = base; while (rubric.sheets[c]) c = `${base} ${++i}`; rubric.sheets[c] = JSON.parse(JSON.stringify(spec)); render(); };
    return wrap;
}

/** Build the editor UI for a single rule and wire up all interaction handlers.
* @param {string} sheetName
* @param {{checks:object[]}} spec
* @param {object} r - Rule object (mutated).
* @param {number} idx - Rule index in the sheet.
* @returns {HTMLElement}
*/
function ruleBlock(sheetName, spec, r, idx) {
    const b = el('div', 'rule');
    const row1 = el('div', 'row');

    // ensure this rule has a stable id
    if (!r.__id) { UI.ruleSeq = (UI.ruleSeq || 1); r.__id = `r${UI.ruleSeq++}`; }

    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.className = 'rule-select';
    // use the same id system as the compact list / rail
    cb.dataset.id = r.__id;
    cb.checked = UI.selected.has(r.__id);
    cb.onchange = (e) => {
        if (e.target.checked) UI.selected.add(r.__id);
        else UI.selected.delete(r.__id);
    };

    // put it at the very left of the row
    const cbWrap = document.createElement('div');
    cbWrap.style.display = 'flex';
    cbWrap.style.alignItems = 'center';
    cbWrap.append(cb);

    // ensure it appears first in row1
    row1.append(cbWrap);

    const typeSel = el('select');
    RULE_TYPES.forEach(t => {
        const o = el('option'); o.value = t; o.textContent = t;
        if (t === r.type) o.selected = true;
        typeSel.append(o);
    });
    const pts = el('input'); pts.type = 'number'; pts.step = '0.25'; pts.min = '0'; pts.value = r.points ?? 1;
    const cell = el('input'); cell.placeholder = 'e.g. B12'; cell.value = r.cell ?? '';
    const range = el('input'); range.placeholder = 'e.g. A2:B11'; range.value = r.range ?? '';
    const note = el('input'); note.placeholder = 'Note'; note.value = r.note ?? '';

    typeSel.style.minWidth = '160px';
    pts.style.width = '84px';
    cell.style.width = '140px';
    range.style.width = '220px';
    note.style.width = '240px';

    const section = el('input');
    section.placeholder = 'Section (optional, e.g., SUMIFS)';
    section.value = r.section ?? '';
    section.onchange = () => { r.section = (section.value || '').trim() || null; refreshJsonPreview(); };
    const labelledSection = labelled(section, 'Section');
    labelledSection.classList.add('section');

    const btnUp = el('button', 'ghost btn-icon', '▲');
    const btnDown = el('button', 'ghost btn-icon', '▼');
    btnUp.title = 'Move up'; btnDown.title = 'Move down';
    const arrows = el('div', 'arrows hstack');
    arrows.style.flexDirection = 'column';
    arrows.style.gap = '6px';
    arrows.append(btnUp, btnDown);

    btnUp.onclick = () => {
        if (idx > 0) { [spec.checks[idx - 1], spec.checks[idx]] = [spec.checks[idx], spec.checks[idx - 1]]; render(); }
    };
    btnDown.onclick = () => {
        if (idx < spec.checks.length - 1) { [spec.checks[idx + 1], spec.checks[idx]] = [spec.checks[idx], spec.checks[idx + 1]]; render(); }
    };

    row1.append(
        arrows,
        labelledSection,
        labelled(typeSel, 'Type'),
        labelled(pts, 'Pts'),
        labelled(cell, 'Cell'),
        labelled(range, 'Range')
    );
    b.append(row1);

    const row2 = el('div', 'row');
    row2.append(labelled(note, 'Note'));
    b.append(row2);

    const exf = el('input'); exf.placeholder = 'Expected formula (=A1+...)'; exf.style.width = '220px'; exf.value = r.expected_formula ?? '';
    const regex = el('input'); regex.placeholder = 'Regex (optional)'; regex.style.width = '200px'; regex.value = r.expected_formula_regex ?? '';
    const fromKey = checkbox('From key', !!r.expected_from_key);
    const absReq = checkbox('Require absolute ($)', !!r.require_absolute);
    const numFmt = el('input'); numFmt.placeholder = 'Number format (e.g. 0.00)'; numFmt.style.width = '200px'; numFmt.value = (r.format && (r.format.number_format || r.format.NumberFormat)) || '';
    const start = el('input'); start.type = 'number'; start.placeholder = 'Start'; start.style.width = '90px'; start.value = r.start ?? '';
    const step = el('input'); step.type = 'number'; step.placeholder = 'Step'; step.style.width = '90px'; step.value = r.step ?? '';
    const btnDup = el('button', 'ghost', 'Duplicate'); const btnDel = el('button', 'ghost danger', 'Remove');

    // NEW: Pivot inputs
    const pvSheet = el('input'); pvSheet.placeholder = 'Sheet (optional)';
    const pvName = el('input'); pvName.placeholder = 'Pivot name contains (optional)';
    pvSheet.style.width = '200px'; pvName.style.width = '260px';

    const pvRows = el('input'); pvRows.placeholder = 'Rows, comma-separated';
    const pvCols = el('input'); pvCols.placeholder = 'Columns, comma-separated';
    const pvFil = el('input'); pvFil.placeholder = 'Filters, comma-separated';
    pvRows.style.width = pvCols.style.width = pvFil.style.width = '260px';

    const pvValuesUI = pivotValuesBlock(r, refreshJsonPreview);

    const FMT_KIND_ORDER = [
        'general', 'number', 'currency', 'accounting', 'percent',
        'short_date', 'long_date', 'time', 'datetime', 'scientific', 'fraction', 'text', 'custom'
    ];

    function fmtPatternFrom(kind, decimals) {
        const d = Math.max(0, Number(decimals || 0));
        const dot = d > 0 ? '.' + '0'.repeat(d) : '';

        switch (kind) {
            case 'general': return 'General';
            case 'number': return '#,##0' + dot;
            case 'currency': return '$#,##0' + dot;
            case 'accounting': return '_($* #,##0' + dot + '_);_($* (#,##0' + dot + '_);_($* "-"??_);_(@_)';
            case 'percent': return '0' + dot + '%';
            case 'short_date': return 'm/d/yy';
            case 'long_date': return 'mmmm d, yyyy';
            case 'time': return 'h:mm AM/PM';
            case 'datetime': return 'm/d/yy h:mm AM/PM';
            case 'scientific': return '0' + (d > 0 ? ('.' + '0'.repeat(d)) : '.0') + 'E+00';
            case 'fraction': return '# ?/?';
            case 'text': return '@';
            case 'custom': return '';
            default: return '';
        }
    }

    function guessKindAndDecimals(fmt) {
        const raw = (fmt || '').trim();
        const first = raw.split(';')[0];
        const lower = first.toLowerCase();

        const decs = (() => {
            const m = /(?:^|;)[^;]*\.(0+)/.exec(first);
            return m ? m[1].length : null;
        })();

        if (!raw || lower === 'general') return { kind: 'general', decimals: null, raw };
        if (lower.includes('%')) return { kind: 'percent', decimals: decs ?? 0, raw };
        if (lower.includes('_(') || lower.includes('* ') || (lower.includes('$') && lower.includes('_')))
            return { kind: 'accounting', decimals: decs ?? 0, raw };
        if (/\$|€|£|¥|₩/.test(first)) return { kind: 'currency', decimals: decs ?? 2, raw };
        if (lower.includes('e+')) return { kind: 'scientific', decimals: decs ?? 1, raw };
        if (lower.includes('@')) return { kind: 'text', decimals: null, raw };
        if (lower.includes('?/')) return { kind: 'fraction', decimals: null, raw };
        const hasD = /d/.test(lower), hasM = /m/.test(lower), hasY = /y/.test(lower),
            hasH = /h/.test(lower), hasS = /s/.test(lower), hasAM = lower.includes('am/pm');
        if ((hasD || hasM || hasY) && (hasH || hasS || hasAM)) return { kind: 'datetime', decimals: null, raw };
        if (hasH || hasS || hasAM) return { kind: 'time', decimals: null, raw };
        if (/(mmm|dddd)/i.test(first)) return { kind: 'long_date', decimals: null, raw };
        if (hasD || hasM || hasY) return { kind: 'short_date', decimals: null, raw };
        if (/[#0]/.test(first)) return { kind: 'number', decimals: decs ?? 0, raw };
        return { kind: 'custom', decimals: null, raw };
    }

    function showFor(type) {
        row2.innerHTML = '';

        function hideCellAndRange() {
            const hide = inp => inp?.closest('.labelled')?.classList.add('hidden');
            hide(b.querySelector('input[placeholder="e.g. B12"]'));
            hide(b.querySelector('input[placeholder="e.g. A2:B11"]'));
        }

        function hideSeriesControls() {
            const seriesLabel = Array.from(row2.querySelectorAll('.tiny, .muted, div'))
                .find(n => /Series \(categories\/values/i.test(n.textContent || ''));
            if (seriesLabel) {
                let p = seriesLabel.parentElement;
                while (p && p !== row2) {
                    if (p.classList.contains('stack') || p.classList.contains('row')) {
                        p.classList.add('hidden');
                        break;
                    }
                    p = p.parentElement;
                }
            }
        }

        if (type === 'formula') {
            row2.append(labelled(exf, 'Expected formula'), labelled(regex, 'Regex'), fromKey.wrap, absReq.wrap);

            const expVal = el('input');
            expVal.placeholder = 'e.g. 0.2154 or 21.54%';
            expVal.style.width = '220px';
            expVal.value = (r.expected ?? '') + '';
            expVal.onchange = () => {
                const v = (expVal.value ?? '').trim();
                r.expected = v === '' ? null : v;
                refreshJsonPreview();
            };

            const tol = el('input');
            tol.type = 'number';
            tol.step = 'any';
            tol.placeholder = 'e.g. 0.0001';
            tol.style.width = '140px';
            tol.value = (r.tolerance ?? '');
            tol.onchange = () => {
                const v = (tol.value ?? '').trim();
                r.tolerance = v === '' ? null : +v;
                refreshJsonPreview();
            };

            row2.append(labelled(expVal, 'Expected value'), labelled(tol, 'Tolerance (abs)'));

            r.any_of = Array.isArray(r.any_of) ? r.any_of : [];

            const list = el('div', 'stack');
            function renderAnyOfFormula() {
                list.innerHTML = '';
                (r.any_of || []).forEach((opt, i) => {
                    const row = el('div', 'hstack');

                    const altF = el('input');
                    altF.placeholder = 'Alternative formula (=...)';
                    altF.style.width = '220px';
                    altF.value = opt.expected_formula ?? '';
                    altF.onchange = () => {
                        r.any_of[i].expected_formula = altF.value.trim() || null;
                        if ((r.any_of || []).length > 0) {
                            r.expected_formula = null;
                            exf.value = '';
                        }
                        refreshJsonPreview();
                    };

                    const altRe = el('input');
                    altRe.placeholder = 'Regex (optional)';
                    altRe.style.width = '200px';
                    altRe.value = opt.expected_formula_regex ?? '';
                    altRe.onchange = () => {
                        r.any_of[i].expected_formula_regex = altRe.value.trim() || null;
                        refreshJsonPreview();
                    };

                    const del = el('button', 'ghost danger', 'Remove');
                    del.onclick = () => {
                        r.any_of.splice(i, 1);
                        if (r.any_of.length === 0) {
                            r.expected_formula = (exf.value.trim() || null);
                        } else {
                            r.expected_formula = null;
                            exf.value = '';
                        }
                        renderAnyOfFormula();
                        refreshJsonPreview();
                    };

                    row.append(labelled(altF, 'Formula'), labelled(altRe, 'Regex'), del);
                    list.append(row);
                });
            }
            renderAnyOfFormula();

            const addAltF = el('button', 'ghost', 'Add formula');
            addAltF.onclick = () => {
                (r.any_of = r.any_of || []).push({ expected_formula: null, expected_formula_regex: null });
                r.expected_formula = null;
                exf.value = '';
                renderAnyOfFormula();
                refreshJsonPreview();
            };

            const anyWrap = el('div', 'stack');
            anyWrap.append(el('div', 'tiny muted', 'Any of these formulas (optional):'), list, addAltF);

            row2.append(anyWrap, btnDup, btnDel);
        } else if (type === 'value') {
            const cast = v => (/^-?\d+(?:\.\d+)?$/.test(v) ? +v : v);
            const clean = s => (s ?? '').trim();

            r.any_of = Array.isArray(r.any_of) ? r.any_of : [];

            const mainExp = el('input');
            mainExp.placeholder = 'Expected (single value)';
            mainExp.style.width = '260px';

            if ((r.any_of || []).length > 0) {
                mainExp.value = '';
                r.expected = null;
            } else {
                mainExp.value = r.expected != null ? (r.expected + '') : '';
            }

            const csRule = document.createElement('input');
            csRule.type = 'checkbox';
            csRule.checked = !!r.case_sensitive;
            csRule.onchange = () => { r.case_sensitive = csRule.checked; refreshJsonPreview(); };

            mainExp.onchange = () => {
                if ((r.any_of || []).length > 0) {
                    mainExp.value = '';
                    r.expected = null;
                } else {
                    const v = clean(mainExp.value);
                    r.expected = v === '' ? null : cast(v);
                }
                refreshJsonPreview();
            };

            const list = el('div', 'stack');

            function renderAnyOf() {
                list.innerHTML = '';
                (r.any_of || []).forEach((opt, i) => {
                    const row = el('div', 'hstack');

                    const exp = el('input');
                    exp.placeholder = 'Alternative expected value';
                    exp.style.width = '220px';
                    exp.value = opt.expected != null ? (opt.expected + '') : '';
                    exp.onchange = () => {
                        const v = clean(exp.value);
                        r.any_of[i].expected = v === '' ? null : cast(v);
                        r.expected = null;
                        mainExp.value = '';
                        refreshJsonPreview();
                    };

                    const cs = document.createElement('input');
                    cs.type = 'checkbox';
                    cs.checked = !!opt.case_sensitive;
                    cs.onchange = () => { r.any_of[i].case_sensitive = cs.checked; refreshJsonPreview(); };

                    const del = el('button', 'ghost danger', 'Remove');
                    del.onclick = () => {
                        r.any_of.splice(i, 1);
                        if (r.any_of.length === 0) {
                            r.expected = clean(mainExp.value) ? cast(mainExp.value) : null;
                        } else {
                            r.expected = null;
                            mainExp.value = '';
                        }
                        renderAnyOf();
                        refreshJsonPreview();
                    };

                    row.append(labelled(exp, 'Answer'), labelled(cs, 'Case sensitive'), del);
                    list.append(row);
                });
            }

            const addAlt = el('button', 'ghost', 'Add answer');
            addAlt.onclick = () => {
                (r.any_of = r.any_of || []).push({ expected: null });
                r.expected = null;
                mainExp.value = '';
                renderAnyOf();
                refreshJsonPreview();
            };

            const btnDup = el('button', 'ghost', 'Duplicate');
            const btnDel = el('button', 'ghost danger', 'Remove');
            btnDup.onclick = () => { spec.checks.splice(idx + 1, 0, JSON.parse(JSON.stringify(r))); render(); };
            btnDel.onclick = () => { spec.checks.splice(idx, 1); render(); };

            const anyWrap = el('div', 'stack');
            anyWrap.append(
                el('div', 'tiny muted', 'Any of these answers (optional):'),
                list,
                addAlt
            );

            row2.append(
                labelled(mainExp, 'Expected (single)'),
                labelled(csRule, 'Case sensitive'),
                anyWrap,
                btnDup,
                btnDel
            );

            renderAnyOf();
            refreshJsonPreview();
        }

        else if (type === 'format' || type === 'range_format') {
            r.format = r.format || {};
            const getFmt = () => r.format.number_format || '';
            const setFmt = v => { r.format.number_format = (v || '').trim() || null; refreshJsonPreview(); };

            const kindSel = document.createElement('select');
            [
                ['general', 'General'],
                ['number', 'Number'],
                ['currency', 'Currency'],
                ['accounting', 'Accounting'],
                ['percent', 'Percent'],
                ['short_date', 'Short date'],
                ['long_date', 'Long date'],
                ['time', 'Time'],
                ['datetime', 'Date & time'],
                ['scientific', 'Scientific'],
                ['fraction', 'Fraction'],
                ['text', 'Text'],
                ['custom', 'Custom…']
            ].forEach(([v, l]) => { const o = document.createElement('option'); o.value = v; o.textContent = l; kindSel.append(o); });

            const decWrap = document.createElement('div'); decWrap.className = 'stack';
            const decInp = document.createElement('input');
            decInp.type = 'number'; decInp.min = '0'; decInp.step = '1'; decInp.value = '2';
            decWrap.append(decInp);

            const numFmt = document.createElement('input');
            numFmt.placeholder = 'e.g., $#,##0.00 or 0% or m/d/yy';
            numFmt.style.width = '320px';

            const g = guessKindAndDecimals(getFmt());
            kindSel.value = g.kind;
            if (g.decimals != null) decInp.value = String(g.decimals);
            numFmt.value = g.raw || fmtPatternFrom(kindSel.value, decInp.value);

            const decimalsRelevant = k => ['number', 'currency', 'accounting', 'percent', 'scientific'].includes(k);

            function refreshVisibility() {
                decWrap.style.display = decimalsRelevant(kindSel.value) ? '' : 'none';
                if (kindSel.value !== 'custom') {
                    numFmt.value = fmtPatternFrom(kindSel.value, decInp.value);
                    setFmt(numFmt.value);
                }
            }

            kindSel.onchange = refreshVisibility;
            decInp.onchange = refreshVisibility;
            numFmt.onchange = () => {
                const g2 = guessKindAndDecimals(numFmt.value);
                setFmt(numFmt.value);
                if (g2.kind && g2.kind !== kindSel.value) kindSel.value = g2.kind;
                if (g2.decimals != null) decInp.value = String(g2.decimals);
                refreshVisibility();
            };

            refreshVisibility();

            r.format.font = r.format.font || {};
            const initBold = (r.format.font.bold ?? r.format.font_bold) ?? null;
            const initSize = (r.format.font.size ?? r.format.font_size) ?? null;

            const boldCk = document.createElement('input');
            boldCk.type = 'checkbox';
            boldCk.checked = !!initBold;
            boldCk.onchange = () => {
                const v = boldCk.checked ? true : null;
                r.format.font.bold = v;
                r.format.font_bold = v;
                refreshJsonPreview();
            };

            const sizeInp = document.createElement('input');
            sizeInp.type = 'number';
            sizeInp.min = '6'; sizeInp.max = '96'; sizeInp.step = '0.5';
            sizeInp.placeholder = 'e.g., 12';
            sizeInp.style.width = '120px';
            sizeInp.value = (initSize ?? '').toString();
            sizeInp.onchange = () => {
                const t = sizeInp.value.trim();
                const v = t === '' ? null : Number(t);
                r.format.font.size = v;
                r.format.font_size = v;
                refreshJsonPreview();
            };

            const rowA = el('div', 'row');
            rowA.style.gridTemplateColumns = '1fr 0.5fr 1.2fr';
            rowA.style.gap = '12px';
            rowA.append(
                labelled(kindSel, 'Kind'),
                labelled(decWrap, 'Decimals'),
                labelled(numFmt, 'Number format (A1 pattern)')
            );

            const rowB = el('div', 'row');
            rowB.style.gridTemplateColumns = '0.6fr 0.8fr';
            rowB.style.gap = '12px';

            const boldWrap = document.createElement('label');
            boldWrap.className = 'checkbox';
            boldWrap.append(boldCk, el('span', '', 'Bold'));

            rowB.append(boldWrap, labelled(sizeInp, 'Font size (pt)'));

            row2.append(rowA, rowB);

            const btnDup = el('button', 'ghost', 'Duplicate');
            const btnDel = el('button', 'ghost danger', 'Remove');
            btnDup.onclick = () => { spec.checks.splice(idx + 1, 0, JSON.parse(JSON.stringify(r))); render(); };
            btnDel.onclick = () => { spec.checks.splice(idx, 1); render(); };
            row2.append(btnDup, btnDel);
        } else if (type === 'range_sequence') { row2.append(labelled(start, 'Start'), labelled(step, 'Step'), btnDup, btnDel); }
        else if (type === 'pivot_layout') {
            const pv = ensurePivot(r);
            pvSheet.value = pv.sheet ?? '';
            pvName.value = pv.tableNameLike ?? '';
            pvRows.value = joinCSV(pv.rows);
            pvCols.value = joinCSV(pv.columns);
            pvFil.value = joinCSV(pv.filters);

            pvSheet.onchange = () => { ensurePivot(r).sheet = pvSheet.value.trim() || null; refreshJsonPreview(); };
            pvName.onchange = () => { ensurePivot(r).tableNameLike = pvName.value.trim() || null; refreshJsonPreview(); };
            pvRows.onchange = () => { ensurePivot(r).rows = splitCSV(pvRows.value); refreshJsonPreview(); };
            pvCols.onchange = () => { ensurePivot(r).columns = splitCSV(pvCols.value); refreshJsonPreview(); };
            pvFil.onchange = () => { ensurePivot(r).filters = splitCSV(pvFil.value); refreshJsonPreview(); };

            const left = el('div', 'stack');
            left.append(labelled(pvSheet, 'Sheet'), labelled(pvName, 'Name contains'), labelled(pvRows, 'Rows'), labelled(pvCols, 'Columns'), labelled(pvFil, 'Filters'));

            const right = el('div', 'stack');
            right.append(el('div', 'tiny muted', 'Values (Field + Aggregation)'), pvValuesUI);

            const grid = el('div', 'row');
            grid.style.gridTemplateColumns = '1fr 1fr';
            grid.style.gap = '12px';
            grid.append(left, right);

            row2.append(grid);
            const btnDup = el('button', 'ghost', 'Duplicate'); const btnDel = el('button', 'ghost danger', 'Remove');
            btnDup.onclick = () => { spec.checks.splice(idx + 1, 0, JSON.parse(JSON.stringify(r))); render(); };
            btnDel.onclick = () => { spec.checks.splice(idx, 1); render(); };
            row2.append(btnDup, btnDel);
        } else if (type === 'conditional_format') {
            const mkLabeled = (labelText, control) => {
                const w = document.createElement('div');
                w.className = 'stack';
                const lab = document.createElement('div');
                lab.className = 'tiny muted';
                lab.textContent = labelText;
                w.append(lab, control);
                return w;
            };

            const sheet = document.createElement('input'); sheet.placeholder = 'Sheet (optional)';
            const range = document.createElement('input'); range.placeholder = 'Range (e.g., B2:B50)';

            const typ = document.createElement('select');
            [['cellIs', 'Cell is…'], ['expression', 'Formula (TRUE/FALSE)'], ['containsText', 'Contains text'], ['top10', 'Top/bottom N'], ['dataBar', 'Data bar'], ['colorScale', 'Color scale'], ['iconSet', 'Icon set']]
                .forEach(([v, l]) => { const o = document.createElement('option'); o.value = v; o.textContent = l; typ.append(o); });

            const op = document.createElement('select');
            [['gt', '> greater than'], ['ge', '≥ greater than or equal'], ['lt', '< less than'], ['le', '≤ less than or equal'], ['eq', '= equal to'], ['ne', '≠ not equal to'], ['between', 'between (inclusive)'], ['notBetween', 'not between']]
                .forEach(([v, l]) => { const o = document.createElement('option'); o.value = v; o.textContent = l; op.append(o); });

            const f1 = document.createElement('input'); f1.placeholder = 'Value / Formula 1 (e.g., 85 or =B2>0)';
            const f2 = document.createElement('input'); f2.placeholder = 'Value / Formula 2 (for between)';
            const txt = document.createElement('input'); txt.placeholder = 'Text to look for';
            const fill = document.createElement('input'); fill.placeholder = 'Fill RGB (optional, e.g., FFFF00)';

            r.cond = r.cond || {};
            const c = r.cond;
            sheet.value = c.sheet ?? '';
            range.value = c.range ?? '';
            typ.value = c.type ?? 'cellIs';
            op.value = c.op ?? 'gt';
            f1.value = c.formula1 ?? '';
            f2.value = c.formula2 ?? '';
            txt.value = c.text ?? '';
            fill.value = c.fillRgb ?? '';

            if (!op.value || ![...op.options].some(o => o.value === op.value)) {
                op.selectedIndex = 0;
            }

            const topRow = document.createElement('div');
            topRow.className = 'hstack';
            topRow.style.gap = '8px';
            topRow.style.flexWrap = 'wrap';

            const wSheet = mkLabeled('Sheet', sheet);
            const wRange = mkLabeled('Range', range);
            const wType = mkLabeled('Type', typ);
            const wOp = mkLabeled('Operator', op);

            topRow.append(wSheet, wRange, wType, wOp);
            row2.append(topRow);

            const botRow = document.createElement('div');
            botRow.className = 'row';
            botRow.style.gridTemplateColumns = '1fr 1fr 1fr 1fr';
            botRow.style.gap = '8px';

            const wF1 = mkLabeled('Value / Formula 1', f1);
            const wF2 = mkLabeled('Value / Formula 2', f2);
            const wText = mkLabeled('Text', txt);
            const wFill = mkLabeled('Fill RGB', fill);

            botRow.append(wF1, wF2, wText, wFill);
            row2.append(botRow);

            const btnDup = el('button', 'ghost', 'Duplicate');
            const btnDel = el('button', 'ghost danger', 'Remove');
            btnDup.onclick = () => { spec.checks.splice(idx + 1, 0, JSON.parse(JSON.stringify(r))); render(); };
            btnDel.onclick = () => { spec.checks.splice(idx, 1); render(); };
            row2.append(btnDup, btnDel);

            const save = () => {
                const norm = s => (s || '').trim();
                r.cond = {
                    sheet: norm(sheet.value) || null,
                    range: norm(range.value) || null,
                    type: typ.value || null,
                    op: op.value || null,
                    formula1: norm(f1.value) || null,
                    formula2: norm(f2.value) || null,
                    text: norm(txt.value) || null,
                    fillRgb: norm(fill.value) || null
                };
                refreshJsonPreview();
                toggle();
            };

            const toggle = () => {
                const t = typ.value;
                wOp.style.display = (t === 'cellIs') ? '' : 'none';
                wF2.style.display = (t === 'cellIs' && (op.value === 'between' || op.value === 'notBetween')) ? '' : 'none';
                wText.style.display = (t === 'containsText') ? '' : 'none';
                if (t === 'expression') {
                    wOp.style.display = 'none';
                    wF2.style.display = 'none';
                }
                if (t === 'dataBar' || t === 'colorScale' || t === 'iconSet' || t === 'top10') {
                    wOp.style.display = 'none';
                    wF1.style.display = 'none';
                    wF2.style.display = 'none';
                    wText.style.display = 'none';
                }
            };

            [sheet, range, typ, op, f1, f2, txt, fill].forEach(x => x.addEventListener('change', save));
            toggle();
        } else if (type === 'chart') {
            const ch = ensureChart(r);

            const cSheet = el('input'); cSheet.placeholder = 'Sheet (optional)'; cSheet.style.width = '200px'; cSheet.value = ch.sheet ?? '';
            const cName = el('input'); cName.placeholder = 'Chart name contains'; cName.style.width = '260px'; cName.value = ch.name_like ?? '';
            const cType = el('select');['', 'column', 'bar', 'line', 'pie', 'pie3D', 'scatter', 'area', 'doughnut', 'radar', 'bubble']
                .forEach(v => { const o = el('option'); o.value = v; o.textContent = v || '(any type)'; if ((ch.type || '') === v) o.selected = true; cType.append(o); });

            const cTitle = el('input'); cTitle.placeholder = 'Title text (exact)'; cTitle.value = ch.title ?? '';
            const cTitleRef = el('input'); cTitleRef.placeholder = 'Title from cell (e.g., Summary!$B$1)'; cTitleRef.value = ch.title_ref ?? '';
            const cLegend = el('select');[['', '(any)'], ['t', 'top'], ['r', 'right'], ['b', 'bottom'], ['l', 'left'], ['tr', 'top-right']]
                .forEach(([v, l]) => { const o = el('option'); o.value = v; o.textContent = l; if ((ch.legend_pos || '') === v) o.selected = true; cLegend.append(o); });
            const cLabels = document.createElement('input'); cLabels.type = 'checkbox'; cLabels.checked = ch.data_labels === true;
            const cX = el('input'); cX.placeholder = 'X-axis title'; cX.value = ch.x_title ?? '';
            const cY = el('input'); cY.placeholder = 'Y-axis title'; cY.value = ch.y_title ?? '';

            const seriesWrap = el('div', 'stack');
            function renderSeries() {
                seriesWrap.innerHTML = '';
                (ch.series || []).forEach((s, i) => {
                    const row = el('div', 'hstack');
                    const sName = el('input'); sName.placeholder = 'Series name (text)'; sName.style.width = '180px'; sName.value = s.name ?? '';
                    const sNameRef = el('input'); sNameRef.placeholder = 'Series name ref (e.g., Sheet!$B$1)'; sNameRef.style.width = '200px'; sNameRef.value = s.name_ref ?? '';
                    const sCat = el('input'); sCat.placeholder = 'Categories ref (e.g., Sheet!$A$2:$A$13)'; sCat.style.width = '260px'; sCat.value = s.cat_ref ?? '';
                    const sVal = el('input'); sVal.placeholder = 'Values ref (e.g., Sheet!$B$2:$B$13)'; sVal.style.width = '260px'; sVal.value = s.val_ref ?? '';
                    const del = el('button', 'ghost danger', 'Remove');

                    sName.onchange = () => { ch.series[i].name = sName.value.trim() || null; refreshJsonPreview(); };
                    sNameRef.onchange = () => { ch.series[i].name_ref = sNameRef.value.trim() || null; refreshJsonPreview(); };
                    sCat.onchange = () => { ch.series[i].cat_ref = sCat.value.trim() || null; refreshJsonPreview(); };
                    sVal.onchange = () => { ch.series[i].val_ref = sVal.value.trim() || null; refreshJsonPreview(); };
                    del.onclick = () => { ch.series.splice(i, 1); renderSeries(); refreshJsonPreview(); };

                    row.append(labelled(sName, 'Name'), labelled(sNameRef, 'Name ref'), labelled(sCat, 'Categories'), labelled(sVal, 'Values'), del);
                    seriesWrap.append(row);
                });
            }
            renderSeries();

            const addSeries = el('button', 'ghost', '+ Add series');
            addSeries.onclick = () => { (ch.series = ch.series || []).push({ name: null, name_ref: null, cat_ref: null, val_ref: null }); renderSeries(); refreshJsonPreview(); };

            cSheet.onchange = () => { ch.sheet = cSheet.value.trim() || null; refreshJsonPreview(); };
            cName.onchange = () => { ch.name_like = cName.value.trim() || null; refreshJsonPreview(); };
            cType.onchange = () => { ch.type = cType.value || null; refreshJsonPreview(); };
            cTitle.onchange = () => { ch.title = cTitle.value.trim() || null; refreshJsonPreview(); };
            cTitleRef.onchange = () => { ch.title_ref = cTitleRef.value.trim() || null; refreshJsonPreview(); };
            cLegend.onchange = () => { ch.legend_pos = cLegend.value || null; refreshJsonPreview(); };
            cLabels.onchange = () => { ch.data_labels = cLabels.checked ? true : null; refreshJsonPreview(); };
            cX.onchange = () => { ch.x_title = cX.value.trim() || null; refreshJsonPreview(); };
            cY.onchange = () => { ch.y_title = cY.value.trim() || null; refreshJsonPreview(); };

            const left = el('div', 'stack');
            left.append(
                labelled(cSheet, 'Sheet'),
                labelled(cName, 'Name contains'),
                labelled(cType, 'Chart type'),
                labelled(cLegend, 'Legend position'),
                labelled(cLabels, 'Data labels')
            );

            const right = el('div', 'stack');
            right.append(
                labelled(cTitle, 'Title'),
                labelled(cTitleRef, 'Title (from cell)'),
                labelled(cX, 'X-axis title'),
                labelled(cY, 'Y-axis title')
            );

            const grid = el('div', 'row');
            grid.style.gridTemplateColumns = '1fr 1fr';
            grid.style.gap = '12px';
            grid.append(left, right);

            row2.append(grid);
            b.append(row2, el('div', 'tiny muted', 'Series (categories/values required for each):'), seriesWrap, addSeries);

            const btnDup = el('button', 'ghost', 'Duplicate');
            const btnDel = el('button', 'ghost danger', 'Remove');
            btnDup.onclick = () => { spec.checks.splice(idx + 1, 0, JSON.parse(JSON.stringify(r))); render(); };
            btnDel.onclick = () => { spec.checks.splice(idx, 1); render(); };
            row2.append(btnDup, btnDel);
        } else if (type === 'table') {
            const t = ensureTable(r);

            hideCellAndRange();
            hideSeriesControls();

            const NUM_RE = /^\s*[$€£]?\s*-?\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*%?\s*$/;

            const rangeRef = el('input'); rangeRef.placeholder = 'A1:D20 or Sheet!A1:D20';
            rangeRef.value = t.range_ref || '';
            rangeRef.onchange = () => { t.range_ref = rangeRef.value.trim() || null; refreshJsonPreview(); };

            const rowsInp = el('input'); rowsInp.type = 'number'; rowsInp.min = '0';
            rowsInp.placeholder = 'rows';
            rowsInp.value = (t.rows ?? '').toString();
            rowsInp.onchange = () => { const v = rowsInp.value.trim(); t.rows = v === '' ? null : Number(v); refreshJsonPreview(); };

            const colsInp = el('input'); colsInp.type = 'number'; colsInp.min = '0';
            colsInp.placeholder = 'cols';
            colsInp.value = (t.cols ?? '').toString();
            colsInp.onchange = () => { const v = colsInp.value.trim(); t.cols = v === '' ? null : Number(v); refreshJsonPreview(); };

            const allowER = document.createElement('input'); allowER.type = 'checkbox';
            allowER.checked = t.allow_extra_rows === true;
            allowER.onchange = () => { t.allow_extra_rows = allowER.checked ? true : null; refreshJsonPreview(); };
            const allowERLbl = el('label', 'checkbox'); allowERLbl.append(allowER, el('span', '', 'Allow extra rows'));

            const allowEC = document.createElement('input'); allowEC.type = 'checkbox';
            allowEC.checked = t.allow_extra_cols === true;
            allowEC.onchange = () => { t.allow_extra_cols = allowEC.checked ? true : null; refreshJsonPreview(); };
            const allowECLbl = el('label', 'checkbox'); allowECLbl.append(allowEC, el('span', '', 'Allow extra cols'));

            const dims = el('div', 'row');
            dims.style.gridTemplateColumns = '1.2fr .6fr .6fr .8fr .9fr';
            dims.style.gap = '12px';
            dims.append(labelled(rangeRef, 'Range (optional)'), labelled(rowsInp, 'Rows'), labelled(colsInp, 'Cols'), allowERLbl, allowECLbl);
            row2.append(dims);

            const matchWhole = document.createElement('input'); matchWhole.type = 'checkbox';
            const orderMatters = document.createElement('input'); orderMatters.type = 'checkbox';
            const caseSens = document.createElement('input'); caseSens.type = 'checkbox';
            const trimCells = document.createElement('input'); trimCells.type = 'checkbox';

            matchWhole.checked = !!t.body_match;
            orderMatters.checked = !!t.body_order_matters;
            caseSens.checked = !!t.body_case_sensitive;
            trimCells.checked = (t.body_trim ?? true);

            matchWhole.onchange = () => { t.body_match = matchWhole.checked ? true : null; refreshJsonPreview(); };
            orderMatters.onchange = () => { t.body_order_matters = orderMatters.checked ? true : null; refreshJsonPreview(); };
            caseSens.onchange = () => { t.body_case_sensitive = caseSens.checked ? true : null; refreshJsonPreview(); };
            trimCells.onchange = () => { t.body_trim = trimCells.checked ? true : null; refreshJsonPreview(); };

            const toggles = el('div', 'row');
            toggles.style.gridTemplateColumns = 'auto auto auto auto';
            toggles.style.gap = '14px';
            toggles.append(
                el('label', 'checkbox', null).appendChild(matchWhole).parentElement.appendChild(el('span', '', 'Match entire contents')).parentElement,
                el('label', 'checkbox', null).appendChild(orderMatters).parentElement.appendChild(el('span', '', 'Row order matters')).parentElement,
                el('label', 'checkbox', null).appendChild(caseSens).parentElement.appendChild(el('span', '', 'Case sensitive')).parentElement,
                el('label', 'checkbox', null).appendChild(trimCells).parentElement.appendChild(el('span', '', 'Trim cell text')).parentElement
            );
            row2.append(toggles);

            const containsWrap = el('div', 'stack');
            const list = el('div', 'stack');
            function renderContains() {
                list.innerHTML = '';
                (t.contains_rows || []).forEach((obj, i) => {
                    const row = el('div', 'hstack');
                    const preview = el('div'); preview.className = 'mono small muted';
                    preview.textContent = Object.entries(obj).map(([k, v]) => `${k}="${v}"`).join('  ·  ');
                    const del = el('button', 'ghost danger', 'Remove');
                    del.onclick = () => { t.contains_rows.splice(i, 1); renderContains(); refreshJsonPreview(); };
                    row.append(preview, del);
                    list.append(row);
                });
            }
            const pasteRows = el('button', 'ghost', 'Paste required rows…');
            pasteRows.onclick = () => {
                const txt = prompt(
                    'Paste CSV/TSV with headers in first row.\n' +
                    'Only the columns you include will be matched.\n' +
                    'Example:\nRegion,Item\nEast,Widgets\nWest,Gadgets', ''
                );
                if (txt != null) {
                    const lines = txt.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
                    if (lines.length > 1) {
                        const delim = (txt.indexOf('\t') >= 0) ? '\t' : ',';
                        const headers = lines[0].split(delim).map(s => s.trim()).filter(Boolean);
                        const rows = lines.slice(1).map(line => line.split(delim));
                        t.contains_rows = rows.map(vals => {
                            const obj = {};
                            headers.forEach((h, i) => { if (vals[i] != null && vals[i].trim() !== '') obj[h] = vals[i].trim(); });
                            return obj;
                        }).filter(o => Object.keys(o).length > 0);
                        renderContains();
                        refreshJsonPreview();
                    }
                }
            };
            containsWrap.append(el('div', 'tiny muted', 'Required rows (must appear at least once):'), list, pasteRows);
            row2.append(containsWrap);
            renderContains();

            t.columns = Array.isArray(t.columns) ? t.columns.filter(Boolean) : [];
            const colList = el('div', 'stack');

            function renderColumns() {
                colList.innerHTML = '';
                (t.columns || []).forEach((name, i) => {
                    const rrow = el('div', 'hstack');

                    const inp = el('input');
                    inp.placeholder = 'Header name';
                    inp.style.width = '260px';
                    inp.value = name || '';
                    inp.onchange = () => { t.columns[i] = inp.value.trim(); refreshJsonPreview(); };

                    const up = el('button', 'ghost', '↑');
                    up.onclick = () => {
                        if (i > 0) { [t.columns[i - 1], t.columns[i]] = [t.columns[i], t.columns[i - 1]]; renderColumns(); refreshJsonPreview(); }
                    };
                    const dn = el('button', 'ghost', '↓');
                    dn.onclick = () => {
                        if (i < t.columns.length - 1) { [t.columns[i + 1], t.columns[i]] = [t.columns[i], t.columns[i + 1]]; renderColumns(); refreshJsonPreview(); }
                    };
                    const del = el('button', 'ghost danger', 'Remove');
                    del.onclick = () => { t.columns.splice(i, 1); renderColumns(); refreshJsonPreview(); };

                    rrow.append(labelled(inp, `Header ${i + 1}`), up, dn, del);
                    colList.append(rrow);
                });
            }

            const addCol = el('button', 'ghost', 'Add column');
            addCol.onclick = () => { (t.columns = t.columns || []).push(''); renderColumns(); refreshJsonPreview(); };

            const bulk = el('button', 'ghost', 'Paste…');
            bulk.onclick = () => {
                const txt = prompt('Paste header names (comma or newline separated):', '');
                if (txt != null) {
                    const parts = txt.split(/[\n,]+/).map(s => s.trim()).filter(Boolean);
                    if (parts.length) { t.columns = parts; renderColumns(); refreshJsonPreview(); }
                }
            };

            const reqOrder = document.createElement('input'); reqOrder.type = 'checkbox';
            reqOrder.checked = t.require_order === true;
            reqOrder.onchange = () => { t.require_order = reqOrder.checked; refreshJsonPreview(); };
            const reqOrderLbl = el('label', 'checkbox'); reqOrderLbl.append(reqOrder, el('span', '', 'Require header order'));

            const colsWrap = el('div', 'stack');
            const colsLabel = el('div', 'tiny muted', 'Columns (in order if "Require header order" is checked):');
            const actions = el('div', 'hstack'); actions.append(addCol, bulk);
            colsWrap.append(colsLabel, colList, actions);

            const nameLike = el('input');
            nameLike.placeholder = 'Table name contains (optional)';
            nameLike.value = t.name_like || '';
            nameLike.onchange = () => { t.name_like = nameLike.value.trim() || null; refreshJsonPreview(); };

            const left = el('div', 'stack');
            left.append(labelled(nameLike, 'Name like'), colsWrap, reqOrderLbl);

            const right = el('div', 'stack');
            const show = document.createElement('input'); show.type = 'checkbox';
            const showLbl = el('label', 'checkbox'); showLbl.append(show, el('span', '', 'Show key table data'));
            const previewPanel = el('div', 'preview-card'); previewPanel.style.display = 'none';
            right.append(showLbl, previewPanel);

            show.onchange = () => {
                previewPanel.style.display = show.checked ? '' : 'none';
                if (show.checked) renderPreview();
            };

            function renderPreview() {
                previewPanel.innerHTML = '';

                const rows = Array.isArray(t.body_rows) ? t.body_rows : null;
                const headers = Array.isArray(t.columns) ? t.columns : null;

                if (!rows || !rows.length || !headers || !headers.length) {
                    previewPanel.append(el('div', 'muted', 'No captured table data.'));
                    return;
                }

                const chipsTitle = el('div', 'preview-title', 'Distinct values by column (sample up to 12):');
                const chipsGrid = el('div', 'chips-grid');

                headers.forEach((h, ci) => {
                    const uniq = new Map();
                    rows.forEach(r => { const v = (r[ci] ?? '').toString(); uniq.set(v, (uniq.get(v) || 0) + 1); });

                    const col = el('div', 'chips-col');
                    const head = el('div', 'head', h);
                    const chips = el('div', 'chips');
                    [...uniq.keys()].slice(0, 12).forEach(v => chips.append(Object.assign(el('span', 'chip'), { textContent: v === '' ? '(blank)' : v })));
                    if (uniq.size > 12) chips.append(Object.assign(el('span', 'chip'), { textContent: `+${uniq.size - 12} more` }));
                    col.append(head, chips);
                    chipsGrid.append(col);
                });

                const N = 25;
                const gridTitle = el('div', 'preview-title', `Previewing first ${Math.min(N, rows.length)} of ${rows.length} data rows:`);
                const tbl = document.createElement('table'); tbl.className = 'mini-table';
                const thead = document.createElement('thead'); const trh = document.createElement('tr');
                headers.forEach(h => { const th = document.createElement('th'); th.textContent = h; trh.append(th); });
                thead.append(trh); tbl.append(thead);

                const tbody = document.createElement('tbody');
                rows.slice(0, N).forEach(r => {
                    const tr = document.createElement('tr');
                    headers.forEach((_, ci) => {
                        const td = document.createElement('td');
                        const text = (r[ci] ?? '').toString();
                        td.textContent = text;
                        if (NUM_RE.test(text)) td.classList.add('td-num');
                        tr.append(td);
                    });
                    tbody.append(tr);
                });
                tbl.append(tbody);

                const scroll = el('div', 'scroll-wrap'); scroll.append(tbl);

                previewPanel.append(chipsTitle, chipsGrid, gridTitle, scroll);
            }

            const grid = el('div', 'row'); grid.style.gridTemplateColumns = '1fr 1fr'; grid.style.gap = '16px';
            grid.append(left, right);
            row2.append(grid);

            const btnDup = el('button', 'ghost', 'Duplicate');
            const btnDel = el('button', 'ghost danger', 'Remove');
            btnDup.onclick = () => { spec.checks.splice(idx + 1, 0, JSON.parse(JSON.stringify(r))); render(); };
            btnDel.onclick = () => { spec.checks.splice(idx, 1); render(); };
            row2.append(btnDup, btnDel);

            renderColumns();
        } else {
            const btnDup = el('button', 'ghost', 'Duplicate'); const btnDel = el('button', 'ghost danger', 'Remove');
            btnDup.onclick = () => { spec.checks.splice(idx + 1, 0, JSON.parse(JSON.stringify(r))); render(); };
            btnDel.onclick = () => { spec.checks.splice(idx, 1); render(); };
            row2.append(btnDup, btnDel);
        }
    }

    showFor(r.type); b.append(row2);

    typeSel.onchange = () => { r.type = typeSel.value; showFor(r.type); refreshJsonPreview(); };
    pts.onchange = () => { r.points = +pts.value || 0; refreshJsonPreview(); };
    cell.onchange = () => { r.cell = cell.value.trim() || null; refreshJsonPreview(); };
    range.onchange = () => { r.range = range.value.trim() || null; refreshJsonPreview(); };
    note.onchange = () => { r.note = note.value.trim() || null; refreshJsonPreview(); };
    exf.onchange = () => { r.expected_formula = exf.value.trim() || null; refreshJsonPreview(); };
    regex.onchange = () => { r.expected_formula_regex = regex.value.trim() || null; refreshJsonPreview(); };
    fromKey.input.onchange = () => { r.expected_from_key = fromKey.input.checked ? true : null; refreshJsonPreview(); };
    absReq.input.onchange = () => { r.require_absolute = absReq.input.checked ? true : null; refreshJsonPreview(); };
    numFmt.onchange = () => { if (!r.format) r.format = {}; r.format.number_format = numFmt.value.trim() || null; refreshJsonPreview(); };
    start.onchange = () => { r.start = start.value === '' ? null : +start.value; refreshJsonPreview(); };
    step.onchange = () => { r.step = step.value === '' ? null : +step.value; refreshJsonPreview(); };
    btnDel.onclick = () => { spec.checks.splice(idx, 1); render(); };
    btnDup.onclick = () => { spec.checks.splice(idx + 1, 0, JSON.parse(JSON.stringify(r))); render(); };
    return b;
}

/** Wrap an input control with a small stacked label for consistent UI.
* @param {HTMLElement} input
* @param {string} label
* @returns {HTMLElement}
*/
function labelled(input, label) {
    const box = el('div', 'stack labelled');
    box.append(el('div', 'tiny muted', label), input);
    return box;
}

/** Create a labeled checkbox control.
 * @param {string} txt - Label text.
 * @param {boolean} checked - Initial state.
 * @returns {{wrap:HTMLLabelElement, input:HTMLInputElement}}
 */
function checkbox(txt, checked) { const w = el('label', 'checkbox'); const c = el('input'); c.type = 'checkbox'; c.checked = checked; const s = el('span', '', txt); w.append(c, s); return { wrap: w, input: c }; }

// ---------- JSON <-> builder

/** Update the JSON preview textbox from the collected rubric structure. */
function refreshJsonPreview() { $('#jsonBox').value = JSON.stringify(collectRubric(), null, 2); }
/** Gather current rubric state from the UI toggles/inputs (mutates global rubric and returns it).
 * @returns {object} rubric
 */
function collectRubric() {
    rubric.points = +($('#totalPoints').value || rubric.points || 0) || 0;
    rubric.report = { include_pass_fail_column: $('#optPassFail').checked, include_comments: $('#optComments').checked };
    return rubric;
}

/** Load a rubric JSON into the builder, normalizing minimal shape and re-rendering.
 * @param {object} json
 */
function loadRubric(json) {
    rubric = { points: json.points ?? 5, report: json.report ?? { include_pass_fail_column: true, include_comments: true }, sheets: {} };
    if (json.sheets) { for (const [name, spec] of Object.entries(json.sheets)) { rubric.sheets[name] = { checks: Array.isArray(spec.checks) ? spec.checks : [] }; } }
    render();
}

// ----- SINGLE ensureTable (complete version) -----

/** Normalize and attach a table spec to a rule (UI-friendly shape).
 * @param {object} r - Rule object (mutated).
 * @returns {{sheet:(string|null),name_like:(string|null),columns:string[],require_order:(boolean|null),range_ref:(string|null),rows:(number|null),cols:(number|null),allow_extra_rows:(boolean|null),allow_extra_cols:(boolean|null),contains_rows:object[],body_match:(boolean|null),body_order_matters:(boolean|null),body_case_sensitive:(boolean|null),body_trim:boolean,body_rows:(Array<Array<string>>|null)}}
 */
function ensureTable(r) {
    const t = r.table || {};
    r.table = {
        sheet: t.sheet ?? null,
        name_like: t.name_like ?? null,
        columns: Array.isArray(t.columns) ? t.columns.filter(Boolean) : [],
        require_order: (typeof t.require_order === 'boolean') ? t.require_order : null,

        range_ref: t.range_ref ?? null,
        rows: (typeof t.rows === 'number') ? t.rows : (t.rows ? Number(t.rows) : null),
        cols: (typeof t.cols === 'number') ? t.cols : (t.cols ? Number(t.cols) : null),
        allow_extra_rows: (typeof t.allow_extra_rows === 'boolean') ? t.allow_extra_rows : (t.allow_extra_rows ? true : null),
        allow_extra_cols: (typeof t.allow_extra_cols === 'boolean') ? t.allow_extra_cols : (t.allow_extra_cols ? true : null),
        contains_rows: Array.isArray(t.contains_rows) ? t.contains_rows : [],

        body_match: (typeof t.body_match === 'boolean') ? t.body_match : (t.body_rows ? true : null),
        body_order_matters: (typeof t.body_order_matters === 'boolean') ? t.body_order_matters : null,
        body_case_sensitive: (typeof t.body_case_sensitive === 'boolean') ? t.body_case_sensitive : null,
        body_trim: (typeof t.body_trim === 'boolean') ? t.body_trim : true,
        body_rows: Array.isArray(t.body_rows) ? t.body_rows : null
    };
    return r.table;
}

/** Normalize section names so tiny differences (“and” vs “&”, double spaces, case) don’t split groups. */
function normalizeSectionName(s) {
    return (s || "(no section)")
        .trim()
        .replace(/\s*&\s*/g, " & ")
        .replace(/\s+and\s+/ig, " & ")
        .replace(/\s+/g, " ")
        .toLowerCase();
}

/** Pick a pretty display name for a section group (prefer meta.sectionOrder spelling if present). */
function chooseDisplayName(normKey, metaOrder) {
    if (metaOrder && metaOrder.length) {
        for (const name of metaOrder) {
            const n = normalizeSectionName(name);
            if (n === normKey) return name;   // use exact spelling from meta
        }
    }
    // Fallback: title-case the normalized key
    return normKey.replace(/\b\w/g, m => m.toUpperCase());
}

/** Sum section totals once. */
function sectionTotals(checks) {
    let points = 0, earned = 0;
    for (const c of checks) {
        points += c.points || 0;
        earned += c.earned || 0;
    }
    return { points, earned };
}

/** Grouped results renderer.
 *  graded: object with .checks (your current grading result for one student)
 *  rubric: the rubric JSON; we read rubric.meta.sectionOrder if present
 */

/** Render grouped grading results by normalized section order, using rubric meta.sectionOrder when present.
 * @param {{checks:Array}} graded - One student's graded payload.
 * @param {{meta?:{sectionOrder?:string[]}}} rubric - Rubric JSON for ordering hints.
 * @returns {HTMLElement}
 */
function renderGroupedResults(graded, rubric) {
    const root = document.createElement('div');
    const metaOrder = rubric?.meta?.sectionOrder || [];

    // 1) Bucket checks by normalized section key
    const byKey = new Map(); // normKey -> { display, checks[] }
    for (const ck of (graded.checks || [])) {
        const key = normalizeSectionName(ck.section || "(no section)");
        if (!byKey.has(key)) byKey.set(key, { display: null, checks: [] });
        byKey.get(key).checks.push(ck);
    }

    // 2) Decide display names using meta.sectionOrder first
    for (const [key, group] of byKey) {
        group.display = chooseDisplayName(key, metaOrder);
    }

    // 3) Build rendering order: meta.sectionOrder first, then any remaining groups
    const orderKeys = [];
    const used = new Set();
    for (const name of metaOrder) {
        const k = normalizeSectionName(name);
        if (byKey.has(k) && !used.has(k)) { orderKeys.push(k); used.add(k); }
    }
    for (const k of byKey.keys()) if (!used.has(k)) orderKeys.push(k);

    // 4) Render each section once: header + its checks
    for (const key of orderKeys) {
        const group = byKey.get(key);
        if (!group) continue;
        const { points, earned } = sectionTotals(group.checks);

        // Header (sticky-friendly)
        const header = document.createElement('div');
        header.className = 'result-section-header';
        header.textContent = `${group.display} — ${earned}/${points}`;
        root.appendChild(header);

        // Body: all checks in this section
        const body = document.createElement('div');
        body.className = 'result-section-body';
        for (const c of group.checks) {
            body.appendChild(renderCheckRow(c)); // ⬅️ your existing per-check row renderer
        }
        root.appendChild(body);
    }

    return root;
}

// ---------- actions
$('#btnAddSheet').onclick = () => { let base = 'Sheet', i = 1, name = base; while (rubric.sheets[name]) name = base + ' ' + (++i); rubric.sheets[name] = { checks: [emptyRule('formula')] }; render(); };
$('#btnClear').onclick = () => { rubric = { points: +($('#totalPoints').value || 5), report: { include_pass_fail_column: $('#optPassFail').checked, include_comments: $('#optComments').checked }, sheets: {} }; render(); };
$('#btnGenerateJson').onclick = () => { $('#jsonBox').value = JSON.stringify(collectRubric(), null, 2); };
document.getElementById('btnDownloadJson')
    ?.addEventListener('click', (e) => {
        e.preventDefault();
        downloadRubricJSON();
    });

// --- Rubric Builder collapse/expand ---
const builderCard = document.querySelector('.builder-card') || document.getElementById('builderBody')?.closest('.card');
const builderBody = document.getElementById('builderBody');
const builderToggle = document.getElementById('btnToggleBuilder');

const builderRail = document.querySelector('.rules-rail');   // right-side sticky card
const sectionsDock = document.getElementById('sectionsDock'); // slide-out sections panel
const sectionsDockBtn = document.getElementById('toggleSectionsDock'); // floating button

/** Collapse/expand the main builder card and associated side panels; persist state in localStorage.
* @param {boolean} on - Whether the builder should be collapsed.
*/
function setBuilderCollapsed(on) {
    // Prefer toggling on the card if present; fall back to inline style
    builderCard?.classList.toggle('collapsed', !!on);
    if (!builderCard && builderBody) builderBody.style.display = on ? 'none' : '';

    // Hide/show the extra pieces so they behave with the builder
    [builderRail, sectionsDock, sectionsDockBtn].forEach(el => {
        if (el) el.classList.toggle('hidden', !!on);
    });

    if (builderToggle) builderToggle.textContent = on ? 'Show Builder' : 'Hide Builder';
    try { localStorage.setItem('builderCollapsed', on ? '1' : '0'); } catch { }
}

setBuilderCollapsed((localStorage.getItem('builderCollapsed') === '1'));

// click handler
builderToggle?.addEventListener('click', (e) => {
    e.preventDefault();
    const next = builderCard
        ? !builderCard.classList.contains('collapsed')
        : (builderBody?.style.display !== 'none');
    setBuilderCollapsed(next);
});

// Auto-rubric from key
async function generateFromKey() {
    const keyInput = $('#keyFile');
    const hint = ($('#sheetHint')?.value || '').trim();
    const allSheets = !!$('#allSheets')?.checked;
    const totalStr = ($('#totalPoints')?.value || '').trim();
    if (!keyInput?.files?.length) { alert('Please choose a key workbook first.'); return; }

    const key = keyInput.files[0];
    const url = `/api/auto-rubric`;
    const form = new FormData();
    form.append('key', key);
    if (hint) form.append('sheet', hint);
    form.append('all', allSheets ? 'true' : 'false');
    if (totalStr) form.append('total', totalStr);

    try {
        const res = await fetch(url, { method: 'POST', body: form });
        if (!res.ok) { const t = await res.text(); throw new Error(t || `HTTP ${res.status}`); }
        const rub = await res.json();
        $('#jsonBox').value = JSON.stringify(rub, null, 2);
        rubric = rub;
        render();
    } catch (err) {
        console.error(err);
        const msg = (err && err.message) ? err.message : String(err ?? 'unknown error');
        alert('Generation failed: ' + msg);
    }
}
$('#btnGenerate').addEventListener('click', e => { e.preventDefault(); generateFromKey(); });

// ---------- grading
$('#btnGradeWithFile').onclick = async () => {
    try {
        const key = $('#keyFile').files[0];
        const rub = $('#rubricFile').files[0];
        const studs = $('#studentFiles').files;
        if (!key || !rub || !studs.length) {
            alert('Pick a key, a rubric file, and at least one student workbook.');
            return;
        }

        const fd = new FormData();
        fd.append('key', key);
        fd.append('rubric', rub);
        for (const f of studs) fd.append('students', f);

        const res = await fetch('/api/grade', { method: 'POST', body: fd });
        const txt = await res.text();
        if (!res.ok) { alert(txt || `HTTP ${res.status}`); return; }
        showReport(txt);
    } catch (err) { console.error(err); alert('Grade failed'); }
};

$('#btnGradeWithJson').onclick = async () => {
    try {
        const key = $('#keyFile').files[0];
        const studs = $('#studentFiles').files;
        if (!key || !studs.length) {
            alert('Pick a key and at least one student workbook.');
            return;
        }

        const jsonBlob = new Blob([$('#jsonBox').value], { type: 'application/json' });
        const fd = new FormData();
        fd.append('key', key);
        fd.append('rubricJson', jsonBlob, 'rubric.json');
        for (const f of studs) fd.append('students', f);

        const res = await fetch('/api/grade', { method: 'POST', body: fd });
        const txt = await res.text();
        if (!res.ok) { alert(txt || `HTTP ${res.status}`); return; }
        showReport(txt);
    } catch (err) { console.error(err); alert('Grade failed'); }
};

// ---------- results viewer
let lastReportRaw = '';
$('#btnDownloadReport').onclick = () => { if (!lastReportRaw) { alert('Nothing to download yet.'); return; } download('report.json', lastReportRaw); };

/** Produce a stable “rule key” string used to map rules to section labels across runs.
* @param {object} r - Rule.
* @returns {string}
*/
function keyForRule(r) {
    const t = (r.type || 'rule');

    if (t === 'chart' && r.chart) {
        const sh = (r.chart.sheet || '').trim();
        const nm = (r.chart.name_like || '').trim();
        if (nm) return `chart:${sh}${(sh && nm ? '/' : '')}${nm}`;

        const typ = (r.chart.type || '').trim().toLowerCase();
        if (sh && typ) return `chart:${sh}|type=${typ}`;

        const title = (r.chart.title || '').trim().toLowerCase().replace(/\s+/g, ' ');
        if (sh && title) return `chart:${sh}|title=${title}`;

        if (sh) return `chart:${sh}`;
        return 'chart';
    }

    if (t === 'table' && r.table) {
        const sh = (r.table.sheet || '').trim();
        const nm = (r.table.name_like || r.table.tableNameLike || '').trim();
        if (nm) return `table:${sh}${(sh && nm ? '/' : '')}${nm}`;
        if (sh) return `table:${sh}`;
    }

    if (t === 'pivot' && r.pivot) {
        const sh = (r.pivot.sheet || '').trim();
        const nm = (r.pivot.tableNameLike || '').trim();
        if (nm) return `pivot:${sh}${(sh && nm ? '/' : '')}${nm}`;
        if (sh) return `pivot:${sh}`;
    }

    const loc = (r.cell && r.cell.trim()) || (r.range && r.range.trim()) || '';
    return t + (loc ? (':' + loc) : '');
}

/** Build a Map of { stableRuleKey -> section } from the current rubric (for inference).
* @returns {Map<string,string>}
*/
function buildSectionIndex() {
    const map = new Map();
    for (const [_sheetName, spec] of Object.entries(rubric.sheets || {})) {
        for (const rr of (spec.checks || [])) {
            const k = keyForRule(rr);
            if (k && rr.section) map.set(k, rr.section);
        }
    }
    return map;
}

/** Heuristically guess a section name from a backend label/check string.
* @param {string} label
* @returns {string} - Suggested section or '(no section)'.
*/
function guessSectionFromLabel(label) {
    const s = String(label || '').toLowerCase();

    // pivots
    if (s.startsWith('pivot_layout') || s.startsWith('pivot') || s.includes(" pivot"))
        return 'Pivot Tables';

    // charts
    if (s.startsWith('chart:') || s.includes("chart '"))
        return 'Chart';

    // excel tables
    if (s.startsWith('table:') || s.includes("table '"))
        return 'Excel Table';

    // conditional formatting
    if (s.startsWith('conditional_') || s.startsWith('cf:'))
        return 'Conditional Formatting';

    // value/range checks often are from “Sort and Filtering” or “Imported Data”;
    // leave them to the normal index (they usually have an entry). If none, no change.
    return '(no section)';
}

/** Render the full results card/table from a raw JSON response string produced by /api/grade.
* Groups by sheet then section, ordered by rubric hints; updates download cache.
* @param {string} text - JSON string from server (single or multi-student).
*/
function showReport(text) {
    // 1) Parse payload
    let data;
    try { data = JSON.parse(text); } catch { alert('Bad JSON from server'); return; }
    lastReportRaw = text;

    // 2) Normalize rows (supports single result, {students}, {results})
    let rows = [];
    if (Array.isArray(data)) rows = data;
    else if (Array.isArray(data.students)) rows = data.students;
    else if (Array.isArray(data.results)) rows = data.results;

    const body = document.getElementById('resultsBody');
    if (!body) { console.error('#resultsBody not found'); alert('Internal error: results container missing.'); return; }
    body.innerHTML = '';

    if (!rows.length) {
        const card = el('div', 'card', '<div class="muted">No results to show.</div>');
        body.appendChild(card);
        document.getElementById('resultsCard').style.display = '';
        return;
    }

    // 3) Try to read rubric (for section ordering)
    let rubricJson = null;
    try {
        const j = document.getElementById('jsonBox')?.value;
        if (j) rubricJson = JSON.parse(j);
    } catch (_) { /* ignore */ }
    const rubricMeta = rubricJson?.meta || {};
    const rubricSheets = rubricJson?.sheets || {};

    // Helper to render a single row (unchanged behavior)
    const renderRow = (d) => {
        const ok = (d.passed ?? d.Passed);
        const okBool = typeof ok === 'boolean' ? ok : (Number(d.earned) >= Number(d.points));
        const partialBool = !okBool && Number(d.earned) > 0;

        const tr = document.createElement('tr');
        tr.dataset.status = okBool ? 'pass' : (partialBool ? 'partial' : 'fail');

        tr.innerHTML = `
      <td class="mono">${escapeHtml(d.check ?? d.rule ?? '')}</td>
      <td class="mono">${fmtNum(d.points)}</td>
      <td class="mono">${fmtNum(d.earned)}</td>
      <td>${okBool
                ? '<span class="tag pass">PASS</span>'
                : (partialBool
                    ? '<span class="tag partial">PARTIAL</span>'
                    : '<span class="tag fail">FAIL</span>')}
      </td>
      <td>${escapeHtml(d.comment ?? d.message ?? '')}</td>`;
        return tr;
    };

    for (const r of rows) {
        // error rows
        if (r && r.error && (r.student || r.name)) {
            const card = el('div', 'card');
            const hdr = el('div', 'hstack');
            hdr.innerHTML = `<h3>${escapeHtml(r.student ?? r.name ?? 'Student')}</h3>
                       <span class="tag fail">FAIL</span>
                       <span class="muted">${escapeHtml(String(r.error))}</span>`;
            card.appendChild(hdr);
            body.appendChild(card);
            continue;
        }

        const isNested = r && typeof r.grade === 'object';
        const grade = isNested ? r.grade : r;

        const name = isNested ? (r.student ?? r.name ?? 'Student')
            : (grade.name ?? r.student ?? 'Student');
        const points = grade.total_points ?? grade.points ?? grade.Points ?? 0;
        const earned = grade.score_numeric ?? grade.earned ?? grade.Earned ?? 0;
        const details = grade.details ?? grade.Details ?? [];

        const pNum = Number(points), eNum = Number(earned);
        const headerPass = pNum > 0 && Math.abs(eNum - pNum) < 1e-9;
        const headerClass = headerPass ? 'pass' : 'fail';
        const headerStatus = headerPass ? 'PASS' : 'FAIL';

        const card = el('div', 'card');
        const hdr = el('div', 'hstack');
        hdr.innerHTML = `<h3>${escapeHtml(name)}</h3>
                     <span class="tag ${headerClass}">${headerStatus}</span>
                     <span class="muted">Score ${fmtNum(earned)}/${fmtNum(points)}</span>`;
        card.appendChild(hdr);

        const tbl = el('table');
        const thead = el('thead');
        thead.innerHTML = `<tr>
                                      <th>Check</th><th>Points</th><th>Earned</th><th>Result</th><th>Comment</th>
                                    </tr>`;
        tbl.appendChild(thead);

        const tb = document.createElement('tbody');

        // ===== NEW: group strictly by SHEET then by SECTION coming from backend =====
        // Build: Map<sheet, Map<section, rows[]>>
        const bySheet = new Map();
        for (const d of details) {
            const sh = (d.sheet || '(Unknown sheet)').toString();
            const sec = (d.section || '(No section)').toString();
            if (!bySheet.has(sh)) bySheet.set(sh, new Map());
            const m = bySheet.get(sh);
            if (!m.has(sec)) m.set(sec, []);
            m.get(sec).push(d);
        }

        // Decide sheet render order (use rubric order of sheets if we can)
        const rubricSheetNames = Object.keys(rubricSheets);
        const isStrict = (typeof window !== 'undefined' && typeof window.isStrict !== 'undefined') ? window.isStrict === true : false;
        const seenSheets = new Set();
        const orderedSheets = [];

        // first: any sheets present in rubric, in rubric order
        for (const sh of rubricSheetNames) if (bySheet.has(sh)) { orderedSheets.push(sh); seenSheets.add(sh); }
        // then: any remaining sheets
        for (const sh of bySheet.keys()) if (!seenSheets.has(sh)) orderedSheets.push(sh);

        for (const sheetName of orderedSheets) {
            // Sheet header row
            const trSheet = document.createElement('tr');
            trSheet.className = 'section-row';
            trSheet.innerHTML = `<td colspan="5" style="font-weight:700;">Sheet: ${escapeHtml(sheetName)}</td>`;
            tb.appendChild(trSheet);

            const secMap = bySheet.get(sheetName);

            // Desired section order: per-sheet 'section_order' or global meta.sectionOrder
            const desiredOrder = (isStrict && Array.isArray(rubricMeta.sectionOrder)) ? rubricMeta.sectionOrder
                : (Array.isArray(rubricMeta.sectionOrder) && rubricMeta.sectionOrder.length
                    ? rubricMeta.sectionOrder
                    : (Array.isArray(rubricSheets?.[sheetName]?.section_order) ? rubricSheets[sheetName].section_order : []));

            // Compute totals & order
            const totals = new Map(); // section -> {points, earned}
            for (const [sec, arr] of secMap) {
                let pts = 0, earn = 0;
                for (const d of arr) { pts += (+d.points || 0); earn += (+d.earned || 0); }
                totals.set(sec, { points: pts, earned: earn });
            }

            const presentSecs = Array.from(secMap.keys());
            const used = new Set();
            const orderedSecs = [];
            for (const s of desiredOrder) if (secMap.has(s) && !used.has(s)) { orderedSecs.push(s); used.add(s); }
            for (const s of presentSecs) if (!used.has(s)) orderedSecs.push(s);

            // Render sections under this sheet
            for (const sec of orderedSecs) {
                const t = totals.get(sec) || { points: 0, earned: 0 };
                const trHeader = document.createElement('tr');
                trHeader.className = 'section-row';
                trHeader.innerHTML = `
                                              <td colspan="5" style="font-weight:600; padding-top:8px;">
                                                ${escapeHtml(sec)} <span class="muted"> — ${fmtNum(t.earned)}/${fmtNum(t.points)}</span>
                                              </td>`;
                tb.appendChild(trHeader);

                for (const d of secMap.get(sec)) {
                    tb.appendChild(renderRow(d));
                }
            }
        }
        // ===== END NEW GROUPING =====

        tbl.appendChild(tb);
        card.appendChild(tbl);
        body.appendChild(card);
    }

    const on = localStorage.getItem('onlyMisses') === '1' || document.getElementById('toggleOnlyMisses')?.checked;
    document.getElementById('resultsCard')?.classList?.toggle('only-misses', !!on);

    document.getElementById('resultsCard').style.display = '';
    window.scrollTo({ top: document.getElementById('resultsCard').offsetTop - 10, behavior: 'smooth' });
}


const toggleOnlyMisses = document.getElementById('toggleOnlyMisses');
if (toggleOnlyMisses) {
    const saved = localStorage.getItem('onlyMisses') === '1';
    toggleOnlyMisses.checked = saved;
    toggleOnlyMisses.addEventListener('change', () => {
        localStorage.setItem('onlyMisses', toggleOnlyMisses.checked ? '1' : '0');
        document.getElementById('resultsCard').classList.toggle('only-misses', toggleOnlyMisses.checked);
    });
}

/** Format a numeric-ish value into a compact string (empty for null/NaN).
* @param {unknown} v
* @returns {string}
*/
function fmtNum(v) { return (v == null || isNaN(+v)) ? '' : (+v).toString(); }

/** Escape minimal HTML entities (& and <) in a string.
 * @param {unknown} s
 * @returns {string}
 */
function escapeHtml(s) { return (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;'); }

// boot
render();
// >>> Forced settings: always group by sections with strict global section order
UI.groupBySection = true;
UI.strictSectionOrder = true;
const _elGroupChk = document.getElementById('groupBySection');
if (_elGroupChk) { try { _elGroupChk.checked = true; _elGroupChk.disabled = true; _elGroupChk.closest('.form-check')?.classList.add('d-none'); } catch { } }
const _elStrictChk = document.getElementById('strictOrder');
if (_elStrictChk) { try { _elStrictChk.checked = true; _elStrictChk.disabled = true; _elStrictChk.closest('.form-check')?.classList.add('d-none'); } catch { } }
// <<< Forced settings



// >>> Override results renderer to always use global grouped-by-section ordering
(function () {
    try {
        const _orig = window.renderResults || null;
        window.renderResults = function () {
            UI.groupBySection = true;
            UI.strictSectionOrder = true;
            if (typeof window.renderGroupedResults === 'function') {
                return window.renderGroupedResults();
            }
            // Fallback: call original if grouped renderer is unavailable
            if (typeof _orig === 'function') return _orig();
        };
    } catch (e) { console.warn('Override renderResults failed:', e); }
})();
// <<< Override
