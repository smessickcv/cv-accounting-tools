/* cv-settings.js
 * Shared "Mappings & Settings" editor for the Cardone Ventures internal tools.
 *
 * Each tool wraps a hard-coded mapping in CVSettings.use(id, defaultValue, schema).
 * use() returns the user's saved version (from this browser) when one exists and
 * the built-in default otherwise, so every later line of the tool sees the edit.
 * The floating "Mappings" button opens an editor with a dropdown of every mapping
 * in the tool, where lines can be edited, added, and deleted. Saved changes live
 * in this browser (localStorage) and take effect after the page reloads.
 *
 * Schema kinds:
 *   map      object of key -> value            columns: ['Key','Value']
 *   list     array of strings                  columns: ['Value']
 *   scalar   one string                        columns: [label]
 *   lines    string of delimiter-split lines   columns: [...], delim: '|'
 *   records  array of objects                  fields: [{key,label,type}]
 *   objmap   object of key -> object           keyLabel, fields: [{key,label,type}]
 * Field types: str (default), nstr (blank = null), list (comma separated), num, bool
 */
(function () {
  'use strict';
  if (window.CVSettings) return;

  var script = document.currentScript;
  var TOOL = (script && script.getAttribute('data-tool')) ||
    (location.pathname.split('/').pop() || 'tool').replace(/\.html?$/i, '').toLowerCase().replace(/[^a-z0-9]+/g, '-');
  var TITLE = (script && script.getAttribute('data-title')) || document.title || TOOL;
  var PREFIX = 'cvset:' + TOOL + ':';

  var sets = [];
  var links = [];
  var ui = null;

  function clone(v) { return v === undefined ? v : JSON.parse(JSON.stringify(v)); }
  function readSaved(id) {
    try { var r = localStorage.getItem(PREFIX + id); return r === null ? undefined : JSON.parse(r); }
    catch (e) { return undefined; }
  }
  function writeSaved(id, v) { localStorage.setItem(PREFIX + id, JSON.stringify(v)); }
  function clearSaved(id) { try { localStorage.removeItem(PREFIX + id); } catch (e) {} }
  function sameShape(a, b) { return typeof a === typeof b && Array.isArray(a) === Array.isArray(b) && a !== null; }

  // ---------- public: use ----------
  function use(id, def, schema) {
    schema = schema || { kind: 'map' };
    var entry = { id: id, schema: schema, def: clone(def), label: schema.label || id, hint: schema.hint || '' };
    sets.push(entry);
    var saved = readSaved(id);
    entry.overridden = saved !== undefined && sameShape(saved, def);
    return entry.overridden ? saved : def;
  }

  // store(): a mapping the tool already keeps in its own storage. read() returns the
  // current value in the editor's shape; write(value) saves it back the way the tool expects.
  function store(opts) {
    var schema = opts.schema || { kind: 'map' };
    var empty = schema.kind === 'scalar' ? '' : (schema.kind === 'lines' ? '' : (schema.kind === 'list' || schema.kind === 'records') ? [] : {});
    var entry = { id: opts.id, schema: schema, label: opts.label || opts.id, hint: opts.hint || schema.hint || '',
      def: opts.def !== undefined ? clone(opts.def) : empty, ext: opts, overridden: false };
    sets.push(entry);
    return entry;
  }

  function link(opts) { links.push(opts); }

  // ---------- value <-> rows ----------
  function fieldToStr(f, v) {
    var t = f.type || 'str';
    if (t === 'list') return (v || []).join(', ');
    if (t === 'nstr') return v == null ? '' : String(v);
    if (t === 'bool') return v ? 'true' : 'false';
    return v == null ? '' : String(v);
  }
  function strToField(f, s) {
    var t = f.type || 'str';
    s = String(s == null ? '' : s);
    if (t === 'list') return s.split(',').map(function (x) { return x.trim(); }).filter(Boolean);
    if (t === 'nstr') return s.trim() === '' ? null : s.trim();
    if (t === 'bool') return /^(true|yes|y|1)$/i.test(s.trim());
    if (t === 'num') { var n = Number(s); if (s.trim() === '' || isNaN(n)) throw new Error('"' + s + '" is not a number (' + f.label + ')'); return n; }
    return s.trim();
  }
  function columnsOf(e) {
    var s = e.schema;
    if (s.kind === 'map') return s.columns || ['Key', 'Value'];
    if (s.kind === 'list') return s.columns || ['Value'];
    if (s.kind === 'scalar') return [e.label];
    if (s.kind === 'lines') return s.columns;
    if (s.kind === 'records') return s.fields.map(function (f) { return f.label; });
    if (s.kind === 'objmap') return [s.keyLabel || 'Key'].concat(s.fields.map(function (f) { return f.label; }));
    return ['Value'];
  }
  function currentValue(e) {
    if (e.ext) {
      try { var xv = e.ext.read(); if (xv != null && sameShape(xv, e.def)) return xv; } catch (err) {}
      return e.def;
    }
    var v = readSaved(e.id);
    return v !== undefined && sameShape(v, e.def) ? v : e.def;
  }
  function toRows(e, value) {
    var s = e.schema, k = s.kind;
    if (k === 'map') return Object.keys(value).map(function (key) { return [key, String(value[key])]; });
    if (k === 'list') return value.map(function (v) { return [String(v)]; });
    if (k === 'scalar') return [[String(value)]];
    if (k === 'lines') {
      var d = s.delim || '|', n = s.columns.length;
      return value.split('\n').filter(function (l) { return l.length; }).map(function (l) {
        var p = l.split(d);
        if (p.length > n) p = p.slice(0, n - 1).concat([p.slice(n - 1).join(d)]);
        while (p.length < n) p.push('');
        return p;
      });
    }
    if (k === 'records') return value.map(function (o) { return s.fields.map(function (f) { return fieldToStr(f, o[f.key]); }); });
    if (k === 'objmap') return Object.keys(value).map(function (key) {
      var o = value[key] || {};
      return [key].concat(s.fields.map(function (f) { return fieldToStr(f, o[f.key]); }));
    });
    return [];
  }
  function blankRow(r) { return r.every(function (c) { return String(c).trim() === ''; }); }
  function fromRows(e, rows) {
    var s = e.schema, k = s.kind, i, r, out;
    rows = rows.filter(function (r) { return !blankRow(r); });
    if (k === 'scalar') return rows.length ? String(rows[0][0]).trim() : '';
    if (k === 'list') return rows.map(function (r) { return String(r[0]).trim(); });
    if (k === 'map') {
      out = {};
      for (i = 0; i < rows.length; i++) {
        r = rows[i];
        var key = String(r[0]).trim();
        if (!key) throw new Error('Row ' + (i + 1) + ': the first column cannot be blank.');
        if (Object.prototype.hasOwnProperty.call(out, key)) throw new Error('Duplicate key "' + key + '". Each key can appear once.');
        out[key] = s.valueType === 'number' ? Number(r[1]) : String(r[1]).trim();
        if (s.valueType === 'number' && isNaN(out[key])) throw new Error('Row ' + (i + 1) + ': "' + r[1] + '" is not a number.');
      }
      return out;
    }
    if (k === 'lines') {
      var d = s.delim || '|';
      for (i = 0; i < rows.length; i++) if (!String(rows[i][0]).trim()) throw new Error('Row ' + (i + 1) + ': the first column cannot be blank.');
      return rows.map(function (r) { return r.map(function (c) { return String(c).trim(); }).join(d); }).join('\n');
    }
    if (k === 'records') {
      return rows.map(function (r, idx) {
        if (!String(r[0]).trim()) throw new Error('Row ' + (idx + 1) + ': ' + s.fields[0].label + ' cannot be blank.');
        var o = {};
        s.fields.forEach(function (f, j) { o[f.key] = strToField(f, r[j]); });
        return o;
      });
    }
    if (k === 'objmap') {
      out = {};
      rows.forEach(function (r, idx) {
        var key = String(r[0]).trim();
        if (!key) throw new Error('Row ' + (idx + 1) + ': ' + (s.keyLabel || 'Key') + ' cannot be blank.');
        if (Object.prototype.hasOwnProperty.call(out, key)) throw new Error('Duplicate key "' + key + '".');
        var o = {};
        s.fields.forEach(function (f, j) { o[f.key] = strToField(f, r[j + 1]); });
        out[key] = o;
      });
      return out;
    }
    return null;
  }

  // ---------- UI ----------
  var CSS = [
    ':host{all:initial}',
    '*{box-sizing:border-box;font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif}',
    '.fab{position:fixed;right:16px;bottom:16px;z-index:2147483000;background:#1f2937;color:#fff;border:1px solid #4b5563;border-radius:999px;padding:8px 14px;font-size:12px;font-weight:600;cursor:pointer;box-shadow:0 2px 10px rgba(0,0,0,.35)}',
    '.fab:hover{background:#374151}',
    '.ov{position:fixed;inset:0;z-index:2147483001;background:rgba(0,0,0,.55);display:none;align-items:center;justify-content:center;padding:16px}',
    '.ov.on{display:flex}',
    '.dlg{background:#111827;color:#e5e7eb;border:1px solid #374151;border-radius:10px;width:min(1100px,100%);max-height:92vh;display:flex;flex-direction:column;box-shadow:0 10px 40px rgba(0,0,0,.5)}',
    '.hd{display:flex;align-items:center;justify-content:space-between;padding:12px 16px;border-bottom:1px solid #374151}',
    '.hd h2{margin:0;font-size:15px;font-weight:700}',
    '.hd .sub{font-size:11px;color:#9ca3af;margin-top:2px}',
    '.x{background:none;border:none;color:#9ca3af;font-size:20px;cursor:pointer;line-height:1}',
    '.x:hover{color:#fff}',
    '.bar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;padding:10px 16px;border-bottom:1px solid #1f2937}',
    'select,input[type=text],input[type=search]{background:#0b1220;color:#e5e7eb;border:1px solid #374151;border-radius:6px;padding:6px 8px;font-size:12px}',
    'select{min-width:260px;max-width:100%}',
    'input:focus,select:focus{outline:2px solid #3b82f6;outline-offset:-1px}',
    '.btn{background:#1f2937;color:#e5e7eb;border:1px solid #4b5563;border-radius:6px;padding:6px 10px;font-size:12px;font-weight:600;cursor:pointer}',
    '.btn:hover{background:#374151}',
    '.btn.pri{background:#2563eb;border-color:#2563eb;color:#fff}',
    '.btn.pri:hover{background:#1d4ed8}',
    '.btn.dng{border-color:#7f1d1d;color:#fca5a5}',
    '.btn:disabled{opacity:.45;cursor:default}',
    '.desc{padding:8px 16px 0;font-size:12px;color:#9ca3af}',
    '.note{padding:6px 16px;font-size:12px}',
    '.note.ok{color:#86efac}.note.err{color:#fca5a5}.note.warn{color:#fcd34d}',
    '.wrap{overflow:auto;padding:8px 16px;flex:1;min-height:120px}',
    'table{border-collapse:collapse;width:100%;font-size:12px}',
    'th{position:sticky;top:0;background:#111827;text-align:left;padding:6px;font-size:10px;letter-spacing:.8px;text-transform:uppercase;color:#9ca3af;border-bottom:1px solid #374151}',
    'td{padding:2px 4px;border-bottom:1px solid #1f2937;vertical-align:top}',
    'td input{width:100%;min-width:90px;font-size:12px;padding:5px 6px}',
    'td.del{width:34px;text-align:center}',
    '.rm{background:none;border:none;color:#f87171;cursor:pointer;font-size:16px;padding:4px}',
    '.rm:hover{color:#fecaca}',
    '.ft{display:flex;flex-wrap:wrap;gap:8px;align-items:center;padding:10px 16px;border-top:1px solid #374151}',
    '.sp{flex:1}',
    '.cnt{font-size:11px;color:#9ca3af}',
    '.empty{padding:24px;text-align:center;color:#9ca3af;font-size:12px}'
  ].join('');

  function h(tag, attrs, kids) {
    var el = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (a) {
      if (a === 'text') el.textContent = attrs[a];
      else if (a.slice(0, 2) === 'on') el.addEventListener(a.slice(2), attrs[a]);
      else el.setAttribute(a, attrs[a]);
    });
    (kids || []).forEach(function (k) { el.appendChild(k); });
    return el;
  }

  function build() {
    if (ui) return;
    var host = h('div', { id: 'cv-settings-host' });
    var root = host.attachShadow({ mode: 'open' });
    root.appendChild(h('style', { text: CSS }));

    var fab = h('button', { 'class': 'fab', type: 'button', text: '\u2699 Mappings', title: 'View and edit this tool\u2019s mappings', onclick: function () { open(); } });
    var ov = h('div', { 'class': 'ov' });
    var dlg = h('div', { 'class': 'dlg', role: 'dialog', 'aria-label': 'Mappings and settings' });
    ov.appendChild(dlg);
    ov.addEventListener('mousedown', function (e) { if (e.target === ov) close(); });

    var title = h('h2', { text: 'Mappings & Settings' });
    var sub = h('div', { 'class': 'sub', text: TITLE });
    var hd = h('div', { 'class': 'hd' }, [h('div', null, [title, sub]), h('button', { 'class': 'x', type: 'button', text: '\u00d7', 'aria-label': 'Close', onclick: function () { close(); } })]);

    var sel = h('select', { 'aria-label': 'Mapping set' });
    var filter = h('input', { type: 'search', placeholder: 'Filter rows\u2026' });
    var cnt = h('span', { 'class': 'cnt' });
    var addBtn = h('button', { 'class': 'btn', type: 'button', text: '+ Add row' });
    var bar = h('div', { 'class': 'bar' }, [sel, filter, addBtn, cnt]);

    var desc = h('div', { 'class': 'desc' });
    var note = h('div', { 'class': 'note' });
    var wrap = h('div', { 'class': 'wrap' });

    var saveBtn = h('button', { 'class': 'btn pri', type: 'button', text: 'Save' });
    var saveReload = h('button', { 'class': 'btn pri', type: 'button', text: 'Save & reload' });
    var discardBtn = h('button', { 'class': 'btn', type: 'button', text: 'Discard changes' });
    var resetBtn = h('button', { 'class': 'btn dng', type: 'button', text: 'Reset to defaults' });
    var reloadBtn = h('button', { 'class': 'btn', type: 'button', text: 'Reload page', onclick: function () { location.reload(); } });
    var exportBtn = h('button', { 'class': 'btn', type: 'button', text: 'Export all' });
    var importBtn = h('button', { 'class': 'btn', type: 'button', text: 'Import' });
    var file = h('input', { type: 'file', accept: '.json,application/json', style: 'display:none' });
    var ft = h('div', { 'class': 'ft' }, [saveBtn, saveReload, discardBtn, resetBtn, h('span', { 'class': 'sp' }), reloadBtn, exportBtn, importBtn, file]);

    [hd, bar, desc, note, wrap, ft].forEach(function (n) { dlg.appendChild(n); });
    root.appendChild(fab);
    root.appendChild(ov);
    document.body.appendChild(host);

    ui = { host: host, ov: ov, sel: sel, filter: filter, cnt: cnt, addBtn: addBtn, desc: desc, note: note, wrap: wrap,
      saveBtn: saveBtn, saveReload: saveReload, discardBtn: discardBtn, resetBtn: resetBtn, reloadBtn: reloadBtn,
      exportBtn: exportBtn, importBtn: importBtn, file: file, work: {}, limit: 200, cur: 0, needsReload: false };
    reloadBtn.style.display = 'none';

    sel.addEventListener('change', function () { ui.cur = Number(sel.value); ui.limit = 200; filter.value = ''; render(); });
    filter.addEventListener('input', function () { ui.limit = 200; renderTable(); });
    addBtn.addEventListener('click', function () {
      var e = sets[ui.cur]; if (!e) return;
      var rows = workRows(e);
      rows.unshift(columnsOf(e).map(function () { return ''; }));
      filter.value = ''; markDirty(e); renderTable();
      var first = wrap.querySelector('tbody input'); if (first) first.focus();
    });
    saveBtn.addEventListener('click', function () { save(false); });
    saveReload.addEventListener('click', function () { save(true); });
    discardBtn.addEventListener('click', function () { var e = sets[ui.cur]; delete ui.work[e.id]; setNote(''); render(); });
    resetBtn.addEventListener('click', function () {
      var e = sets[ui.cur];
      var msg = e.ext ? 'Clear every row in "' + e.label + '"? This empties the list saved in this browser.' : 'Reset "' + e.label + '" to the built-in defaults? Your saved edits for this mapping will be removed.';
      if (!window.confirm(msg)) return;
      try { if (e.ext) { if (e.ext.reset) e.ext.reset(); else e.ext.write(e.def); } else clearSaved(e.id); } catch (err3) { setNote(err3.message, 'err'); return; }
      delete ui.work[e.id]; e.overridden = false; ui.needsReload = true;
      setNote('Reset. Reload the page to apply.', 'ok'); render();
    });
    exportBtn.addEventListener('click', exportAll);
    importBtn.addEventListener('click', function () { file.click(); });
    file.addEventListener('change', function () { importFile(file.files && file.files[0]); file.value = ''; });
    document.addEventListener('keydown', function (ev) { if (ev.key === 'Escape' && ui.ov.classList.contains('on')) close(); });
  }

  function workRows(e) {
    if (!ui.work[e.id]) ui.work[e.id] = { rows: toRows(e, currentValue(e)), dirty: false };
    return ui.work[e.id].rows;
  }
  function markDirty(e) { ui.work[e.id].dirty = true; refreshButtons(); refreshOptionLabels(); }
  function isDirty(e) { return !!(ui.work[e.id] && ui.work[e.id].dirty); }
  function setNote(msg, kind) { ui.note.textContent = msg || ''; ui.note.className = 'note' + (kind ? ' ' + kind : ''); }

  function refreshOptionLabels() {
    var opts = ui.sel.options, i;
    for (i = 0; i < sets.length; i++) {
      var e = sets[i];
      opts[i].textContent = e.label + (isDirty(e) ? '  *unsaved' : (e.overridden ? '  (customized)' : ''));
    }
  }
  function refreshButtons() {
    var e = sets[ui.cur]; if (!e) return;
    var d = isDirty(e);
    ui.saveBtn.disabled = !d; ui.saveReload.disabled = !d; ui.discardBtn.disabled = !d;
    ui.resetBtn.disabled = !e.overridden && !d && !e.ext;
    ui.resetBtn.textContent = e.ext ? 'Clear all rows' : 'Reset to defaults';
    ui.reloadBtn.style.display = ui.needsReload ? '' : 'none';
  }

  function save(thenReload) {
    var e = sets[ui.cur], rows = workRows(e), value;
    try { value = fromRows(e, rows); } catch (err) { setNote(err.message, 'err'); return; }
    try { if (e.ext) e.ext.write(value); else writeSaved(e.id, value); } catch (err2) { setNote('Could not save in this browser: ' + err2.message, 'err'); return; }
    if (!e.ext) e.overridden = true;
    ui.needsReload = true;
    ui.work[e.id] = { rows: toRows(e, value), dirty: false };
    setNote('Saved in this browser. Reload the page to apply it (files you have loaded will need to be loaded again).', 'ok');
    render();
    if (thenReload) location.reload();
  }

  function render() {
    var e = sets[ui.cur];
    ui.sel.value = String(ui.cur);
    refreshOptionLabels();
    var isLink = ui.cur >= sets.length;
    ui.desc.textContent = e ? (e.hint || '') : '';
    renderTable();
    refreshButtons();
  }

  function renderTable() {
    var e = sets[ui.cur];
    if (!e) {
      ui.wrap.textContent = ''; ui.cnt.textContent = '';
      ui.wrap.appendChild(h('div', { 'class': 'empty', text: links.length ? 'This tool keeps its mappings in its own editors. Use the buttons above to open them.' : 'This tool has no mapping tables. Nothing here is hard-coded.' }));
      return;
    }
    var rows = workRows(e), cols = columnsOf(e);
    var q = ui.filter.value.trim().toLowerCase();
    var wrap = ui.wrap; wrap.textContent = '';
    var matches = [];
    rows.forEach(function (r, i) {
      if (!q || r.some(function (c) { return String(c).toLowerCase().indexOf(q) !== -1; })) matches.push(i);
    });
    ui.cnt.textContent = matches.length + (q ? ' of ' + rows.length : '') + ' row' + (rows.length === 1 ? '' : 's');
    if (!matches.length) { wrap.appendChild(h('div', { 'class': 'empty', text: rows.length ? 'No rows match the filter.' : 'No rows yet. Use "+ Add row".' })); return; }
    var thead = h('thead', null, [h('tr', null, cols.map(function (c) { return h('th', { text: c }); }).concat([h('th', { text: '' })]))]);
    var tbody = h('tbody');
    matches.slice(0, ui.limit).forEach(function (ri) {
      var tr = h('tr');
      rows[ri].forEach(function (val, ci) {
        var inp = h('input', { type: 'text', value: val, 'aria-label': cols[ci] });
        inp.value = val;
        inp.addEventListener('input', function () { rows[ri][ci] = inp.value; if (!isDirty(e)) markDirty(e); });
        tr.appendChild(h('td', null, [inp]));
      });
      var rm = h('button', { 'class': 'rm', type: 'button', title: 'Delete this row', text: '\u00d7', 'aria-label': 'Delete row' });
      rm.addEventListener('click', function () { rows.splice(ri, 1); markDirty(e); renderTable(); });
      tr.appendChild(h('td', { 'class': 'del' }, [rm]));
      tbody.appendChild(tr);
    });
    wrap.appendChild(h('table', null, [thead, tbody]));
    if (matches.length > ui.limit) {
      var more = h('button', { 'class': 'btn', type: 'button', text: 'Show ' + Math.min(200, matches.length - ui.limit) + ' more (' + (matches.length - ui.limit) + ' hidden)' });
      more.style.margin = '10px 0';
      more.addEventListener('click', function () { ui.limit += 200; renderTable(); });
      wrap.appendChild(more);
    }
  }

  function exportAll() {
    var out = { format: 'cv-settings-bundle', tool: TOOL, exported: new Date().toISOString(), sets: {} };
    sets.forEach(function (e) { out.sets[e.id] = currentValue(e); });
    var blob = new Blob([JSON.stringify(out, null, 2)], { type: 'application/json' });
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob);
    a.download = TOOL + '-mappings.json'; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
    setNote('Exported ' + sets.length + ' mapping set' + (sets.length === 1 ? '' : 's') + '. Keep this file out of the repo if it holds company data.', 'ok');
  }
  function importFile(f) {
    if (!f) return;
    var r = new FileReader();
    r.onload = function () {
      try {
        var b = JSON.parse(r.result);
        if (!b || b.format !== 'cv-settings-bundle' || !b.sets) throw new Error('This is not a CV mappings bundle.');
        if (b.tool && b.tool !== TOOL && !window.confirm('This bundle is for "' + b.tool + '", not "' + TOOL + '". Import anyway?')) return;
        var n = 0, skipped = [];
        sets.forEach(function (e) {
          if (!(e.id in b.sets)) return;
          if (!sameShape(b.sets[e.id], e.def)) { skipped.push(e.label); return; }
          if (e.ext) e.ext.write(b.sets[e.id]); else { writeSaved(e.id, b.sets[e.id]); e.overridden = true; }
          delete ui.work[e.id]; n++;
        });
        ui.needsReload = ui.needsReload || n > 0;
        setNote('Imported ' + n + ' mapping set' + (n === 1 ? '' : 's') + (skipped.length ? '. Skipped (wrong shape): ' + skipped.join(', ') : '') + '. Reload the page to apply.', n ? 'ok' : 'warn');
        render();
      } catch (err) { setNote('Import failed: ' + err.message, 'err'); }
    };
    r.readAsText(f);
  }

  function open(id) {
    build();
    ui.sel.textContent = '';
    sets.forEach(function (e, i) { ui.sel.appendChild(h('option', { value: String(i), text: e.label })); });
    if (id) { for (var i = 0; i < sets.length; i++) if (sets[i].id === id) ui.cur = i; }
    if (ui.cur >= sets.length) ui.cur = 0;
    ui.ov.classList.add('on'); ui.filter.value = ''; ui.limit = 200;
    renderLinks();
    render();
  }
  function renderLinks() {
    // Linked editors that already exist inside a tool are listed under the table.
    var old = ui.ov.querySelector('.links'); if (old) old.remove();
    if (!links.length) return;
    var box = h('div', { 'class': 'note links' });
    box.appendChild(h('span', { text: 'Also editable in this tool: ' }));
    links.forEach(function (l, i) {
      var b = h('button', { 'class': 'btn', type: 'button', text: l.label, title: l.hint || '' });
      b.style.marginRight = '6px';
      b.addEventListener('click', function () { close(); try { l.open(); } catch (e) {} });
      box.appendChild(b);
    });
    ui.desc.parentNode.insertBefore(box, ui.desc.nextSibling);
  }
  function close() { if (ui) ui.ov.classList.remove('on'); }

  function mount() { build(); }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', mount);
  else mount();

  window.CVSettings = {
    tool: TOOL, use: use, store: store, link: link, open: open, close: close,
    sets: function () { return sets.map(function (e) { return { id: e.id, label: e.label, overridden: !!e.overridden }; }); },
    get: function (id) { for (var i = 0; i < sets.length; i++) if (sets[i].id === id) return currentValue(sets[i]); },
    _toRows: function (id) { for (var i = 0; i < sets.length; i++) if (sets[i].id === id) return toRows(sets[i], currentValue(sets[i])); },
    _fromRows: function (id, rows) { for (var i = 0; i < sets.length; i++) if (sets[i].id === id) return fromRows(sets[i], rows); }
  };
})();
