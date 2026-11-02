/* global S, $, say, fmt, unit, draw, requestDraw, setTool, fit, layerStats, send, sendCount, writeTable, layerRows, takeoffRows,
   updateCheck, undo, redo, openDxf, openPdf, showMethod, UNIT_M, renderTakeoff, renderScaleList, saveSoon, SHAPES, COLS,
   savePrefs, normalizeColumns, useSelectedAsFixed, updateLinks, insertPicture, toolFinish, toolBackspace, toolCancel */

/* ---------- files ---------- */
$('file').addEventListener('change', async (ev) => {
  const file = ev.target.files[0];
  ev.target.value = ''; // lets the same file be opened again
  if (!file) return;
  if (/\.pdf$/i.test(file.name)) await openPdf(file);
  else if (/\.dxf$/i.test(file.name)) await openDxf(file);
  else say('Open a .dxf or .pdf file. DWG files need to be saved as DXF first.');
});

/* ---------- tools ---------- */
document.querySelectorAll('[data-tool]').forEach((b) => b.addEventListener('click', () => setTool(b.dataset.tool)));
$('fit').addEventListener('click', () => fit());
$('finish').addEventListener('click', () => toolFinish());
$('back').addEventListener('click', () => toolBackspace());
$('cancel').addEventListener('click', () => { toolCancel(); say('Cancelled.'); });
$('undo').addEventListener('click', undo);
$('redo').addEventListener('click', redo);

/* ---------- distance + scale check ---------- */
$('measure-send').addEventListener('click', () => {
  const m = S.lastDist;
  if (m && m.f) send('Distance', new Map([[m.f, m.raw]]), 1, unit());
});
$('chk-val').addEventListener('input', updateCheck);
$('chk-unit').addEventListener('change', updateCheck);
$('chk-cal').addEventListener('click', () => {
  const exp = parseFloat($('chk-val').value);
  const m = S.lastDist;
  if (!(exp > 0) || !m) { say('Enter the expected real distance first.'); return; }
  const u = $('chk-unit').value;
  S.calibration = { factor: exp / m.raw, unit: u, label: `Calibrated (${exp} ${u} checked)` };
  $('cal-result').textContent = `1 drawing unit = ${+S.calibration.factor.toPrecision(6)} ${u}`;
  showMethod('calibrate');
  say('Distance recorded as a calibration. Choose where to apply it under Scale, then click Apply scale.');
});

/* ---------- send everything ---------- */
$('send-takeoff').addEventListener('click', () => writeTable(takeoffRows()));
$('send-layers').addEventListener('click', () => writeTable(layerRows()));

/* ---------- layers ---------- */
function chip(text, onClick) {
  const b = document.createElement('button');
  b.className = 'btn';
  b.textContent = text;
  b.addEventListener('click', onClick);
  return b;
}

const MAX_LAYER_ROWS = 200;

function renderLayers() {
  const list = $('layers');
  list.replaceChildren();
  const u = unit();
  const filter = $('layer-filter').value.trim().toLowerCase();
  const stats = [...layerStats()].filter(([name]) => !filter || name.toLowerCase().includes(filter));
  $('layers-empty').hidden = stats.length > 0;
  $('send-layers').hidden = !S.items.length;

  for (const [name, s] of stats.slice(0, MAX_LAYER_ROWS)) {
    const li = document.createElement('li');
    li.addEventListener('pointerenter', () => { S.focus = name; requestDraw(); });
    li.addEventListener('pointerleave', () => { S.focus = null; requestDraw(); });

    const label = document.createElement('label');
    const box = document.createElement('input');
    box.type = 'checkbox';
    box.checked = !S.hidden.has(name);
    box.addEventListener('change', () => { if (box.checked) S.hidden.delete(name); else S.hidden.add(name); requestDraw(); });
    label.append(box, document.createTextNode(name));
    li.append(label);

    const chips = document.createElement('div');
    chips.className = 'chips';
    if (s.len > 0) chips.append(chip(`Length ${fmt(s.len)} ${u}`, () => send(name, s.lenT, 1, u)));
    if (s.area > 0) chips.append(chip(`Area ${fmt(s.area)} ${u}²`, () => send(name, s.areaT, 2, `${u}²`)));
    if (s.count > 0) chips.append(chip(`Count ${s.count}`, () => sendCount(name, s.count)));
    li.append(chips);

    if (s.unscaled > 0) {
      const warn = document.createElement('p');
      warn.className = 'warn';
      warn.textContent = `${s.unscaled} object${s.unscaled > 1 ? 's have' : ' has'} no scale and ${s.unscaled > 1 ? 'are' : 'is'} left out.`;
      li.append(warn);
    }
    list.append(li);
  }
  if (stats.length > MAX_LAYER_ROWS) {
    const more = document.createElement('li');
    more.className = 'none';
    more.textContent = `Showing ${MAX_LAYER_ROWS} of ${stats.length} layers. Use the filter to narrow the list.`;
    list.append(more);
  }
}
$('layer-filter').addEventListener('input', renderLayers);

/* ---------- shapes ---------- */
for (const kind of ['length', 'area', 'count']) {
  const sel = $(`shape-${kind}`);
  for (const [key, def] of Object.entries(SHAPES[kind])) {
    const o = document.createElement('option');
    o.value = key;
    o.textContent = def.label;
    sel.append(o);
  }
  sel.value = S.shape[kind];
  sel.addEventListener('change', () => { S.shape[kind] = sel.value; savePrefs(); setTool(kind); });
}
$('show-labels').addEventListener('change', requestDraw);

/* ---------- where quantities go ---------- */
const picks = $('col-picks');
for (const [key, text] of COLS) {
  const l = document.createElement('label');
  l.className = 'check';
  const cb = document.createElement('input');
  cb.type = 'checkbox';
  cb.dataset.col = key;
  if (key === 'qty') cb.disabled = true;
  cb.addEventListener('change', () => {
    const chosen = [...picks.querySelectorAll('[data-col]')].filter((x) => x.checked).map((x) => x.dataset.col);
    S.prefs.columns = normalizeColumns(chosen);
    savePrefs();
    syncDestUi();
  });
  l.append(cb, document.createTextNode(text));
  picks.append(l);
}

function syncDestUi() {
  const d = S.prefs.dest;
  $('dest-mode').value = d.mode;
  $('dest-fixed').hidden = d.mode !== 'fixed';
  $('dest-addr').value = d.address;
  $('dest-advance').value = d.advance;
  $('dest-write').value = d.write;
  $('dest-where').value = S.prefs.where;
  $('decimals').value = S.prefs.decimals;
  picks.querySelectorAll('[data-col]').forEach((cb) => { cb.checked = S.prefs.columns.includes(cb.dataset.col); });
}

$('dest-mode').addEventListener('change', () => { S.prefs.dest.mode = $('dest-mode').value; savePrefs(); syncDestUi(); });
$('dest-addr').addEventListener('change', () => { S.prefs.dest.address = $('dest-addr').value.trim(); savePrefs(); });
$('dest-use').addEventListener('click', useSelectedAsFixed);
$('dest-advance').addEventListener('change', () => { S.prefs.dest.advance = $('dest-advance').value; savePrefs(); });
$('dest-write').addEventListener('change', () => { S.prefs.dest.write = $('dest-write').value; savePrefs(); });
$('dest-where').addEventListener('change', () => { S.prefs.where = $('dest-where').value; savePrefs(); });
$('decimals').addEventListener('change', () => {
  S.prefs.decimals = Math.max(0, Math.min(4, parseInt($('decimals').value, 10) || 0));
  savePrefs();
  syncDestUi();
  renderTakeoff();
  renderLayers();
});
$('update-links').addEventListener('click', updateLinks);
$('insert-picture').addEventListener('click', insertPicture);

/* ---------- start ---------- */
syncDestUi();
setTool('pan');
renderScaleList();
renderLayers();
renderTakeoff();
