/* global Excel, Office, S, $, say, fmt, unit, toScreen, toWorld, ctx, W, H, requestDraw, setTool, fit, polyLen, bboxCenter,
   scaleAtPoint, UNIT_M, showMethod, showPage, drawSymbols, geomFor, markersAlong, simplify, angleDeg, polyDist, pointInPoly,
   segDist, SHAPES, syncDestUi, renderLayers */

const PALETTE = ['#e4570e', '#1b8a5a', '#2f6fdd', '#a23bd1', '#c9a100', '#d6336c', '#0b7285', '#7a5230'];
const KIND_LABEL = { length: 'Length', area: 'Area', count: 'Count', angle: 'Angle' };
const uid = () => Math.random().toString(36).slice(2, 9);
const num = (v, d) => (v === '' || v === null || v === undefined || !Number.isFinite(+v) ? d : +v);

/* ---------- preferences (per user, kept in the browser) ---------- */
const COLS = [['code', 'Code'], ['name', 'Description'], ['file', 'Drawing'], ['page', 'Page'], ['type', 'Type'],
  ['qty', 'Quantity'], ['unit', 'Unit'], ['rate', 'Rate'], ['amount', 'Amount'], ['notes', 'Notes']];
const DEFAULT_PREFS = {
  dest: { mode: 'selection', address: '', advance: 'down', write: 'row' },
  columns: ['name', 'page', 'type', 'qty', 'unit'],
  decimals: 2,
  where: 'here',
  shape: { length: 'polyline', area: 'polygon', count: 'markers' },
};

function normalizeColumns(cols) {
  const set = new Set(cols);
  set.add('qty');
  if (set.has('amount')) set.add('rate');
  return COLS.map((c) => c[0]).filter((k) => set.has(k));
}

function loadPrefs() {
  let p = {};
  try { p = JSON.parse(localStorage.getItem('takeoff:prefs') || '{}'); } catch { /* ignore */ }
  return {
    ...DEFAULT_PREFS, ...p,
    dest: { ...DEFAULT_PREFS.dest, ...(p.dest || {}) },
    shape: { ...DEFAULT_PREFS.shape, ...(p.shape || {}) },
    columns: normalizeColumns(p.columns || DEFAULT_PREFS.columns),
  };
}
function savePrefs() { try { localStorage.setItem('takeoff:prefs', JSON.stringify(S.prefs)); } catch { /* ignore */ } }
S.prefs = loadPrefs();
S.shape = S.prefs.shape;
S.edit = null;
S.pickHover = null;

/* ---------- undo / redo (snapshots of the takeoff list) ---------- */
const snapshot = () => JSON.stringify(S.takeoff);
function pushUndo() { S.undo.push(snapshot()); if (S.undo.length > 100) S.undo.shift(); S.redo = []; }
function afterChange() { renderTakeoff(); requestDraw(); saveSoon(); }
function commit(fn) { pushUndo(); fn(); afterChange(); }
function undo() {
  if (!S.undo.length) { say('Nothing to undo.'); return; }
  S.redo.push(snapshot());
  S.takeoff = JSON.parse(S.undo.pop());
  S.activeCount = null;
  afterChange();
}
function redo() {
  if (!S.redo.length) { say('Nothing to redo.'); return; }
  S.undo.push(snapshot());
  S.takeoff = JSON.parse(S.redo.pop());
  S.activeCount = null;
  afterChange();
}

/* ---------- items and their geometry ---------- */
const GEO = new WeakMap(); // cached outline + exact length/area per item; call touch() after changing pts or params
const touch = (t) => GEO.delete(t);

function addItem(kind, pts, extra = {}) {
  const n = S.takeoff.filter((t) => t.kind === kind).length + 1;
  const t = {
    id: uid(), kind, name: `${KIND_LABEL[kind]} ${n}`, page: S.page, pts,
    shape: kind === 'angle' ? 'angle' : S.shape[kind],
    color: PALETTE[S.takeoff.length % PALETTE.length], ...extra,
  };
  if (t.kind === 'count' && t.shape === 'spaced') { t.spacing = 3; t.endPost = true; t.dimUnit = unit(); }
  S.takeoff.push(t);
  return t;
}

function itemGeom(t) {
  let g = GEO.get(t);
  if (g) return g;
  if (t.exact) g = { poly: t.pts, closed: !!t.closed, len: t.exact.len, area: t.exact.area };
  else if (t.kind === 'count' || t.kind === 'angle') g = geomFor('length', 'polyline', t.pts);
  else g = geomFor(t.kind, t.shape || (t.kind === 'length' ? 'polyline' : 'polygon'), t.pts, { flip: t.flip });
  GEO.set(t, g);
  return g;
}

function deleteSelected() {
  if (!S.sel) return;
  commit(() => { S.takeoff = S.takeoff.filter((t) => t.id !== S.sel); S.sel = null; });
}

/* ---------- quantities ---------- */
function typeLabel(t) {
  if (t.kind === 'count') return 'Count';
  if (t.kind === 'angle') return 'Angle';
  if (t.out === 'volume') return 'Volume';
  if (t.kind === 'length') return t.out === 'area' ? 'Area' : 'Length';
  return t.out === 'perimeter' ? 'Length' : 'Area';
}

// Markers for a spaced count (needs the scale: the spacing is a real-world distance).
function markersFor(t) {
  const f = scaleAtPoint(bboxCenter(t.pts), t.page);
  if (f == null) return null;
  const step = (num(t.spacing, 0) * (UNIT_M[t.dimUnit || unit()] / UNIT_M[unit()])) / f;
  return markersAlong(itemGeom(t).poly, step, t.endPost !== false);
}

// Returns null (no scale here), { missing } (needs a width, depth or spacing) or { value, unit, expr }.
// expr is the Excel formula body, built from the same factors as value so the two always agree.
function qtyOf(t) {
  const u = unit();
  if (t.kind === 'angle') {
    if (t.pts.length < 3) return null;
    const v = angleDeg(t.pts[0], t.pts[1], t.pts[2]);
    return { value: v, unit: '°', expr: `${+v.toFixed(4)}` };
  }
  const base = [];
  let unitLabel;
  if (t.kind === 'count') {
    let n = t.pts.length;
    if (t.shape === 'spaced') {
      if (!(num(t.spacing, 0) > 0)) return { missing: 'spacing' };
      const m = markersFor(t);
      if (m === null) return null;
      n = m.length;
    }
    base.push({ n, s: `${n}` });
    unitLabel = 'ea';
  } else {
    if (t.pts.length < 2) return null;
    const g = itemGeom(t);
    const f = scaleAtPoint(bboxCenter(t.pts), t.page);
    if (f == null) return null;
    const fs = +f.toPrecision(8);
    const dimF = UNIT_M[t.dimUnit || u] / UNIT_M[u];
    const W_ = num(t.width, 0) * dimF;
    const D_ = num(t.depth, 0) * dimF;
    const dim = (v) => ({ n: v, s: `${+v.toPrecision(8)}` });
    const out = t.out || 'base';
    if (t.kind === 'length') {
      base.push({ n: g.len, s: `${+g.len.toFixed(4)}` }, { n: f, s: `${fs}` });
      unitLabel = u;
      if (out === 'area' || out === 'volume') { if (!(W_ > 0)) return { missing: 'width' }; base.push(dim(W_)); unitLabel = `${u}²`; }
      if (out === 'volume') { if (!(D_ > 0)) return { missing: 'depth' }; base.push(dim(D_)); unitLabel = `${u}³`; }
    } else if (out === 'perimeter') {
      base.push({ n: g.len, s: `${+g.len.toFixed(4)}` }, { n: f, s: `${fs}` });
      unitLabel = u;
    } else {
      base.push({ n: g.area, s: `${+g.area.toFixed(4)}` }, { n: f * f, s: `${fs}^2` });
      unitLabel = `${u}²`;
      if (out === 'volume') { if (!(D_ > 0)) return { missing: 'depth' }; base.push(dim(D_)); unitLabel = `${u}³`; }
    }
  }
  const mult = num(t.mult, 1);
  const waste = num(t.waste, 0);
  const extra = [];
  if (mult !== 1) extra.push({ n: mult, s: `${mult}` });
  if (waste !== 0) extra.push({ n: 1 + waste / 100, s: `(1+${+(waste / 100).toPrecision(6)})` });
  let value = [...base, ...extra].reduce((a, x) => a * x.n, 1);
  let expr = base.map((x) => x.s).join('*');
  if (extra.length) expr = `(${expr})*${extra.map((x) => x.s).join('*')}`;
  if (t.deduct) { value = -value; expr = `-${expr}`; }
  return { value, unit: unitLabel, expr };
}

// terms: Map(scale factor -> raw drawing-unit sum). Mixed scales become a sum, e.g. =120.5*0.05+80*0.1
function termsExpr(terms, power) {
  return [...terms]
    .filter(([, raw]) => raw > 0)
    .map(([f, raw]) => `${+raw.toFixed(4)}*${+f.toPrecision(8)}${power === 2 ? '^2' : ''}`)
    .join('+');
}

/* ---------- Excel: addresses and destination ---------- */
const safeText = (t) => (/^[=+\-@]/.test(t) ? `'${t}` : t);

function splitAddress(a) {
  const i = a.lastIndexOf('!');
  if (i < 0) return { sheet: null, ref: a.trim() };
  let sheet = a.slice(0, i);
  if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
  return { sheet, ref: a.slice(i + 1).trim() };
}

function rangeAt(c, addr) {
  const { sheet, ref } = splitAddress(addr);
  const ws = sheet ? c.workbook.worksheets.getItem(sheet) : c.workbook.worksheets.getActiveWorksheet();
  return ws.getRange(ref);
}

async function inExcel(fn) {
  if (!S.inExcel) { say('Open this pane inside Excel to send quantities.'); return; }
  try {
    await Excel.run(async (c) => { say(await fn(c)); });
  } catch (err) {
    say(`Excel could not do that: ${err.message}`);
  }
}

const anchorCell = (c) => {
  const d = S.prefs.dest;
  return d.mode === 'fixed' && d.address.trim()
    ? rangeAt(c, d.address.trim()).getCell(0, 0)
    : c.workbook.getSelectedRange().getCell(0, 0);
};

const ADVANCE = { down: [1, 0], right: [0, 1], none: [0, 0] };

// Moves the destination past what was just written (the selection, or the fixed address).
async function advanceFrom(c, anchor, rows, cols) {
  const d = S.prefs.dest;
  const [dr, dc] = ADVANCE[d.advance] || [0, 0];
  if (dr || dc) {
    const next = anchor.getOffsetRange(dr ? rows : 0, dc ? cols : 0);
    if (d.mode === 'fixed') {
      next.load('address');
      await c.sync();
      d.address = next.address;
      savePrefs();
      if (typeof syncDestUi === 'function') syncDestUi();
    } else {
      next.select();
    }
  }
  await c.sync();
}

// One quantity. Either "label | quantity | unit" or just the quantity, depending on the Write setting.
function writeRow(label, formula, unitText) {
  return inExcel(async (c) => {
    const first = anchorCell(c);
    first.load('address');
    let cols = 3;
    if (S.prefs.dest.write === 'value') { first.formulas = [[formula]]; cols = 1; }
    else first.getResizedRange(0, 2).formulas = [[safeText(label), formula, unitText]];
    await advanceFrom(c, first, 1, cols);
    return `Sent “${label}” to ${first.address.split('!').pop()}.`;
  });
}

// Quantity only, into the exact cell an item is linked to.
function writeLinked(t, q) {
  return inExcel(async (c) => {
    rangeAt(c, t.link.address).getCell(0, 0).formulas = [[`=${q.expr}`]];
    await c.sync();
    return `Updated ${t.link.address}.`;
  });
}

function buildTable(rows, prefs) {
  const cols = normalizeColumns(prefs.columns);
  const idx = Object.fromEntries(cols.map((k, i) => [k, i]));
  const dec = Math.max(0, Math.min(4, prefs.decimals | 0));
  const qtyFmt = dec ? `#,##0.${'0'.repeat(dec)}` : '#,##0';
  const header = cols.map((k) => COLS.find((c) => c[0] === k)[1]);
  const body = rows.map((r) => cols.map((k, i) => {
    switch (k) {
      case 'code': return safeText(r.code || '');
      case 'name': return safeText(r.name);
      case 'file': return safeText(r.file || '');
      case 'page': return r.page;
      case 'type': return r.type;
      case 'qty': return r.formula;
      case 'unit': return r.unit;
      case 'rate': return r.rate === undefined || r.rate === null || r.rate === '' ? '' : r.rate;
      case 'amount': return r.rate === undefined || r.rate === null || r.rate === '' ? '' : `=RC[${idx.qty - i}]*RC[${idx.rate - i}]`;
      default: return safeText(r.notes || '');
    }
  }));
  const formats = rows.map((r) => cols.map((k) => {
    if (k === 'page') return '0';
    if (k === 'qty') return r.type === 'Count' ? '0' : r.type === 'Angle' ? '0.0' : qtyFmt;
    return k === 'rate' || k === 'amount' ? '#,##0.00' : 'General';
  }));
  return { cols, idx, header, body, formats };
}

// A header plus rows, optionally on a fresh "Takeoff" sheet, with a total under Amount when rates are used.
function writeTable(rows) {
  if (!rows.length) { say('Nothing to send yet. Add items or set a scale first.'); return Promise.resolve(); }
  return inExcel(async (c) => {
    const t = buildTable(rows, S.prefs);
    const n = rows.length;
    const w = t.header.length;
    let anchor;
    let ws = null;
    if (S.prefs.where === 'newsheet') {
      const sheets = c.workbook.worksheets;
      sheets.load('items/name');
      await c.sync();
      const names = new Set(sheets.items.map((s) => s.name));
      let name = 'Takeoff';
      let k = 1;
      while (names.has(name)) name = `Takeoff ${++k}`;
      ws = sheets.add(name);
      const title = ws.getRange('A1');
      title.values = [['Quantity takeoff']];
      title.format.font.bold = true;
      title.format.font.size = 14;
      ws.getRange('A2').values = [[`${S.fileName || 'Drawing'}  ·  ${new Date().toLocaleDateString()}`]];
      anchor = ws.getRange('A4');
    } else {
      anchor = anchorCell(c);
    }
    anchor.load('address');

    const head = anchor.getResizedRange(0, w - 1);
    head.values = [t.header];
    head.format.font.bold = true;
    head.format.font.color = '#FFFFFF';
    head.format.fill.color = '#17232B';
    const body = anchor.getOffsetRange(1, 0).getResizedRange(n - 1, w - 1);
    body.formulasR1C1 = t.body;
    body.numberFormat = t.formats;

    let used = n + 1;
    if (t.idx.amount !== undefined) {
      const labelCell = anchor.getOffsetRange(n + 1, t.idx.name !== undefined ? t.idx.name : 0);
      labelCell.values = [['Total']];
      labelCell.format.font.bold = true;
      const tot = anchor.getOffsetRange(n + 1, t.idx.amount);
      tot.formulasR1C1 = [[`=SUM(R[-${n}]C:R[-1]C)`]];
      tot.numberFormat = [['#,##0.00']];
      tot.format.font.bold = true;
      used += 1;
    }
    anchor.getResizedRange(used - 1, w - 1).format.autofitColumns();
    if (ws) {
      if (Office.context.requirements.isSetSupported('ExcelApi', '1.7')) ws.freezePanes.freezeRows(4);
      ws.activate();
    } else {
      await advanceFrom(c, anchor, used + 1, w + 1);
    }
    await c.sync();
    return `Sent ${n} row${n > 1 ? 's' : ''} to ${anchor.address}.`;
  });
}

function send(label, terms, power, unitText) {
  const expr = termsExpr(terms, power);
  if (!expr) { say('Nothing to send yet. Set a scale first.'); return; }
  return writeRow(label, `=${expr}`, unitText);
}
const sendCount = (label, n) => writeRow(label, `=${n}`, 'ea');

function rowOf(t) {
  const q = qtyOf(t);
  if (!q || q.missing) return null;
  return { name: t.name, code: t.code, file: S.fileName, page: t.page, type: typeLabel(t), formula: `=${q.expr}`, unit: q.unit, rate: t.rate, notes: t.notes };
}
const takeoffRows = () => S.takeoff.map(rowOf).filter(Boolean);

function layerRows() {
  const rows = [];
  const u = unit();
  for (const [name, s] of layerStats()) {
    const add = (suffix, type, formula, unitText) => rows.push({ name: `${name} (${suffix})`, file: S.fileName, page: S.page, type, formula, unit: unitText });
    if (s.len > 0) add('length', 'Length', `=${termsExpr(s.lenT, 1)}`, u);
    if (s.area > 0) add('area', 'Area', `=${termsExpr(s.areaT, 2)}`, `${u}²`);
    if (s.count > 0) add('count', 'Count', `=${s.count}`, 'ea');
  }
  return rows;
}

/* ---------- Excel: links, picture, destination helpers ---------- */
function sendItem(t) {
  const q = qtyOf(t);
  if (!q) { say('Set a scale for this spot first.'); return; }
  if (q.missing) { say(`Open the item's details and enter a ${q.missing}.`); return; }
  if (t.link) writeLinked(t, q); else writeRow(t.name, `=${q.expr}`, q.unit);
}

function linkItem(t) {
  return inExcel(async (c) => {
    const r = c.workbook.getSelectedRange().getCell(0, 0);
    r.load('address');
    await c.sync();
    commit(() => { t.link = { address: r.address }; });
    return `Linked “${t.name}” to ${r.address}. Use Send or Update linked cells to write its quantity there.`;
  });
}

function updateLinks() {
  const linked = S.takeoff.filter((t) => t.link);
  if (!linked.length) { say('No items are linked to cells yet. Use “Link to selected cell” on an item.'); return Promise.resolve(); }
  return inExcel(async (c) => {
    let done = 0;
    let skipped = 0;
    for (const t of linked) {
      const q = qtyOf(t);
      if (!q || q.missing) { skipped++; continue; }
      rangeAt(c, t.link.address).getCell(0, 0).formulas = [[`=${q.expr}`]];
      done++;
    }
    await c.sync();
    return `Updated ${done} linked cell${done === 1 ? '' : 's'}${skipped ? `, skipped ${skipped} without a scale or missing size` : ''}.`;
  });
}

function useSelectedAsFixed() {
  return inExcel(async (c) => {
    const r = c.workbook.getSelectedRange().getCell(0, 0);
    r.load('address');
    await c.sync();
    S.prefs.dest.mode = 'fixed';
    S.prefs.dest.address = r.address;
    savePrefs();
    if (typeof syncDestUi === 'function') syncDestUi();
    return `Quantities will go to ${r.address}.`;
  });
}

function insertPicture() {
  if (S.inExcel && !Office.context.requirements.isSetSupported('ExcelApi', '1.9')) {
    say('Inserting pictures needs a newer version of Excel.');
    return Promise.resolve();
  }
  return inExcel(async (c) => {
    const b64 = document.getElementById('plan').toDataURL('image/png').split(',')[1];
    const anchor = anchorCell(c);
    const shape = anchor.worksheet.shapes.addImage(b64);
    shape.name = `Takeoff markup ${new Date().toLocaleTimeString()}`;
    try {
      anchor.load(['left', 'top']);
      await c.sync();
      shape.left = anchor.left;
      shape.top = anchor.top;
    } catch { /* older Excel: the picture lands at the default position */ }
    await c.sync();
    return 'Inserted the current view as a picture.';
  });
}

/* ---------- tools ---------- */
const HINTS_SHAPE = {
  line: 'Click the start and end, or drag.',
  polyline: 'Click along the run. Double-click or press Enter to finish.',
  arc: 'Click the start, a point on the arc, then the end.',
  bezier: 'Click the start, two control points, then the end.',
  spline: 'Click points the curve passes through. Double-click or press Enter to finish.',
  freehand: 'Press and drag to trace. Release to finish.',
  polygon: 'Click the corners. Double-click or press Enter to close the shape.',
  rect: 'Click two opposite corners, or drag.',
  rrect: 'Click both ends of one side, then a point that sets the width.',
  circle: 'Click the centre, then a point on the edge, or drag.',
  ellipse: 'Click two opposite corners of its bounding box, or drag.',
  sector: 'Click the centre, the start of the arc, then the end (counter-clockwise).',
  segment: 'Click the start, a point on the arc, then the end.',
  markers: 'Click each item to count. Enter ends the group.',
  spaced: 'Click along the path, then press Enter. Set the spacing in the item details.',
};
function hintFor(tool) {
  if (SHAPES[tool]) return HINTS_SHAPE[S.shape[tool]];
  return {
    distance: 'Click two points, or drag, to measure a distance.',
    angle: 'Click a point on one arm, the corner, then a point on the other arm.',
    pick: 'Click a line, or inside a closed shape, to measure it exactly from the drawing.',
    symbol: 'Drag a box around one example of the symbol.',
    calibrate: 'Click two points a known distance apart.',
  }[tool] || '';
}

const curShape = () => (SHAPES[S.tool] ? SHAPES[S.tool][S.shape[S.tool]] : null);
const isDragShape = () => !!(curShape() && curShape().drag);
const isTwoPoint = () => S.tool === 'distance' || (!!curShape() && curShape().need === 2);
const needsFinishBar = () => !!(curShape() && curShape().open);

function finishShape() {
  const kind = S.tool;
  const pts = S.pts.slice();
  commit(() => { S.sel = addItem(kind, pts).id; });
  S.pts = [];
  S.hover = null;
}

function toolClick(w) {
  if (S.tool === 'pick') { pickClick(w); return; }
  switch (S.tool) {
    case 'distance':
      if (S.pts.length >= 2) S.pts = [];
      S.pts.push(w);
      if (S.pts.length === 2) showDistance(); else $('measure-out').hidden = true;
      break;
    case 'calibrate':
      if (S.pts.length >= 2) S.pts = [];
      S.pts.push(w);
      if (S.pts.length === 2) {
        S.calRaw = Math.hypot(S.pts[1][0] - S.pts[0][0], S.pts[1][1] - S.pts[0][1]);
        $('cal-panel').hidden = false;
        $('cal-dist').focus();
      } else {
        $('cal-panel').hidden = true;
      }
      break;
    case 'angle':
      if (S.pts.length >= 3) S.pts = [];
      S.pts.push(w);
      if (S.pts.length === 3) { commit(() => { S.sel = addItem('angle', S.pts.slice()).id; }); S.pts = []; }
      break;
    case 'count':
      if (S.shape.count === 'markers') {
        pushUndo();
        let t = S.takeoff.find((x) => x.id === S.activeCount);
        if (!t) { t = addItem('count', []); S.activeCount = t.id; }
        t.pts.push(w);
        touch(t);
        afterChange();
        return;
      }
      S.pts.push(w);
      break;
    case 'length':
    case 'area': {
      S.pts.push(w);
      const sh = curShape();
      if (sh && sh.need && S.pts.length === sh.need) finishShape();
      break;
    }
    default: break;
  }
  requestDraw();
}

function toolDouble() {
  const sh = curShape();
  if (!sh || !sh.open || S.tool === 'count' && S.shape.count === 'markers') return;
  const [a, b] = [S.pts.at(-1), S.pts.at(-2)];
  if (a && b && a[0] === b[0] && a[1] === b[1]) S.pts.pop(); // the two clicks of the double-click
  toolFinish();
}

function toolFinish() {
  if (S.tool === 'count' && S.shape.count === 'markers') {
    if (S.activeCount) say('Count group finished. The next click starts a new group.');
    S.activeCount = null;
    return;
  }
  const sh = curShape();
  if (!sh || !sh.open) return;
  if (S.pts.length < sh.min) { say(`Add at least ${sh.min} points first.`); return; }
  finishShape();
  if (S.tool === 'count') say('Set the spacing in the item details below, then send the count.');
}

function toolCancel() {
  S.pts = [];
  S.activeCount = null;
  S.hover = null;
  $('measure-out').hidden = true;
  $('cal-panel').hidden = true;
  requestDraw();
}

function toolBackspace() {
  if (S.tool === 'count' && S.shape.count === 'markers') { if (S.activeCount) undo(); return; }
  S.pts.pop();
  if (S.pts.length < 2) { $('measure-out').hidden = true; $('cal-panel').hidden = true; }
  requestDraw();
}

function addFreePoint(w) {
  const last = S.pts.at(-1);
  if (!last || Math.hypot(w[0] - last[0], w[1] - last[1]) * S.view.k > 3) { S.pts.push(w); requestDraw(); }
}

function finishFreehand() {
  const need = S.tool === 'area' ? 3 : 2;
  const pts = simplify(S.pts, 1.5 / S.view.k);
  S.pts = [];
  if (pts.length < need) { requestDraw(); return; }
  commit(() => { S.sel = addItem(S.tool, pts).id; });
}

/* ---------- pick: measure an existing vector shape exactly ---------- */
function itemBox(it) {
  if (it.bb) return it.bb;
  if (it.type === 'path') {
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    for (const p of it.pts) { if (p[0] < x0) x0 = p[0]; if (p[1] < y0) y0 = p[1]; if (p[0] > x1) x1 = p[0]; if (p[1] > y1) y1 = p[1]; }
    it.bb = [x0, y0, x1, y1];
  } else if (it.type === 'insert') it.bb = [it.p[0], it.p[1], it.p[0], it.p[1]];
  else it.bb = [it.c[0] - it.r, it.c[1] - it.r, it.c[0] + it.r, it.c[1] + it.r];
  return it.bb;
}

function circlePoly(c, r) {
  return Array.from({ length: 96 }, (_, i) => [c[0] + r * Math.cos((2 * Math.PI * i) / 96), c[1] + r * Math.sin((2 * Math.PI * i) / 96)]);
}

function arcPoly(it) {
  const sweep = (((it.a1 - it.a0) % (2 * Math.PI)) + 2 * Math.PI) % (2 * Math.PI);
  return Array.from({ length: 49 }, (_, i) => [it.c[0] + it.r * Math.cos(it.a0 + (sweep * i) / 48), it.c[1] + it.r * Math.sin(it.a0 + (sweep * i) / 48)]);
}

function pickTarget(w) {
  const tol = 10 / S.view.k;
  let best = null;
  let bd = tol;
  let inside = null;
  for (const it of S.items) {
    if (it.type === 'insert' || S.hidden.has(it.layer)) continue;
    const [x0, y0, x1, y1] = itemBox(it);
    if (w[0] < x0 - tol || w[0] > x1 + tol || w[1] < y0 - tol || w[1] > y1 + tol) continue;
    let d;
    if (it.type === 'path') d = polyDist(it.pts, w, it.closed);
    else if (it.type === 'circle') d = Math.abs(Math.hypot(w[0] - it.c[0], w[1] - it.c[1]) - it.r);
    else d = polyDist(arcPoly(it), w, false);
    if (d < bd) { bd = d; best = it; }
    const closed = it.type === 'circle' || (it.type === 'path' && it.closed);
    if (closed && it.area > 0) {
      const hit = it.type === 'circle' ? Math.hypot(w[0] - it.c[0], w[1] - it.c[1]) <= it.r : pointInPoly(it.pts, w);
      if (hit && (!inside || it.area < inside.area)) inside = it;
    }
  }
  return best || inside;
}

function pickClick(w) {
  const it = pickTarget(w);
  if (!it) { say(S.items.length ? 'Nothing to pick there. Click right on a line or inside a closed shape.' : 'This page has no vector lines to pick. Use the measuring tools instead.'); return; }
  const closed = it.type === 'circle' || (it.type === 'path' && it.closed);
  const pts = it.type === 'path' ? it.pts.map((p) => p.slice()) : it.type === 'circle' ? circlePoly(it.c, it.r) : arcPoly(it);
  commit(() => {
    const t = addItem(closed ? 'area' : 'length', pts, { shape: 'trace', closed, exact: { len: it.len, area: it.area }, name: it.layer });
    S.sel = t.id;
  });
  say(`Picked from layer “${it.layer}”. Lengths and areas come straight from the drawing.`);
}

function updatePickHover(w) {
  if (S.items.length > 50000) { S.pickHover = null; return; }
  S.pickHover = pickTarget(w);
}

/* ---------- selecting and editing items on the canvas ---------- */
function hitTakeoff(sx, sy) {
  const p = toWorld(sx, sy);
  const tol = 8 / S.view.k;
  for (let i = S.takeoff.length - 1; i >= 0; i--) {
    const t = S.takeoff[i];
    if (t.page !== S.page || !t.pts.length) continue;
    if (t.kind === 'count') {
      const pts = t.shape === 'spaced' ? (markersFor(t) || t.pts) : t.pts;
      if (pts.some((q) => Math.hypot(q[0] - p[0], q[1] - p[1]) <= tol)) return t.id;
    } else {
      const g = itemGeom(t);
      if (polyDist(g.poly, p, g.closed) <= tol || (g.closed && pointInPoly(g.poly, p))) return t.id;
    }
  }
  return null;
}

function selectAt(sx, sy) {
  const id = hitTakeoff(sx, sy);
  if (id === S.sel) return;
  S.sel = id;
  renderTakeoff();
  requestDraw();
}

// Control points of the selected item can be dragged when the Pan tool is active.
function handleAt(sx, sy) {
  const t = S.takeoff.find((x) => x.id === S.sel);
  if (!t || t.page !== S.page || t.exact || t.pts.length > 40) return null;
  for (let i = 0; i < t.pts.length; i++) {
    const [x, y] = toScreen(t.pts[i][0], t.pts[i][1]);
    if (Math.hypot(x - sx, y - sy) <= 10) return { id: t.id, i };
  }
  return null;
}
function beginEdit(h) { S.edit = { ...h, snap: snapshot(), moved: false }; }
function moveHandle(w) {
  const t = S.takeoff.find((x) => x.id === S.edit.id);
  if (!t) return;
  t.pts[S.edit.i] = w;
  S.edit.moved = true;
  touch(t);
  requestDraw();
}
function endEdit() {
  const e = S.edit;
  S.edit = null;
  if (!e || !e.moved) return;
  S.undo.push(e.snap);
  S.redo = [];
  afterChange();
}

/* ---------- distance + scale check ---------- */
function showDistance() {
  const [a, b] = S.pts;
  const raw = Math.hypot(b[0] - a[0], b[1] - a[1]);
  const f = scaleAtPoint([(a[0] + b[0]) / 2, (a[1] + b[1]) / 2]);
  S.lastDist = { raw, f };
  $('measure-text').textContent = f ? `Distance ${fmt(raw * f)} ${unit()}` : 'No scale set here yet';
  $('measure-send').disabled = !f;
  $('chk-val').value = '';
  $('chk-out').textContent = '';
  $('measure-out').hidden = false;
}

function updateCheck() {
  const exp = parseFloat($('chk-val').value);
  const m = S.lastDist;
  const out = $('chk-out');
  if (!m || !(exp > 0)) { out.textContent = ''; return; }
  if (!m.f) { out.textContent = 'No scale here yet. Use “Calibrate from this”.'; return; }
  const expectedOut = (exp * UNIT_M[$('chk-unit').value]) / UNIT_M[unit()];
  const pct = ((m.raw * m.f - expectedOut) / expectedOut) * 100;
  out.textContent = Math.abs(pct) < 0.5
    ? `Scale checks out (${pct.toFixed(2)}% off).`
    : `Measured ${fmt(m.raw * m.f)} ${unit()}: ${Math.abs(pct).toFixed(1)}% ${pct > 0 ? 'over' : 'under'} the expected value.`;
}

/* ---------- overlays on the canvas ---------- */
function label(text, x, y, color) {
  ctx.font = '12px "Segoe UI", system-ui, sans-serif';
  const w = ctx.measureText(text).width + 10;
  const bx = Math.min(Math.max(2, x), Math.max(2, W - w - 2));
  const by = Math.min(Math.max(2, y), H - 20);
  ctx.fillStyle = 'rgba(23,35,43,0.88)';
  ctx.fillRect(bx, by, w, 18);
  ctx.fillStyle = color || '#fff';
  ctx.textBaseline = 'middle';
  ctx.fillText(text, bx + 5, by + 9);
}

function pathOf(poly, closed) {
  ctx.beginPath();
  poly.forEach((p, i) => { const [x, y] = toScreen(p[0], p[1]); if (i) ctx.lineTo(x, y); else ctx.moveTo(x, y); });
  if (closed) ctx.closePath();
}

function previewLabel(kind, g) {
  const f = scaleAtPoint(bboxCenter(g.poly));
  if (f == null) return '';
  const u = unit();
  if (g.closed) return `${fmt(g.area * f * f)} ${u}²  ·  ${fmt(g.len * f)} ${u}`;
  return `${fmt(g.len * f)} ${u}`;
}

function drawOverlays(flag) {
  ctx.lineJoin = 'round';
  const showLabels = $('show-labels').checked;

  if (S.pickHover && S.tool === 'pick') {
    const it = S.pickHover;
    ctx.strokeStyle = flag;
    ctx.lineWidth = 3;
    pathOf(it.type === 'path' ? it.pts : it.type === 'circle' ? circlePoly(it.c, it.r) : arcPoly(it), it.type === 'circle' || (it.type === 'path' && it.closed));
    ctx.stroke();
  }

  for (const t of S.takeoff) {
    if (t.page !== S.page || !t.pts.length) continue;
    const sel = t.id === S.sel;
    ctx.strokeStyle = t.color;
    ctx.fillStyle = t.color;
    ctx.lineWidth = sel ? 3.5 : 2;
    if (t.kind === 'count') {
      const pts = t.shape === 'spaced' ? (markersFor(t) || []) : t.pts;
      if (t.shape === 'spaced') { pathOf(itemGeom(t).poly, false); ctx.setLineDash([4, 4]); ctx.stroke(); ctx.setLineDash([]); }
      for (const p of pts) {
        const [x, y] = toScreen(p[0], p[1]);
        ctx.beginPath();
        ctx.arc(x, y, sel ? 7 : 6, 0, 2 * Math.PI);
        ctx.globalAlpha = 0.85;
        ctx.fill();
        ctx.globalAlpha = 1;
        ctx.strokeStyle = '#fff';
        ctx.lineWidth = 1.5;
        ctx.stroke();
        ctx.strokeStyle = t.color;
      }
      if (showLabels && pts.length) { const [x, y] = toScreen(pts[0][0], pts[0][1]); label(`${t.name}: ${pts.length}`, x + 10, y - 9, '#fff'); }
      continue;
    }
    if (t.kind === 'angle') {
      pathOf(t.pts, false);
      ctx.stroke();
      const q = qtyOf(t);
      if (q) { const [x, y] = toScreen(t.pts[1][0], t.pts[1][1]); label(`${fmt(q.value)}°`, x + 8, y - 22, '#fff'); }
      continue;
    }
    const g = itemGeom(t);
    pathOf(g.poly, g.closed);
    if (g.closed) {
      ctx.globalAlpha = t.deduct ? 0.1 : 0.2;
      ctx.fill();
      ctx.globalAlpha = 1;
      if (t.deduct) ctx.setLineDash([5, 4]);
    }
    ctx.stroke();
    ctx.setLineDash([]);
    if (showLabels) {
      const q = qtyOf(t);
      if (q && !q.missing) {
        const c = toScreen(...bboxCenter(g.poly));
        label(`${t.name}: ${fmt(q.value)} ${q.unit}`, c[0] - 30, c[1] - 9, '#fff');
      }
    }
  }

  // handles on the selected item
  const st = S.takeoff.find((x) => x.id === S.sel);
  if (st && st.page === S.page && !st.exact && st.pts.length <= 40 && S.tool === 'pan') {
    ctx.fillStyle = '#fff';
    ctx.strokeStyle = st.color;
    ctx.lineWidth = 2;
    for (const p of st.pts) { const [x, y] = toScreen(p[0], p[1]); ctx.fillRect(x - 4, y - 4, 8, 8); ctx.strokeRect(x - 4, y - 4, 8, 8); }
  }

  // shape being drawn, previewed with the cursor as its next point
  if (S.pts.length && ['length', 'area', 'count', 'distance', 'calibrate', 'angle'].includes(S.tool) && !(S.tool === 'count' && S.shape.count === 'markers')) {
    const free = S.drag && S.drag.mode === 'free';
    const live = S.hover && !free && !(S.tool === 'angle' && S.pts.length >= 3) && !((S.tool === 'distance' || S.tool === 'calibrate') && S.pts.length >= 2);
    const pts = live ? [...S.pts, S.hover] : S.pts;
    ctx.strokeStyle = flag;
    ctx.fillStyle = flag;
    ctx.lineWidth = 2;
    let g;
    if (SHAPES[S.tool]) g = geomFor(S.tool, S.shape[S.tool], pts, {});
    else g = { poly: pts, closed: false, len: polyLen(pts, false), area: 0 };
    pathOf(g.poly, g.closed);
    if (g.closed) { ctx.globalAlpha = 0.12; ctx.fill(); ctx.globalAlpha = 1; }
    ctx.stroke();
    for (const p of S.pts) { const [x, y] = toScreen(p[0], p[1]); ctx.beginPath(); ctx.arc(x, y, 4, 0, 2 * Math.PI); ctx.fill(); }
    if (S.hover && SHAPES[S.tool] && S.tool !== 'count') {
      const text = previewLabel(S.tool, g);
      if (text) { const [hx, hy] = toScreen(S.hover[0], S.hover[1]); label(text, hx + 12, hy + 10, '#fff'); }
    } else if (S.hover && S.tool === 'angle' && S.pts.length === 2) {
      const [hx, hy] = toScreen(S.hover[0], S.hover[1]);
      label(`${fmt(angleDeg(S.pts[0], S.pts[1], S.hover))}°`, hx + 12, hy + 10, '#fff');
    }
  }
  if (S.hover && S.tool !== 'pan') {
    const [x, y] = toScreen(S.hover[0], S.hover[1]);
    ctx.strokeStyle = flag;
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.arc(x, y, 7, 0, 2 * Math.PI);
    ctx.stroke();
  }
  if (typeof drawSymbols === 'function') drawSymbols(flag);
}

/* ---------- takeoff list UI ---------- */
function mkBtn(text, onClick, labelText) {
  const b = document.createElement('button');
  b.className = 'btn';
  b.textContent = text;
  if (labelText) b.setAttribute('aria-label', labelText);
  b.addEventListener('click', onClick);
  return b;
}

function field(text, input) {
  const l = document.createElement('label');
  l.className = 'fld';
  l.append(document.createTextNode(text), input);
  return l;
}
function numInput(t, key, step = 'any', min) {
  const i = document.createElement('input');
  i.type = 'number';
  i.step = step;
  if (min !== undefined) i.min = min;
  i.value = t[key] === undefined || t[key] === null ? '' : t[key];
  i.addEventListener('change', () => commit(() => { if (i.value === '') delete t[key]; else t[key] = +i.value; touch(t); }));
  return i;
}
function textInput(t, key) {
  const i = document.createElement('input');
  i.type = 'text';
  i.value = t[key] || '';
  i.addEventListener('change', () => commit(() => { t[key] = i.value.trim(); }));
  return i;
}
function selectInput(options, value, onChange) {
  const s = document.createElement('select');
  for (const [v, text] of options) { const o = document.createElement('option'); o.value = v; o.textContent = text; s.append(o); }
  s.value = value;
  s.addEventListener('change', () => onChange(s.value));
  return s;
}
function checkInput(text, checked, onChange) {
  const l = document.createElement('label');
  l.className = 'check';
  const c = document.createElement('input');
  c.type = 'checkbox';
  c.checked = checked;
  c.addEventListener('change', () => onChange(c.checked));
  l.append(c, document.createTextNode(text));
  return l;
}

function itemDetails(t) {
  const d = document.createElement('details');
  d.className = 'more';
  const sum = document.createElement('summary');
  sum.textContent = 'Details';
  d.append(sum);
  const grid = document.createElement('div');
  grid.className = 'grid';
  const unitOpts = Object.keys(UNIT_M).map((u) => [u, u]);
  const dimUnit = () => selectInput(unitOpts, t.dimUnit || unit(), (v) => commit(() => { t.dimUnit = v; touch(t); }));

  grid.append(field('Code', textInput(t, 'code')));
  if (t.kind === 'length') {
    grid.append(field('Output', selectInput([['base', 'Length'], ['area', 'Area (× width)'], ['volume', 'Volume (× width × depth)']], t.out || 'base', (v) => commit(() => { t.out = v; }))));
    if (t.out === 'area' || t.out === 'volume') grid.append(field('Width', numInput(t, 'width', 'any', 0)));
    if (t.out === 'volume') grid.append(field('Depth', numInput(t, 'depth', 'any', 0)));
    if (t.out === 'area' || t.out === 'volume') grid.append(field('Size unit', dimUnit()));
  } else if (t.kind === 'area') {
    grid.append(field('Output', selectInput([['base', 'Area'], ['perimeter', 'Perimeter'], ['volume', 'Volume (× depth)']], t.out || 'base', (v) => commit(() => { t.out = v; }))));
    if (t.out === 'volume') { grid.append(field('Depth', numInput(t, 'depth', 'any', 0))); grid.append(field('Size unit', dimUnit())); }
    if (t.shape === 'sector') grid.append(checkInput('Other side of the sector', !!t.flip, (v) => commit(() => { t.flip = v; touch(t); })));
    grid.append(checkInput('Deduct (opening or void)', !!t.deduct, (v) => commit(() => { t.deduct = v; })));
  } else if (t.kind === 'count' && t.shape === 'spaced') {
    grid.append(field('Spacing', numInput(t, 'spacing', 'any', 0)));
    grid.append(field('Spacing unit', dimUnit()));
    grid.append(checkInput('Post at the far end', t.endPost !== false, (v) => commit(() => { t.endPost = v; })));
  }
  if (t.kind !== 'angle') {
    grid.append(field('Repeat ×', numInput(t, 'mult', 'any', 0)));
    grid.append(field('Waste %', numInput(t, 'waste', 'any')));
    grid.append(field('Rate', numInput(t, 'rate', 'any')));
  }
  grid.append(field('Notes', textInput(t, 'notes')));
  d.append(grid);
  return d;
}

function renderTakeoff() {
  const ul = $('takeoff');
  ul.replaceChildren();
  $('takeoff-empty').hidden = S.takeoff.length > 0;
  const totals = new Map();
  for (const t of S.takeoff) {
    const li = document.createElement('li');
    li.className = `item${t.id === S.sel ? ' selected' : ''}`;

    const head = document.createElement('div');
    head.className = 'item-head';
    const sw = document.createElement('span');
    sw.className = 'swatch';
    sw.style.background = t.color;
    const name = document.createElement('input');
    name.type = 'text';
    name.value = t.name;
    name.setAttribute('aria-label', 'Item name');
    name.addEventListener('change', () => commit(() => { t.name = name.value.trim() || t.name; }));
    head.append(sw, name);

    const q = qtyOf(t);
    const info = document.createElement('div');
    info.className = 'item-info';
    const shapeName = t.shape === 'trace' ? 'picked' : ((SHAPES[t.kind] && SHAPES[t.kind][t.shape] && SHAPES[t.kind][t.shape].label) || '').split(' (')[0].toLowerCase();
    const where = `${shapeName ? ` · ${shapeName}` : ''}${S.pdf ? ` · page ${t.page}` : ''}`;
    if (q && !q.missing) {
      info.textContent = `${typeLabel(t)} ${fmt(q.value)} ${q.unit}${where}`;
      const key = `${typeLabel(t)}|${q.unit}`;
      totals.set(key, (totals.get(key) || 0) + q.value);
    } else if (q && q.missing) info.textContent = `${typeLabel(t)}: enter a ${q.missing} under Details${where}`;
    else info.textContent = `${typeLabel(t)}: no scale here yet${where}`;

    const actions = document.createElement('div');
    actions.className = 'chips';
    actions.append(
      mkBtn('Show', () => selectItem(t.id)),
      mkBtn(t.link ? 'Send to link' : 'Send', () => sendItem(t)),
      mkBtn('Delete', () => commit(() => { S.takeoff = S.takeoff.filter((x) => x.id !== t.id); if (S.sel === t.id) S.sel = null; }), `Delete ${t.name}`),
    );
    li.append(head, info, actions);

    const link = document.createElement('div');
    link.className = 'chips';
    if (t.link) {
      const a = document.createElement('span');
      a.className = 'note inline';
      a.textContent = `Linked to ${t.link.address}`;
      link.append(a, mkBtn('Unlink', () => commit(() => { delete t.link; })));
    } else {
      link.append(mkBtn('Link to selected cell', () => linkItem(t)));
    }
    li.append(link, itemDetails(t));
    ul.append(li);
  }
  const text = [...totals].map(([k, v]) => { const [type, u] = k.split('|'); return `${type} ${fmt(v)} ${u}`; }).join('  ·  ');
  $('takeoff-total').textContent = text ? `Totals: ${text}` : '';
  $('undo').disabled = !S.undo.length;
  $('redo').disabled = !S.redo.length;
  $('update-links').hidden = !S.takeoff.some((t) => t.link);
}

async function selectItem(id) {
  const t = S.takeoff.find((x) => x.id === id);
  if (!t) return;
  S.sel = id;
  if (S.pdf && t.page !== S.page) await showPage(t.page);
  const g = t.kind === 'count' || t.kind === 'angle' ? { poly: t.pts } : itemGeom(t);
  const xs = g.poly.map((p) => p[0]);
  const ys = g.poly.map((p) => p[1]);
  const pad = Math.max(Math.max(...xs) - Math.min(...xs), Math.max(...ys) - Math.min(...ys), 1) * 0.2;
  fit({ x0: Math.min(...xs) - pad, x1: Math.max(...xs) + pad, y0: Math.min(...ys) - pad, y1: Math.max(...ys) + pad });
  renderTakeoff();
}

/* ---------- saved sessions (stored with the workbook, browser storage as a fallback) ---------- */
function storeGet(key) {
  try { if (S.inExcel) { const v = Office.context.document.settings.get(key); if (v) return v; } } catch { /* ignore */ }
  try { return localStorage.getItem(key); } catch { return null; }
}
function storeSet(key, val) {
  try {
    if (S.inExcel) { Office.context.document.settings.set(key, val); Office.context.document.settings.saveAsync(() => {}); return; }
  } catch { /* fall through */ }
  try { localStorage.setItem(key, val); } catch { /* ignore */ }
}

let saveTimer = 0;
function saveSoon() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(saveSession, 600);
}

function saveSession() {
  if (!S.fileKey) return;
  storeSet(`takeoff:${S.fileKey}`, JSON.stringify({
    v: 1,
    unit: unit(),
    rules: S.rules.map((r) => ({ ...r, pages: r.pages ? [...r.pages] : undefined })),
    calibration: S.calibration,
    takeoff: S.takeoff,
  }));
}

// Returns a status message when something was restored.
function restoreSession() {
  const raw = storeGet(`takeoff:${S.fileKey}`);
  if (!raw) return '';
  try {
    const d = JSON.parse(raw);
    if (d.v !== 1) return '';
    if (d.rules && d.rules.length) S.rules = d.rules.map((r) => ({ ...r, pages: r.pages ? new Set(r.pages) : undefined }));
    S.calibration = d.calibration || null;
    if (S.calibration) $('cal-result').textContent = `1 drawing unit = ${+S.calibration.factor.toPrecision(6)} ${S.calibration.unit}`;
    S.takeoff = d.takeoff || [];
    if (d.unit && UNIT_M[d.unit]) $('unit').value = d.unit;
    return `Restored ${S.rules.length} scale${S.rules.length === 1 ? '' : 's'} and ${S.takeoff.length} takeoff item${S.takeoff.length === 1 ? '' : 's'} for this file.`;
  } catch { return ''; }
}
