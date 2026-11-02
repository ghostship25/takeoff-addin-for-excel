/* global S, $, say, draw, renderLayers, renderTakeoff, saveSoon, setTool, unit */
// Scale rules. Each rule says "1 drawing unit = factor <unit>" and where it applies:
//   scope 'region' (a box on one page)  beats  'pages' (one or more pages)  beats  'all'.
// Later rules of the same kind win over earlier ones.

const UNIT_M = { ft: 0.3048, m: 1, in: 0.0254, mm: 0.001, cm: 0.01 };

// perInch = real distance that one inch on the printed sheet stands for.
const ARCH = [['1/16"', 16], ['3/32"', 32 / 3], ['1/8"', 8], ['3/16"', 16 / 3], ['1/4"', 4], ['3/8"', 8 / 3],
  ['1/2"', 2], ['3/4"', 4 / 3], ['1"', 1], ['1-1/2"', 2 / 3], ['3"', 1 / 3]];
const PRESETS = [
  ...ARCH.map(([l, p]) => ({ group: 'Architectural', label: `${l} = 1'-0"`, perInch: p, unit: 'ft' })),
  ...[10, 20, 30, 40, 50, 60, 100].map((n) => ({ group: 'Engineering', label: `1" = ${n}'`, perInch: n, unit: 'ft' })),
  ...[10, 20, 25, 50, 75, 100, 125, 150, 200, 250, 500, 1000, 1250, 2000, 2500, 5000].map((n) => ({ group: 'Metric ratio', label: `1:${n}`, perInch: n, unit: 'in' })),
];

/* ---------- resolving which scale applies ---------- */
const inBox = (b, p) => p[0] >= b.x0 && p[0] <= b.x1 && p[1] >= b.y0 && p[1] <= b.y1;

function ruleAt(pt, page = S.page) {
  for (let i = S.rules.length - 1; i >= 0; i--) {
    const r = S.rules[i];
    if (r.scope === 'region' && r.page === page && inBox(r.box, pt)) return r;
  }
  for (let i = S.rules.length - 1; i >= 0; i--) {
    const r = S.rules[i];
    if (r.scope === 'pages' && r.pages.has(page)) return r;
  }
  for (let i = S.rules.length - 1; i >= 0; i--) {
    if (S.rules[i].scope === 'all') return S.rules[i];
  }
  return null;
}

// Real distance (in the "show quantities in" unit) per drawing unit at a point, or null.
function scaleAtPoint(pt, page = S.page) {
  const r = ruleAt(pt, page);
  return r ? (r.factor * UNIT_M[r.unit]) / UNIT_M[unit()] : null;
}

/* ---------- page lists ---------- */
function parsePages(text, max) {
  const set = new Set();
  for (const part of text.split(',')) {
    const m = part.trim().match(/^(\d+)(?:\s*-\s*(\d+))?$/);
    if (!m) continue;
    const a = +m[1];
    const b = m[2] ? +m[2] : a;
    for (let p = Math.max(1, Math.min(a, b)); p <= Math.min(max, Math.max(a, b)); p++) set.add(p);
  }
  return set;
}

function compactPages(set) {
  const list = [...set].sort((a, b) => a - b);
  const out = [];
  for (let i = 0; i < list.length;) {
    let j = i;
    while (j + 1 < list.length && list[j + 1] === list[j] + 1) j++;
    out.push(j > i ? `${list[i]}-${list[j]}` : `${list[i]}`);
    i = j + 1;
  }
  return out.join(', ');
}

function scopeText(r) {
  if (r.scope === 'all') return S.pdf ? 'All pages' : 'Whole drawing';
  if (r.scope === 'region') return `Page ${r.page}, region`;
  if (!S.pdf) return 'Whole drawing';
  return r.pages.size === 1 ? `Page ${[...r.pages][0]}` : `Pages ${compactPages(r.pages)}`;
}

/* ---------- UI ---------- */
function refreshScales() {
  renderScaleList();
  renderLayers();
  renderTakeoff();
  draw();
  saveSoon();
}

function addRule(rule) {
  S.rules.push(rule);
  refreshScales();
}

function renderScaleList() {
  const ul = $('rules');
  ul.replaceChildren();
  if (!S.rules.length) {
    const li = document.createElement('li');
    li.className = 'none';
    li.textContent = 'No scale set yet.';
    ul.append(li);
    return;
  }
  S.rules.forEach((r, i) => {
    const li = document.createElement('li');
    li.className = 'rule';
    const text = document.createElement('span');
    text.textContent = `${scopeText(r)}: ${r.label}`;
    const rm = document.createElement('button');
    rm.className = 'btn';
    rm.textContent = 'Remove';
    rm.setAttribute('aria-label', `Remove scale for ${scopeText(r)}`);
    rm.addEventListener('click', () => { S.rules.splice(i, 1); refreshScales(); });
    li.append(text, rm);
    ul.append(li);
  });
}

function showMethod(m) {
  $('method').value = m;
  for (const id of ['preset', 'calibrate', 'factor']) $(`m-${id}`).hidden = id !== m;
}

// Turns the current method's inputs into { factor, unit, label }, or says what is missing.
function chosenScale() {
  const m = $('method').value;
  if (m === 'preset') {
    if (!S.pdf) { say('Paper scales apply to PDF sheets. For a DXF, type a factor instead.'); return null; }
    const v = $('preset').value;
    if (v === 'custom') {
      const n = parseFloat($('ratio').value);
      if (!(n > 0)) { say('Enter a ratio greater than 0.'); return null; }
      return { factor: n / 72, unit: 'in', label: `1:${n}` };
    }
    const p = PRESETS[+v];
    return { factor: p.perInch / 72, unit: p.unit, label: p.label }; // PDF units are 1/72 inch
  }
  if (m === 'calibrate') {
    if (!S.calibration) { say('Calibrate first: pick two points and enter the real distance.'); return null; }
    return S.calibration;
  }
  const f = parseFloat($('factor').value);
  if (!(f > 0)) { say('Enter a factor greater than 0.'); return null; }
  const u = $('factor-unit').value;
  return { factor: f, unit: u, label: `1 drawing unit = ${f} ${u}` };
}

function finishRegion(box) {
  const base = S.pendingScale;
  if (!base) return;
  addRule({ ...base, scope: 'region', page: S.page, box });
  setTool('pan');
  say(`Scale applied to the region: ${base.label}.`);
}

(function initScaleUi() {
  const sel = $('preset');
  const groups = new Map();
  PRESETS.forEach((p, i) => {
    if (!groups.has(p.group)) {
      const g = document.createElement('optgroup');
      g.label = p.group;
      groups.set(p.group, g);
      sel.append(g);
    }
    const o = document.createElement('option');
    o.value = String(i);
    o.textContent = p.label;
    groups.get(p.group).append(o);
  });
  const custom = document.createElement('option');
  custom.value = 'custom';
  custom.textContent = 'Custom ratio';
  sel.append(custom);
  sel.value = String(PRESETS.findIndex((p) => p.label.startsWith('1/4"')));

  sel.addEventListener('change', () => { $('ratio-wrap').hidden = sel.value !== 'custom'; });
  $('method').addEventListener('change', () => showMethod($('method').value));
  $('scope').addEventListener('change', () => { $('pages-wrap').hidden = $('scope').value !== 'pages'; });
  $('unit').addEventListener('change', refreshScales);
  $('cal-start').addEventListener('click', () => setTool('calibrate'));

  $('cal-set').addEventListener('click', () => {
    const known = parseFloat($('cal-dist').value);
    if (!(known > 0) || !(S.calRaw > 0)) { say('Enter the real distance, and pick two different points.'); return; }
    const u = $('cal-unit').value;
    S.calibration = { factor: known / S.calRaw, unit: u, label: `Calibrated (${known} ${u} measured)` };
    $('cal-result').textContent = `1 drawing unit = ${+S.calibration.factor.toPrecision(6)} ${u}`;
    setTool('pan');
    say('Distance recorded. Choose where to apply it, then click Apply scale.');
  });

  $('apply-scale').addEventListener('click', () => {
    const base = chosenScale();
    if (!base) return;
    const scope = $('scope').value;
    if (scope === 'region') {
      setTool('region');
      S.pendingScale = base;
      say('Drag a box around the area this scale applies to.');
      return;
    }
    if (scope === 'all') {
      addRule({ ...base, scope: 'all' });
      say(`Scale applied to ${S.pdf ? 'all pages' : 'the drawing'}: ${base.label}.`);
      return;
    }
    const pages = scope === 'page' ? new Set([S.page]) : parsePages($('pages').value, S.pdf ? S.pdf.numPages : 1);
    if (!pages.size) { say('Enter page numbers like 1-3, 5.'); return; }
    addRule({ ...base, scope: 'pages', pages });
    say(`Scale applied to ${scopeText({ scope: 'pages', pages }).toLowerCase()}: ${base.label}.`);
  });

  showMethod('preset');
  renderScaleList();
})();
