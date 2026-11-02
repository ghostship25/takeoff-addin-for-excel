/* global scaleAtPoint, scheduleSharp, drawOverlays, toolClick, toolDouble, toolFinish, toolCancel, toolBackspace,
   finishRegion, symbolDrag, symbolClick, showMethod, undo, redo, deleteSelected, hintFor, needsFinishBar, isDragShape,
   isTwoPoint, handleAt, beginEdit, moveHandle, endEdit, selectAt, addFreePoint, finishFreehand, updatePickHover */
const $ = (id) => document.getElementById(id);
const cv = $('plan');
const ctx = cv.getContext('2d');

const S = {
  // geometry for the current page / drawing
  items: [], inserts: [], layerPaths: new Map(), grid: null, cell: 1, gridOrigin: [0, 0],
  hidden: new Set(), focus: null,
  // viewer
  view: { x: 0, y: 0, k: 1 }, hover: null, tool: 'pan', pts: [], drag: null, box: null,
  // document
  rules: [], page: 1, pdf: null, pageObj: null, bg: null, hi: null, fileName: '', fileKey: '',
  pendingScale: null, calibration: null, calRaw: 0, lastDist: null,
  // takeoff
  takeoff: [], sel: null, activeCount: null, undo: [], redo: [], sym: null,
  inExcel: false,
};
let W = 0;
let H = 0;
let drawQueued = false;

const dpr = () => window.devicePixelRatio || 1;
const say = (msg) => { $('status').textContent = msg; };
const fmt = (n) => n.toLocaleString(undefined, { maximumFractionDigits: S.prefs ? S.prefs.decimals : 2 });
const unit = () => $('unit').value;
const toScreen = (x, y) => [(x - S.view.x) * S.view.k, H - (y - S.view.y) * S.view.k];
const toWorld = (sx, sy) => [sx / S.view.k + S.view.x, (H - sy) / S.view.k + S.view.y];

if (typeof Office !== 'undefined') {
  Office.onReady((info) => { S.inExcel = info.host === Office.HostType.Excel; });
}

/* ---------- geometry ---------- */
function polyLen(pts, closed) {
  let d = 0;
  for (let i = 1; i < pts.length; i++) d += Math.hypot(pts[i][0] - pts[i - 1][0], pts[i][1] - pts[i - 1][1]);
  if (closed && pts.length > 2) d += Math.hypot(pts[0][0] - pts.at(-1)[0], pts[0][1] - pts.at(-1)[1]);
  return d;
}

function polyArea(pts) {
  if (pts.length < 3) return 0;
  let a = 0;
  for (let i = 0; i < pts.length; i++) {
    const [x1, y1] = pts[i];
    const [x2, y2] = pts[(i + 1) % pts.length];
    a += x1 * y2 - x2 * y1;
  }
  return Math.abs(a) / 2;
}

function bboxCenter(pts) {
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  for (const [x, y] of pts) { x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y); }
  return [(x0 + x1) / 2, (y0 + y1) / 2];
}

// Every geometry item carries its raw length / area in drawing units, computed once.
const pathItem = (layer, pts, closed, len, area) => (pts.length < 2 ? null : {
  type: 'path', layer, pts, closed,
  len: len ?? polyLen(pts, closed),
  area: area ?? (closed ? polyArea(pts) : 0),
});
const circleItem = (layer, c, r) => ({ type: 'circle', layer, c, r, len: 2 * Math.PI * r, area: Math.PI * r * r });
const arcItem = (layer, c, r, a0, a1) => {
  const sweep = (((a1 - a0) % (2 * Math.PI)) + 2 * Math.PI) % (2 * Math.PI);
  return { type: 'arc', layer, c, r, a0, a1, len: r * sweep, area: 0 };
};
const insertItem = (layer, name, p) => ({ type: 'insert', layer, name, p, len: 0, area: 0 });

// The point that decides which scale an object uses.
function repPoint(it) {
  if (!it.rp) it.rp = it.type === 'path' ? bboxCenter(it.pts) : it.type === 'insert' ? it.p : it.c;
  return it.rp;
}

function layerStats() {
  const m = new Map();
  for (const it of S.items) {
    let s = m.get(it.layer);
    if (!s) { s = { len: 0, area: 0, count: 0, unscaled: 0, lenT: new Map(), areaT: new Map() }; m.set(it.layer, s); }
    if (it.type === 'insert') { s.count += 1; continue; }
    const f = scaleAtPoint(repPoint(it));
    if (f == null) { s.unscaled += 1; continue; }
    s.len += it.len * f;
    s.area += it.area * f * f;
    s.lenT.set(f, (s.lenT.get(f) || 0) + it.len);
    s.areaT.set(f, (s.areaT.get(f) || 0) + it.area);
  }
  return m;
}

/* ---------- loading geometry: Path2D per layer + a grid index for snapping ---------- */
function bounds() {
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  const add = (x, y) => { if (x < x0) x0 = x; if (y < y0) y0 = y; if (x > x1) x1 = x; if (y > y1) y1 = y; };
  for (const it of S.items) {
    if (it.type === 'path') for (const p of it.pts) add(p[0], p[1]);
    else if (it.type === 'insert') add(it.p[0], it.p[1]);
    else { add(it.c[0] - it.r, it.c[1] - it.r); add(it.c[0] + it.r, it.c[1] + it.r); }
  }
  if (S.bg) { add(0, 0); add(S.bg.w, S.bg.h); }
  return x0 === Infinity ? null : { x0, y0, x1, y1 };
}

function setItems(items) {
  S.items = items;
  S.inserts = items.filter((i) => i.type === 'insert');
  S.hidden.clear();
  S.focus = null;

  const paths = new Map();
  for (const it of items) {
    if (it.type === 'insert') continue;
    let p = paths.get(it.layer);
    if (!p) { p = new Path2D(); paths.set(it.layer, p); }
    if (it.type === 'path') {
      it.pts.forEach(([x, y], i) => (i ? p.lineTo(x, y) : p.moveTo(x, y)));
      if (it.closed) p.closePath();
    } else if (it.type === 'circle') {
      p.moveTo(it.c[0] + it.r, it.c[1]);
      p.arc(it.c[0], it.c[1], it.r, 0, 2 * Math.PI);
    } else {
      p.moveTo(it.c[0] + Math.cos(it.a0) * it.r, it.c[1] + Math.sin(it.a0) * it.r);
      p.arc(it.c[0], it.c[1], it.r, it.a0, it.a1, false);
    }
  }
  S.layerPaths = paths;
  buildGrid();
}

function buildGrid() {
  S.grid = null;
  const b = bounds();
  if (!b) return;
  let n = 0;
  for (const it of S.items) n += it.type === 'path' ? it.pts.length : 1;
  if (n > 2e6) return; // too heavy: snapping to drawing vertices is switched off
  const cell = Math.max(b.x1 - b.x0, b.y1 - b.y0, 1e-9) / 128;
  const g = new Map();
  const add = (x, y, layer) => {
    const key = Math.floor((x - b.x0) / cell) * 100000 + Math.floor((y - b.y0) / cell);
    const arr = g.get(key);
    if (arr) arr.push(x, y, layer); else g.set(key, [x, y, layer]);
  };
  for (const it of S.items) {
    if (it.type === 'path') for (const p of it.pts) add(p[0], p[1], it.layer);
    else add((it.p || it.c)[0], (it.p || it.c)[1], it.layer);
  }
  S.grid = g;
  S.cell = cell;
  S.gridOrigin = [b.x0, b.y0];
}

function resetDocument(file) {
  S.pdf = null; S.pageObj = null; S.bg = null; S.hi = null; S.page = 1;
  S.rules = []; S.calibration = null; S.box = null; S.pts = []; S.hover = null;
  S.takeoff = []; S.sel = null; S.activeCount = null; S.undo = []; S.redo = []; S.sym = null;
  S.fileName = file.name;
  S.fileKey = `${file.name}|${file.size}`;
  $('cal-result').textContent = 'Not calibrated yet.';
  for (const id of ['cal-panel', 'measure-out', 'symbol-panel', 'pdf-controls']) $(id).hidden = true;
}

/* ---------- snapping ---------- */
function snapPoint(sx, sy, ortho) {
  const [wx, wy] = toWorld(sx, sy);
  const base = S.pts.length ? S.pts.at(-1) : null;
  if (ortho && base) return Math.abs(wx - base[0]) > Math.abs(wy - base[1]) ? [wx, base[1]] : [base[0], wy];
  if (!$('snap').checked) return [wx, wy];

  const tol = 12 / S.view.k;
  let best = null;
  let bd = tol;
  const test = (x, y) => { const d = Math.hypot(x - wx, y - wy); if (d < bd) { bd = d; best = [x, y]; } };

  if (S.grid) {
    const [ox, oy] = S.gridOrigin;
    const r = Math.ceil(tol / S.cell);
    const cx = Math.floor((wx - ox) / S.cell);
    const cy = Math.floor((wy - oy) / S.cell);
    for (let i = cx - r; i <= cx + r; i++) {
      for (let j = cy - r; j <= cy + r; j++) {
        const arr = S.grid.get(i * 100000 + j);
        if (!arr) continue;
        for (let n = 0; n < arr.length; n += 3) if (!S.hidden.has(arr[n + 2])) test(arr[n], arr[n + 1]);
      }
    }
  }
  for (const t of S.takeoff) if (t.page === S.page) for (const p of t.pts) test(p[0], p[1]);
  return best || [wx, wy];
}

/* ---------- drawing ---------- */
function requestDraw() {
  if (drawQueued) return;
  drawQueued = true;
  requestAnimationFrame(draw);
}

function viewChanged() {
  S.hi = null;
  if (typeof scheduleSharp === 'function') scheduleSharp();
  requestDraw();
}

function resize() {
  const r = cv.getBoundingClientRect();
  const d = dpr();
  W = r.width;
  H = r.height;
  cv.width = Math.round(W * d);
  cv.height = Math.round(H * d);
  viewChanged();
}
new ResizeObserver(resize).observe(cv);

function fit(box) {
  const b = box || bounds();
  if (!b || !W) return;
  const k = Math.min(W / Math.max(b.x1 - b.x0, 1e-9), H / Math.max(b.y1 - b.y0, 1e-9)) * 0.9;
  S.view.k = k;
  S.view.x = (b.x0 + b.x1) / 2 - W / 2 / k;
  S.view.y = (b.y0 + b.y1) / 2 - H / 2 / k;
  viewChanged();
}

function draw() {
  drawQueued = false;
  const d = dpr();
  const { x: vx, y: vy, k } = S.view;
  ctx.setTransform(d, 0, 0, d, 0, 0);
  ctx.clearRect(0, 0, W, H);
  const css = getComputedStyle(document.documentElement);
  const ink = css.getPropertyValue('--drawing').trim();
  const flag = css.getPropertyValue('--flag').trim();

  const bgOn = S.bg && $('show-bg').checked;
  if (bgOn) {
    const [bx, by] = toScreen(0, S.bg.h);
    ctx.drawImage(S.bg.canvas, bx, by, S.bg.w * k, S.bg.h * k);
    const hi = S.hi;
    if (hi && hi.page === S.page && hi.x === vx && hi.y === vy && hi.k === k) ctx.drawImage(hi.canvas, 0, 0, W, H);
  }

  // Vectors: one Path2D per layer, drawn in world space (the browser does the transform).
  ctx.save();
  ctx.setTransform(d * k, 0, 0, -d * k, -d * k * vx, d * (H + vy * k));
  for (const focused of [false, true]) {
    if (bgOn && !focused) continue; // the page image already shows them
    ctx.strokeStyle = focused ? flag : ink;
    ctx.lineWidth = (focused ? 2 : 1) / k;
    for (const [layer, p] of S.layerPaths) {
      if (S.hidden.has(layer) || (layer === S.focus) !== focused) continue;
      ctx.stroke(p);
    }
  }
  ctx.restore();

  if (S.inserts.length && !(bgOn && !S.focus)) {
    ctx.strokeStyle = ink;
    ctx.lineWidth = 1;
    ctx.beginPath();
    for (const it of S.inserts) {
      if (S.hidden.has(it.layer)) continue;
      const [x, y] = toScreen(it.p[0], it.p[1]);
      ctx.moveTo(x - 4, y); ctx.lineTo(x + 4, y);
      ctx.moveTo(x, y - 4); ctx.lineTo(x, y + 4);
    }
    ctx.stroke();
  }

  // dashed outlines for regions that carry their own scale
  ctx.strokeStyle = flag;
  ctx.lineWidth = 1.5;
  ctx.setLineDash([6, 4]);
  const strokeBox = (b) => { const [x, y] = toScreen(b.x0, b.y1); ctx.strokeRect(x, y, (b.x1 - b.x0) * k, (b.y1 - b.y0) * k); };
  for (const r of S.rules) if (r.scope === 'region' && r.page === S.page) strokeBox(r.box);
  if (S.box) strokeBox(S.box);
  ctx.setLineDash([]);

  if (typeof drawOverlays === 'function') drawOverlays(flag, ink);
}

/* ---------- tools ---------- */
const HOVER_TOOLS = new Set(['distance', 'length', 'area', 'count', 'calibrate', 'angle']);

function setTool(t) {
  S.tool = t;
  S.pts = [];
  S.box = null;
  S.hover = null;
  S.pickHover = null;
  S.activeCount = null;
  if (t !== 'region') S.pendingScale = null;
  document.querySelectorAll('[data-tool]').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.tool === t)));
  cv.style.cursor = t === 'pan' ? 'grab' : 'crosshair';
  $('measure-out').hidden = true;
  $('cal-panel').hidden = true;
  $('finish-bar').hidden = !needsFinishBar();
  $('symbol-panel').hidden = !(S.sym || t === 'symbol');
  if (t === 'calibrate') showMethod('calibrate');
  const hint = hintFor(t);
  if (hint) say(hint);
  requestDraw();
}

/* ---------- pointer, wheel, keyboard ---------- */
// Right or middle button always pans. Drag modes: pan, edit (a handle), free (freehand), box (region/symbol), two (2-point shapes).
cv.addEventListener('contextmenu', (e) => e.preventDefault());

cv.addEventListener('pointerdown', (e) => {
  cv.setPointerCapture(e.pointerId);
  S.drag = { x: e.clientX, y: e.clientY, vx: S.view.x, vy: S.view.y, moved: false, sx: e.offsetX, sy: e.offsetY, w0: toWorld(e.offsetX, e.offsetY), mode: 'pan', button: e.button };
  if (e.button !== 0) return;
  if (S.tool === 'pan') {
    const h = handleAt(e.offsetX, e.offsetY);
    if (h) { S.drag.mode = 'edit'; beginEdit(h); }
  } else if (isDragShape()) {
    S.drag.mode = 'free';
    S.pts = [S.drag.w0];
  } else if (S.tool === 'region' || S.tool === 'symbol') {
    S.drag.mode = 'box';
  } else if (isTwoPoint() && S.pts.length === 0) {
    S.drag.mode = 'two';
  }
});

cv.addEventListener('pointermove', (e) => {
  if (!S.drag) {
    if (HOVER_TOOLS.has(S.tool)) { S.hover = snapPoint(e.offsetX, e.offsetY, e.shiftKey); requestDraw(); }
    else if (S.tool === 'pick') { updatePickHover(toWorld(e.offsetX, e.offsetY)); requestDraw(); }
    return;
  }
  const dx = e.clientX - S.drag.x;
  const dy = e.clientY - S.drag.y;
  if (Math.hypot(dx, dy) > 3) S.drag.moved = true;
  if (!S.drag.moved) return;
  switch (S.drag.mode) {
    case 'edit': moveHandle(snapPoint(e.offsetX, e.offsetY, false)); break;
    case 'free': addFreePoint(toWorld(e.offsetX, e.offsetY)); break;
    case 'box': {
      const w1 = toWorld(e.offsetX, e.offsetY);
      S.box = {
        x0: Math.min(S.drag.w0[0], w1[0]), x1: Math.max(S.drag.w0[0], w1[0]),
        y0: Math.min(S.drag.w0[1], w1[1]), y1: Math.max(S.drag.w0[1], w1[1]),
      };
      requestDraw();
      break;
    }
    case 'two':
      if (!S.pts.length) S.pts = [snapPoint(S.drag.sx, S.drag.sy, false)];
      S.hover = snapPoint(e.offsetX, e.offsetY, e.shiftKey);
      requestDraw();
      break;
    default:
      S.hover = null;
      S.view.x = S.drag.vx - dx / S.view.k;
      S.view.y = S.drag.vy + dy / S.view.k;
      requestDraw();
  }
});

cv.addEventListener('pointerleave', () => { if (!S.drag) { S.hover = null; S.pickHover = null; requestDraw(); } });

cv.addEventListener('pointerup', (e) => {
  const d = S.drag;
  S.drag = null;
  if (!d) return;
  switch (d.mode) {
    case 'edit': endEdit(); return;
    case 'free': finishFreehand(); return;
    case 'box':
      if (S.tool === 'region') { if (d.moved && S.box) finishRegion(S.box); }
      else if (d.moved && S.box) symbolDrag(S.box);
      else if (!d.moved) symbolClick(e.offsetX, e.offsetY);
      return;
    case 'two':
      if (d.moved) { toolClick(snapPoint(e.offsetX, e.offsetY, e.shiftKey)); return; }
      break;
    default: break;
  }
  if (d.moved) { viewChanged(); return; }
  if (d.button !== 0) return;
  if (S.tool === 'pan') { selectAt(e.offsetX, e.offsetY); return; }
  toolClick(snapPoint(e.offsetX, e.offsetY, e.shiftKey));
});

cv.addEventListener('dblclick', () => toolDouble());

cv.addEventListener('wheel', (e) => {
  e.preventDefault();
  const f = e.deltaY < 0 ? 1.15 : 1 / 1.15;
  const [wx, wy] = toWorld(e.offsetX, e.offsetY);
  S.view.k *= f;
  S.view.x = wx - e.offsetX / S.view.k;
  S.view.y = wy - (H - e.offsetY) / S.view.k;
  viewChanged();
}, { passive: false });

const KEY_TOOLS = { p: 'pan', d: 'distance', l: 'length', a: 'area', c: 'count', t: 'pick', g: 'angle', s: 'symbol', k: 'calibrate' };
document.addEventListener('keydown', (e) => {
  if (/INPUT|SELECT|TEXTAREA/.test(e.target.tagName)) return;
  if (e.ctrlKey || e.metaKey) {
    const key = e.key.toLowerCase();
    if (key === 'z') { e.preventDefault(); if (e.shiftKey) redo(); else undo(); }
    else if (key === 'y') { e.preventDefault(); redo(); }
    return;
  }
  const key = e.key.toLowerCase();
  if (KEY_TOOLS[key]) setTool(KEY_TOOLS[key]);
  else if (key === 'f') fit();
  else if (e.key === 'Enter') toolFinish();
  else if (e.key === 'Escape') toolCancel();
  else if (e.key === 'Backspace') { e.preventDefault(); toolBackspace(); }
  else if (e.key === 'Delete') deleteSelected();
});
