/* global pdfjsLib, S, $, W, H, say, dpr, pathItem, setItems, resetDocument, requestDraw, fit, resize, setTool,
   renderLayers, renderScaleList, renderTakeoff, restoreSession, showMethod */
// PDF support: the page is drawn as an image, and the vector lines are pulled out of the PDF's
// drawing commands (grouped by color) so they can be totaled. Written for pdf.js 3.11.x
// (the operator-list format changed in v4).

if (typeof pdfjsLib !== 'undefined') {
  pdfjsLib.GlobalWorkerOptions.workerSrc = window.PDF_LOCAL
    ? 'vendor/pdf.worker.min.js'
    : 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

const MAX_ITEMS = 200000;

// Apply matrix a first, then b (PDF row-vector convention).
const mul = (a, b) => [
  a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3],
  a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3],
  a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5],
];

function hexColor(a) {
  if (typeof a === 'string') return a;
  if (Array.isArray(a) && typeof a[0] === 'string') return a[0];
  if (Array.isArray(a) && a.length >= 3 && a.slice(0, 3).every(Number.isFinite)) {
    return `#${a.slice(0, 3).map((v) => Math.round(v).toString(16).padStart(2, '0')).join('')}`;
  }
  return '#000000';
}

function bezier(p0, c1, c2, p3, n = 8) {
  const pts = [];
  for (let i = 1; i <= n; i++) {
    const t = i / n;
    const u = 1 - t;
    pts.push([
      u * u * u * p0[0] + 3 * u * u * t * c1[0] + 3 * u * t * t * c2[0] + t * t * t * p3[0],
      u * u * u * p0[1] + 3 * u * u * t * c1[1] + 3 * u * t * t * c2[1] + t * t * t * p3[1],
    ]);
  }
  return pts;
}

function extractVector(opList, vp) {
  const OPS = pdfjsLib.OPS;
  const items = [];
  const stack = [];
  let gs = { ctm: [1, 0, 0, 1, 0, 0], stroke: '#000000', fill: '#000000' };
  let subs = [];
  let cur = null;
  let user = [0, 0]; // current point in PDF user space

  const toWorld = ([x, y]) => {
    const px = gs.ctm[0] * x + gs.ctm[2] * y + gs.ctm[4];
    const py = gs.ctm[1] * x + gs.ctm[3] * y + gs.ctm[5];
    const [vx, vy] = vp.convertToViewportPoint(px, py);
    return [vx, vp.height - vy];
  };
  const startSub = (p) => { cur = { pts: [toWorld(p)], closed: false }; subs.push(cur); };
  const lineTo = (p) => { if (!cur) startSub(user); cur.pts.push(toWorld(p)); user = p; };

  const build = (ops, c) => {
    let i = 0;
    for (const op of ops) {
      if (op === OPS.moveTo) { user = [c[i], c[i + 1]]; startSub(user); i += 2; }
      else if (op === OPS.lineTo) { lineTo([c[i], c[i + 1]]); i += 2; }
      else if (op === OPS.curveTo) {
        if (!cur) startSub(user);
        const end = [c[i + 4], c[i + 5]];
        bezier(user, [c[i], c[i + 1]], [c[i + 2], c[i + 3]], end).forEach((p) => cur.pts.push(toWorld(p)));
        user = end; i += 6;
      } else if (op === OPS.curveTo2) {
        if (!cur) startSub(user);
        const end = [c[i + 2], c[i + 3]];
        bezier(user, user, [c[i], c[i + 1]], end).forEach((p) => cur.pts.push(toWorld(p)));
        user = end; i += 4;
      } else if (op === OPS.curveTo3) {
        if (!cur) startSub(user);
        const end = [c[i + 2], c[i + 3]];
        bezier(user, [c[i], c[i + 1]], end, end).forEach((p) => cur.pts.push(toWorld(p)));
        user = end; i += 4;
      } else if (op === OPS.closePath) {
        if (cur) cur.closed = true;
        cur = null;
      } else if (op === OPS.rectangle) {
        const [x, y, w, h] = [c[i], c[i + 1], c[i + 2], c[i + 3]];
        subs.push({ pts: [[x, y], [x + w, y], [x + w, y + h], [x, y + h]].map(toWorld), closed: true });
        cur = null; user = [x, y]; i += 4;
      }
    }
  };

  const flush = (mode, closeLast = false) => {
    if (closeLast && subs.length) subs.at(-1).closed = true;
    const color = mode === 'fill' ? gs.fill : gs.stroke;
    for (const s of subs) {
      const it = pathItem(`Color ${color}`, s.pts, s.closed || mode !== 'stroke');
      if (it) items.push(it);
    }
    subs = [];
    cur = null;
  };

  const { fnArray: fn, argsArray: args } = opList;
  for (let i = 0; i < fn.length && items.length < MAX_ITEMS; i++) {
    const a = args[i];
    switch (fn[i]) {
      case OPS.save: stack.push({ ...gs, ctm: gs.ctm.slice() }); break;
      case OPS.restore: if (stack.length) gs = stack.pop(); break;
      case OPS.transform: gs.ctm = mul(a, gs.ctm); break;
      case OPS.paintFormXObjectBegin:
        stack.push({ ...gs, ctm: gs.ctm.slice() });
        if (a && a[0]) gs.ctm = mul(a[0], gs.ctm);
        break;
      case OPS.paintFormXObjectEnd: if (stack.length) gs = stack.pop(); break;
      case OPS.setStrokeRGBColor: gs.stroke = hexColor(a); break;
      case OPS.setFillRGBColor: gs.fill = hexColor(a); break;
      case OPS.constructPath: build(a[0], a[1]); break;
      case OPS.stroke: flush('stroke'); break;
      case OPS.closeStroke: flush('stroke', true); break;
      case OPS.fill:
      case OPS.eoFill: flush('fill'); break;
      case OPS.fillStroke:
      case OPS.eoFillStroke: flush('both'); break;
      case OPS.closeFillStroke:
      case OPS.closeEOFillStroke: flush('both', true); break;
      case OPS.endPath: subs = []; cur = null; break;
      default: break;
    }
  }
  return items;
}

/* ---------- sharp zoom: re-render just the visible part at the current zoom ---------- */
let sharpTimer = 0;
let sharpTask = null;

function scheduleSharp() {
  clearTimeout(sharpTimer);
  if (sharpTask) { sharpTask.cancel(); sharpTask = null; }
  if (!S.pageObj || !S.bg) return;
  const nativeRes = S.bg.canvas.width / S.bg.w; // bitmap pixels per world unit
  if (S.view.k * dpr() <= nativeRes * 1.15) return; // the page image is already sharp enough
  sharpTimer = setTimeout(renderSharp, 250);
}

async function renderSharp() {
  const snap = { x: S.view.x, y: S.view.y, k: S.view.k, page: S.page };
  const d = dpr();
  const cw = Math.round(W * d);
  const ch = Math.round(H * d);
  if (!cw || !ch || cw * ch > 16e6) return;
  const c = document.createElement('canvas');
  c.width = cw;
  c.height = ch;
  const g = c.getContext('2d');
  g.fillStyle = '#fff';
  g.fillRect(0, 0, cw, ch);
  const scale = snap.k * d;
  const topWorldY = snap.y + H / snap.k;
  try {
    sharpTask = S.pageObj.render({
      canvasContext: g,
      viewport: S.pageObj.getViewport({ scale }),
      transform: [1, 0, 0, 1, -scale * snap.x, -scale * (S.bg.h - topWorldY)],
    });
    await sharpTask.promise;
  } catch { return; } // cancelled by a newer view
  sharpTask = null;
  if (S.page === snap.page && S.view.x === snap.x && S.view.y === snap.y && S.view.k === snap.k) {
    S.hi = { canvas: c, ...snap };
    requestDraw();
  }
}

/* ---------- pages ---------- */
async function showPage(n) {
  const page = await S.pdf.getPage(n);
  S.page = n;
  S.pageObj = page;
  say('Reading page…');
  const vp = page.getViewport({ scale: 1 });

  const rs = Math.min(2, 3000 / Math.max(vp.width, vp.height));
  const rvp = page.getViewport({ scale: rs });
  const c = document.createElement('canvas');
  c.width = Math.ceil(rvp.width);
  c.height = Math.ceil(rvp.height);
  const g = c.getContext('2d');
  g.fillStyle = '#fff';
  g.fillRect(0, 0, c.width, c.height);
  await page.render({ canvasContext: g, viewport: rvp }).promise;
  S.bg = { canvas: c, w: vp.width, h: vp.height };
  S.hi = null;

  setItems(extractVector(await page.getOperatorList(), vp));
  S.pts = [];
  S.box = null;
  S.sym = null;
  $('measure-out').hidden = true;
  $('cal-panel').hidden = true;
  $('empty').hidden = true;
  $('pdf-controls').hidden = false;
  $('page-info').textContent = `Page ${n} of ${S.pdf.numPages}`;
  $('page-jump').value = String(n);
  $('prev-page').disabled = n <= 1;
  $('next-page').disabled = n >= S.pdf.numPages;
  $('file-info').textContent = `${S.fileName} · ${S.items.length} lines`;
  renderScaleList();
  renderLayers();
  renderTakeoff();
  resize();
  fit();
  say(S.items.length
    ? 'Vector lines found. Set the scale below (preset or calibrate), then pick quantities or use the measuring tools.'
    : 'No vector lines on this page, so it is probably a scan. Set the scale, then measure with the tools above.');
}

async function openPdf(file) {
  if (typeof pdfjsLib === 'undefined') {
    say('PDF support did not load. Check your connection and reload the pane.');
    return;
  }
  try {
    const doc = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
    resetDocument(file);
    S.pdf = doc;
    const restored = restoreSession();
    showMethod('preset');
    setTool('pan');
    await showPage(1);
    if (restored) say(restored);
  } catch (err) {
    say('Could not read that PDF.');
    console.error(err);
  }
}

$('prev-page').addEventListener('click', () => showPage(S.page - 1));
$('next-page').addEventListener('click', () => showPage(S.page + 1));
$('page-jump').addEventListener('change', () => {
  const n = parseInt($('page-jump').value, 10);
  if (S.pdf && n >= 1 && n <= S.pdf.numPages) showPage(n); else $('page-jump').value = String(S.page);
});
$('show-bg').addEventListener('change', requestDraw);
