/* global S, $, say, toScreen, ctx, requestDraw, commit, addItem, setTool, renderTakeoff */
// Finds look-alikes of a symbol the user boxes on the page image (normalized cross-correlation).
// Works on the page bitmap, so it handles vector and scanned PDFs alike. It does not handle
// rotated or resized copies of the symbol.

// Self-contained on purpose: this function's source is also run inside a Web Worker.
function nccSearch(img, iw, ih, tpl, tw, th, minScore) {
  const n = tw * th;
  let mean = 0;
  for (let i = 0; i < n; i++) mean += tpl[i];
  mean /= n;
  const t = new Float32Array(n);
  let tn = 0;
  for (let i = 0; i < n; i++) { t[i] = tpl[i] - mean; tn += t[i] * t[i]; }
  if (tn < 1e-6) return [];
  tn = Math.sqrt(tn);

  // integral images give each window's sum and sum of squares in O(1)
  const w1 = iw + 1;
  const S1 = new Float64Array(w1 * (ih + 1));
  const S2 = new Float64Array(w1 * (ih + 1));
  for (let y = 0; y < ih; y++) {
    let r1 = 0;
    let r2 = 0;
    for (let x = 0; x < iw; x++) {
      const v = img[y * iw + x];
      r1 += v; r2 += v * v;
      S1[(y + 1) * w1 + x + 1] = S1[y * w1 + x + 1] + r1;
      S2[(y + 1) * w1 + x + 1] = S2[y * w1 + x + 1] + r2;
    }
  }

  const found = [];
  for (let y = 0; y <= ih - th; y++) {
    for (let x = 0; x <= iw - tw; x++) {
      const a = y * w1 + x;
      const b = a + tw;
      const c = (y + th) * w1 + x;
      const d = c + tw;
      const s1 = S1[d] - S1[b] - S1[c] + S1[a];
      const v = S2[d] - S2[b] - S2[c] + S2[a] - (s1 * s1) / n;
      if (v < 1e-3 * n) continue; // flat window (blank paper)
      let num = 0;
      for (let j = 0; j < th; j++) {
        const row = (y + j) * iw + x;
        const tr = j * tw;
        for (let i = 0; i < tw; i++) num += img[row + i] * t[tr + i];
      }
      const score = num / (tn * Math.sqrt(v));
      if (score >= minScore) found.push({ x, y, s: score });
    }
  }

  // keep the best hit in each neighbourhood
  found.sort((p, q) => q.s - p.s);
  if (found.length > 5000) found.length = 5000;
  const keep = [];
  const r = Math.max(tw, th) * 0.5;
  for (const f of found) {
    let ok = true;
    for (const k of keep) if (Math.abs(k.x - f.x) < r && Math.abs(k.y - f.y) < r) { ok = false; break; }
    if (ok) keep.push(f);
  }
  return keep;
}

function grayDownsample(canvas, ds) {
  const w = canvas.width;
  const h = canvas.height;
  const px = canvas.getContext('2d').getImageData(0, 0, w, h).data;
  const gw = Math.floor(w / ds);
  const gh = Math.floor(h / ds);
  const out = new Uint8Array(gw * gh);
  for (let y = 0; y < gh; y++) {
    for (let x = 0; x < gw; x++) {
      let s = 0;
      for (let j = 0; j < ds; j++) {
        let o = ((y * ds + j) * w + x * ds) * 4;
        for (let i = 0; i < ds; i++) { s += px[o] * 0.299 + px[o + 1] * 0.587 + px[o + 2] * 0.114; o += 4; }
      }
      out[y * gw + x] = s / (ds * ds);
    }
  }
  return { data: out, w: gw, h: gh };
}

function runNcc(g, tpl, tw, th, min) {
  return new Promise((resolve) => {
    const direct = () => resolve(nccSearch(g.data, g.w, g.h, tpl, tw, th, min));
    try {
      const src = `self.onmessage=(e)=>{const d=e.data;self.postMessage((${nccSearch.toString()})(d.img,d.iw,d.ih,d.tpl,d.tw,d.th,d.min));};`;
      const url = URL.createObjectURL(new Blob([src], { type: 'text/javascript' }));
      const worker = new Worker(url);
      worker.onmessage = (e) => { resolve(e.data); worker.terminate(); URL.revokeObjectURL(url); };
      worker.onerror = () => { worker.terminate(); URL.revokeObjectURL(url); direct(); };
      worker.postMessage({ img: g.data, iw: g.w, ih: g.h, tpl, tw, th, min });
    } catch { direct(); }
  });
}

const symThreshold = () => parseInt($('sym-thresh').value, 10) / 100;
const symVisible = () => (S.sym ? S.sym.cands.filter((c) => c.s >= symThreshold()) : []);

function updateSymbolInfo() {
  const vis = symVisible();
  const on = vis.filter((c) => !c.off).length;
  $('sym-thresh-val').textContent = `${$('sym-thresh').value}%`;
  $('sym-info').textContent = S.sym
    ? `${on} match${on === 1 ? '' : 'es'}${vis.length !== on ? ` (${vis.length - on} ignored)` : ''}. Click a box to ignore or restore it.`
    : 'Drag a box around one example of the symbol.';
  $('sym-add').disabled = on === 0;
}

async function symbolDrag(box) {
  S.box = null;
  if (!S.bg) { say('Symbol search needs the page image. Open a PDF first.'); requestDraw(); return; }
  const sx = S.bg.canvas.width / S.bg.w;
  const sy = S.bg.canvas.height / S.bg.h;
  const bw = (box.x1 - box.x0) * sx;
  const bh = (box.y1 - box.y0) * sy;
  if (Math.min(bw, bh) < 8) { say('That box is too small. Drag around the whole symbol.'); requestDraw(); return; }

  say('Searching the page…');
  await new Promise((r) => setTimeout(r, 30)); // let the status paint
  const ds = Math.max(1, Math.round(Math.max(bw, bh) / 28));
  const g = grayDownsample(S.bg.canvas, ds);
  const tw = Math.max(4, Math.min(48, Math.round(bw / ds)));
  const th = Math.max(4, Math.min(48, Math.round(bh / ds)));
  const tx = Math.min(Math.max(0, Math.round((box.x0 * sx) / ds)), g.w - tw);
  const ty = Math.min(Math.max(0, Math.round(((S.bg.h - box.y1) * sy) / ds)), g.h - th);
  const tpl = new Uint8Array(tw * th);
  for (let j = 0; j < th; j++) for (let i = 0; i < tw; i++) tpl[j * tw + i] = g.data[(ty + j) * g.w + tx + i];

  const hits = await runNcc(g, tpl, tw, th, 0.5);
  S.sym = {
    page: S.page,
    ww: (tw * ds) / sx,
    wh: (th * ds) / sy,
    cands: hits.map((h) => ({
      x: ((h.x + tw / 2) * ds) / sx,
      y: S.bg.h - ((h.y + th / 2) * ds) / sy,
      s: h.s,
      off: false,
    })),
  };
  $('symbol-panel').hidden = false;
  updateSymbolInfo();
  const on = symVisible().length;
  say(on ? `Found ${on} possible match${on === 1 ? '' : 'es'}. Adjust sensitivity, then add them as a count.` : 'No matches. Try lowering the sensitivity or boxing the symbol more tightly.');
  requestDraw();
}

function symbolClick(sx, sy) {
  if (!S.sym) return;
  let best = null;
  let bd = 16;
  for (const c of symVisible()) {
    const [x, y] = toScreen(c.x, c.y);
    const d = Math.hypot(x - sx, y - sy);
    if (d < bd) { bd = d; best = c; }
  }
  if (best) { best.off = !best.off; updateSymbolInfo(); requestDraw(); }
}

function drawSymbols(flag) {
  if (!S.sym || S.sym.page !== S.page) return;
  ctx.lineWidth = 1.5;
  for (const c of symVisible()) {
    const [x, y] = toScreen(c.x, c.y);
    const w = Math.max(10, S.sym.ww * S.view.k);
    const h = Math.max(10, S.sym.wh * S.view.k);
    ctx.strokeStyle = flag;
    ctx.globalAlpha = c.off ? 0.35 : 1;
    ctx.setLineDash(c.off ? [3, 3] : []);
    ctx.strokeRect(x - w / 2, y - h / 2, w, h);
  }
  ctx.globalAlpha = 1;
  ctx.setLineDash([]);
}

$('sym-thresh').addEventListener('input', () => { updateSymbolInfo(); requestDraw(); });
$('sym-clear').addEventListener('click', () => { S.sym = null; updateSymbolInfo(); requestDraw(); });
$('sym-add').addEventListener('click', () => {
  const pts = symVisible().filter((c) => !c.off).map((c) => [c.x, c.y]);
  if (!pts.length) return;
  commit(() => { S.sel = addItem('count', pts, { source: 'symbol' }).id; });
  S.sym = null;
  updateSymbolInfo();
  say(`Added a count of ${pts.length}. You can rename it in the Takeoff list.`);
});
updateSymbolInfo();
