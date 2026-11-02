/* global DxfParser, S, $, say, setItems, resetDocument, pathItem, circleItem, arcItem, insertItem,
   showMethod, renderScaleList, renderLayers, renderTakeoff, restoreSession, resize, fit, setTool */

const DXF_UNITS = { 1: 'in', 2: 'ft', 4: 'mm', 5: 'cm', 6: 'm' };

// A bulge is tan(sweep / 4) for the arc between two polyline vertices (positive = counter-clockwise).
function bulgeArc(p1, p2, b) {
  const theta = 4 * Math.atan(b);
  const dx = p2.x - p1.x;
  const dy = p2.y - p1.y;
  const c = Math.hypot(dx, dy);
  if (c === 0) return { len: 0, seg: 0, mid: [] };
  const r = c / (2 * Math.sin(Math.abs(theta) / 2));
  const off = Math.sign(b) * r * Math.cos(Math.abs(theta) / 2); // centre offset along the chord's left normal
  const cx = (p1.x + p2.x) / 2 + (-dy / c) * off;
  const cy = (p1.y + p2.y) / 2 + (dx / c) * off;
  const a0 = Math.atan2(p1.y - cy, p1.x - cx);
  const n = Math.max(2, Math.ceil(Math.abs(theta) / (Math.PI / 24)));
  const mid = [];
  for (let i = 1; i < n; i++) {
    const a = a0 + (theta * i) / n;
    mid.push([cx + r * Math.cos(a), cy + r * Math.sin(a)]);
  }
  return { len: r * Math.abs(theta), seg: (r * r / 2) * (theta - Math.sin(theta)), mid };
}

// Polyline with exact length and area, including curved (bulge) segments.
function dxfPolyline(layer, verts, closed) {
  if (!verts || verts.length < 2) return null;
  const n = verts.length;
  const pts = [[verts[0].x, verts[0].y]];
  let len = 0;
  let shoelace = 0;
  let segArea = 0;
  const segs = closed ? n : n - 1;
  for (let i = 0; i < segs; i++) {
    const p1 = verts[i];
    const p2 = verts[(i + 1) % n];
    shoelace += p1.x * p2.y - p2.x * p1.y;
    if (p1.bulge) {
      const a = bulgeArc(p1, p2, p1.bulge);
      len += a.len;
      segArea += a.seg;
      pts.push(...a.mid);
    } else {
      len += Math.hypot(p2.x - p1.x, p2.y - p1.y);
    }
    if (!(closed && i === n - 1)) pts.push([p2.x, p2.y]);
  }
  const area = closed ? Math.abs(shoelace / 2 + segArea) : 0;
  return pathItem(layer, pts, closed, len, area);
}

function dxfEllipse(layer, e) {
  const mx = e.majorAxisEndPoint.x;
  const my = e.majorAxisEndPoint.y;
  const ratio = e.axisRatio;
  let a0 = e.startAngle || 0;
  let a1 = e.endAngle == null ? 2 * Math.PI : e.endAngle;
  if (a1 <= a0) a1 += 2 * Math.PI;
  const full = Math.abs(a1 - a0 - 2 * Math.PI) < 1e-6;
  const n = Math.max(24, Math.ceil(((a1 - a0) / (2 * Math.PI)) * 96));
  const pts = [];
  for (let i = 0; i <= (full ? n - 1 : n); i++) {
    const t = a0 + ((a1 - a0) * i) / n;
    pts.push([e.center.x + Math.cos(t) * mx - Math.sin(t) * ratio * my, e.center.y + Math.cos(t) * my + Math.sin(t) * ratio * mx]);
  }
  return pathItem(layer, pts, full);
}

function toDxfItem(e) {
  const layer = e.layer || '0';
  switch (e.type) {
    case 'LINE': return pathItem(layer, (e.vertices || []).map((v) => [v.x, v.y]), false);
    case 'LWPOLYLINE':
    case 'POLYLINE': return dxfPolyline(layer, e.vertices, !!e.shape);
    case 'CIRCLE': return circleItem(layer, [e.center.x, e.center.y], e.radius);
    case 'ARC': return arcItem(layer, [e.center.x, e.center.y], e.radius, e.startAngle, e.endAngle);
    case 'ELLIPSE': return e.majorAxisEndPoint ? dxfEllipse(layer, e) : null;
    case 'SPLINE': return e.fitPoints && e.fitPoints.length > 1 ? pathItem(layer, e.fitPoints.map((v) => [v.x, v.y]), !!e.closed) : null;
    case 'INSERT': return insertItem(layer, e.name, [e.position.x, e.position.y]);
    default: return null;
  }
}

async function openDxf(file) {
  if (typeof DxfParser === 'undefined') { say('DXF support did not load. Check your connection and reload the pane.'); return; }
  try {
    say('Reading drawing…');
    const dxf = new DxfParser().parseSync(await file.text());
    const items = [];
    for (const e of dxf.entities || []) { const it = toDxfItem(e); if (it) items.push(it); }

    resetDocument(file);
    setItems(items);

    const fromFile = DXF_UNITS[dxf.header && dxf.header.$INSUNITS];
    S.rules = [{
      scope: 'all', factor: 1, unit: fromFile || 'ft',
      label: fromFile ? `1 drawing unit = 1 ${fromFile} (from file)` : '1 drawing unit = 1 ft (assumed)',
    }];
    const restored = restoreSession();
    showMethod('factor');
    $('empty').hidden = items.length > 0;
    $('file-info').textContent = `${file.name} · ${items.length} objects`;
    renderScaleList();
    renderLayers();
    renderTakeoff();
    setTool('pan');
    resize();
    fit();
    say(restored || (items.length
      ? (fromFile ? 'Drawing loaded. Units read from the file.' : 'Drawing loaded. The file has no units, so feet are assumed. Change the scale if needed.')
      : 'No supported objects found in this file.'));
  } catch (err) {
    say('Could not read that file. Save it as an ASCII DXF and try again.');
    console.error(err);
  }
}
