/* global polyLen, polyArea */
// Pure geometry for the measuring shapes. No DOM. Every shape returns
//   { poly, closed, len, area }
// poly is a tessellated outline for drawing and snapping. len and area are analytic wherever a closed
// form exists (arcs, circles, ellipses, sectors, segments), so curved quantities do not depend on how
// finely the outline is drawn. For closed shapes, len is the perimeter.

const TAU = Math.PI * 2;
const norm = (a) => ((a % TAU) + TAU) % TAU;
const dist = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1]);

const openGeom = (poly) => ({ poly, closed: false, len: polyLen(poly, false), area: 0 });
const closedGeom = (poly) => ({ poly, closed: true, len: polyLen(poly, true), area: polyArea(poly) });

/* ---------- circles and arcs ---------- */
function circumcircle(p1, p2, p3) {
  // work relative to p1 so large coordinates (survey grids) keep their precision
  const bx = p2[0] - p1[0], by = p2[1] - p1[1];
  const cx = p3[0] - p1[0], cy = p3[1] - p1[1];
  const d = 2 * (bx * cy - by * cx);
  if (Math.abs(d) < 1e-9 * Math.hypot(bx, by) * Math.hypot(cx, cy)) return null; // collinear
  const b2 = bx * bx + by * by;
  const c2 = cx * cx + cy * cy;
  const ux = (cy * b2 - by * c2) / d;
  const uy = (bx * c2 - cx * b2) / d;
  return { c: [p1[0] + ux, p1[1] + uy], r: Math.hypot(ux, uy) };
}

function arcPoints(c, r, a0, theta) {
  const n = Math.max(6, Math.ceil(Math.abs(theta) / (TAU / 180)));
  const out = [];
  for (let i = 0; i <= n; i++) {
    const a = a0 + (theta * i) / n;
    out.push([c[0] + r * Math.cos(a), c[1] + r * Math.sin(a)]);
  }
  return out;
}

// Arc from p1 through p2 to p3.
function arc3(p1, p2, p3) {
  const cc = circumcircle(p1, p2, p3);
  if (!cc) return openGeom([p1, p2, p3]);
  const a1 = Math.atan2(p1[1] - cc.c[1], p1[0] - cc.c[0]);
  const a3 = Math.atan2(p3[1] - cc.c[1], p3[0] - cc.c[0]);
  const turnsLeft = (p2[0] - p1[0]) * (p3[1] - p1[1]) - (p2[1] - p1[1]) * (p3[0] - p1[0]) > 0;
  const theta = turnsLeft ? norm(a3 - a1) : -norm(a1 - a3);
  return { poly: arcPoints(cc.c, cc.r, a1, theta), closed: false, len: cc.r * Math.abs(theta), area: 0, c: cc.c, r: cc.r, theta };
}

// The same arc closed by its chord: a circular segment.
function segment3(p1, p2, p3) {
  const a = arc3(p1, p2, p3);
  if (a.r === undefined) return openGeom([p1, p2, p3]);
  const th = Math.abs(a.theta);
  return { poly: a.poly, closed: true, len: a.len + dist(p1, p3), area: ((a.r * a.r) / 2) * (th - Math.sin(th)) };
}

// Pie slice: centre, start of the arc, then the direction of the end (counter-clockwise unless flipped).
function sector3(c, p2, p3, flip) {
  const r = dist(c, p2);
  if (r === 0) return openGeom([c, p2]);
  const a0 = Math.atan2(p2[1] - c[1], p2[0] - c[0]);
  const a1 = Math.atan2(p3[1] - c[1], p3[0] - c[0]);
  let th = norm(a1 - a0);
  if (flip) th -= TAU;
  const t = Math.abs(th);
  return { poly: [c, ...arcPoints(c, r, a0, th)], closed: true, len: 2 * r + r * t, area: 0.5 * r * r * t };
}

function circle2(c, edge) {
  const r = dist(c, edge);
  if (r === 0) return openGeom([c, edge]);
  const poly = [];
  for (let i = 0; i < 144; i++) poly.push([c[0] + r * Math.cos((TAU * i) / 144), c[1] + r * Math.sin((TAU * i) / 144)]);
  return { poly, closed: true, len: TAU * r, area: Math.PI * r * r };
}

// Ramanujan's second approximation (error below 1e-9 for ordinary aspect ratios).
function ellipsePerimeter(a, b) {
  const h = ((a - b) / (a + b)) ** 2;
  return Math.PI * (a + b) * (1 + (3 * h) / (10 + Math.sqrt(4 - 3 * h)));
}

// Axis-aligned ellipse inscribed in the box with opposite corners p1 and p2.
function ellipse2(p1, p2) {
  const rx = Math.abs(p2[0] - p1[0]) / 2;
  const ry = Math.abs(p2[1] - p1[1]) / 2;
  if (rx === 0 || ry === 0) return openGeom([p1, p2]);
  const cx = (p1[0] + p2[0]) / 2;
  const cy = (p1[1] + p2[1]) / 2;
  const poly = [];
  for (let i = 0; i < 144; i++) poly.push([cx + rx * Math.cos((TAU * i) / 144), cy + ry * Math.sin((TAU * i) / 144)]);
  return { poly, closed: true, len: ellipsePerimeter(rx, ry), area: Math.PI * rx * ry };
}

/* ---------- rectangles ---------- */
const rect2 = (p1, p2) => closedGeom([p1, [p2[0], p1[1]], p2, [p1[0], p2[1]]]);

// One side given by p1-p2, then a third point that sets the width.
function rotatedRect3(p1, p2, p3) {
  const L = dist(p1, p2);
  if (L === 0) return openGeom([p1, p2]);
  const nx = -(p2[1] - p1[1]) / L;
  const ny = (p2[0] - p1[0]) / L;
  const w = (p3[0] - p1[0]) * nx + (p3[1] - p1[1]) * ny;
  return closedGeom([p1, p2, [p2[0] + nx * w, p2[1] + ny * w], [p1[0] + nx * w, p1[1] + ny * w]]);
}

/* ---------- curves ---------- */
function sampleBezier(p, n = 96) {
  const out = [];
  for (let i = 0; i <= n; i++) {
    const t = i / n;
    const u = 1 - t;
    out.push([
      u * u * u * p[0][0] + 3 * u * u * t * p[1][0] + 3 * u * t * t * p[2][0] + t * t * t * p[3][0],
      u * u * u * p[0][1] + 3 * u * u * t * p[1][1] + 3 * u * t * t * p[2][1] + t * t * t * p[3][1],
    ]);
  }
  return out;
}

// Centripetal Catmull-Rom spline through the points: no loops or overshoot with uneven spacing.
function spline(points, per = 20) {
  const pts = points.filter((p, i) => i === 0 || dist(p, points[i - 1]) > 0); // drop repeats
  if (pts.length < 3) return pts;
  const ext = [
    [2 * pts[0][0] - pts[1][0], 2 * pts[0][1] - pts[1][1]],
    ...pts,
    [2 * pts.at(-1)[0] - pts.at(-2)[0], 2 * pts.at(-1)[1] - pts.at(-2)[1]],
  ];
  const lerp = (a, b, ta, tb, t) => {
    const w = (t - ta) / (tb - ta);
    return [a[0] + (b[0] - a[0]) * w, a[1] + (b[1] - a[1]) * w];
  };
  const out = [pts[0]];
  for (let i = 1; i < ext.length - 2; i++) {
    const [p0, p1, p2, p3] = [ext[i - 1], ext[i], ext[i + 1], ext[i + 2]];
    const t0 = 0;
    const t1 = t0 + Math.sqrt(dist(p0, p1));
    const t2 = t1 + Math.sqrt(dist(p1, p2));
    const t3 = t2 + Math.sqrt(dist(p2, p3));
    for (let k = 1; k <= per; k++) {
      const t = t1 + ((t2 - t1) * k) / per;
      const a1 = lerp(p0, p1, t0, t1, t);
      const a2 = lerp(p1, p2, t1, t2, t);
      const a3 = lerp(p2, p3, t2, t3, t);
      out.push(lerp(lerp(a1, a2, t0, t2, t), lerp(a2, a3, t1, t3, t), t1, t2, t));
    }
  }
  return out;
}

/* ---------- helpers ---------- */
function segDist(p, a, b) {
  const dx = b[0] - a[0], dy = b[1] - a[1];
  const l2 = dx * dx + dy * dy;
  const u = l2 ? Math.max(0, Math.min(1, ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / l2)) : 0;
  return Math.hypot(p[0] - (a[0] + u * dx), p[1] - (a[1] + u * dy));
}

function polyDist(poly, p, closed) {
  let d = Infinity;
  for (let i = 1; i < poly.length; i++) d = Math.min(d, segDist(p, poly[i - 1], poly[i]));
  if (closed && poly.length > 2) d = Math.min(d, segDist(p, poly.at(-1), poly[0]));
  return d;
}

function pointInPoly(poly, p) {
  let inside = false;
  for (let i = 0, j = poly.length - 1; i < poly.length; j = i++) {
    const [xi, yi] = poly[i];
    const [xj, yj] = poly[j];
    if ((yi > p[1]) !== (yj > p[1]) && p[0] < ((xj - xi) * (p[1] - yi)) / (yj - yi) + xi) inside = !inside;
  }
  return inside;
}

// Douglas-Peucker: thins a freehand trace without changing its shape beyond tol.
function simplify(pts, tol) {
  if (pts.length < 3) return pts;
  const keep = new Uint8Array(pts.length);
  keep[0] = 1; keep[pts.length - 1] = 1;
  const stack = [[0, pts.length - 1]];
  while (stack.length) {
    const [s, e] = stack.pop();
    let md = 0;
    let mi = -1;
    for (let i = s + 1; i < e; i++) {
      const d = segDist(pts[i], pts[s], pts[e]);
      if (d > md) { md = d; mi = i; }
    }
    if (md > tol) { keep[mi] = 1; stack.push([s, mi], [mi, e]); }
  }
  return pts.filter((_, i) => keep[i]);
}

// Points every `step` along a path, starting at its first point; optionally one more at the very end.
function markersAlong(poly, step, endPost = true) {
  if (!(step > 0) || poly.length < 2) return [];
  const out = [poly[0]];
  let next = step;
  let walked = 0;
  for (let i = 1; i < poly.length; i++) {
    const a = poly[i - 1];
    const b = poly[i];
    const seg = dist(a, b);
    if (seg === 0) continue;
    while (next <= walked + seg + 1e-9) {
      const t = Math.min(1, (next - walked) / seg);
      out.push([a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t]);
      next += step;
    }
    walked += seg;
  }
  const end = poly.at(-1);
  if (endPost && dist(out.at(-1), end) > step * 1e-6) out.push(end);
  return out;
}

function angleDeg(a, b, c) {
  const v1 = [a[0] - b[0], a[1] - b[1]];
  const v2 = [c[0] - b[0], c[1] - b[1]];
  const l = Math.hypot(...v1) * Math.hypot(...v2);
  if (!l) return 0;
  return (Math.acos(Math.max(-1, Math.min(1, (v1[0] * v2[0] + v1[1] * v2[1]) / l))) * 180) / Math.PI;
}

/* ---------- dispatcher ---------- */
// kind is 'length' | 'area' | 'count'. Fewer points than a shape needs gives a plain polyline (for previews).
function geomFor(kind, shape, pts, o = {}) {
  const n = pts.length;
  switch (shape) {
    case 'line': return openGeom(pts.slice(0, 2));
    case 'polyline':
    case 'spaced': return openGeom(pts);
    case 'arc': return n >= 3 ? arc3(pts[0], pts[1], pts[2]) : openGeom(pts);
    case 'bezier': return n >= 4 ? openGeom(sampleBezier(pts.slice(0, 4))) : openGeom(pts);
    case 'spline': return openGeom(spline(pts));
    case 'freehand': return kind === 'area' ? closedGeom(pts) : openGeom(pts);
    case 'polygon': return n >= 3 ? closedGeom(pts) : openGeom(pts);
    case 'rect': return n >= 2 ? rect2(pts[0], pts[1]) : openGeom(pts);
    case 'rrect': return n >= 3 ? rotatedRect3(pts[0], pts[1], pts[2]) : openGeom(pts);
    case 'circle': return n >= 2 ? circle2(pts[0], pts[1]) : openGeom(pts);
    case 'ellipse': return n >= 2 ? ellipse2(pts[0], pts[1]) : openGeom(pts);
    case 'sector': return n >= 3 ? sector3(pts[0], pts[1], pts[2], o.flip) : openGeom(pts);
    case 'segment': return n >= 3 ? segment3(pts[0], pts[1], pts[2]) : openGeom(pts);
    default: return openGeom(pts);
  }
}

// What each tool offers. need = points that complete the shape on their own; open = finished with Enter / double-click.
const SHAPES = {
  length: {
    line: { label: 'Straight line', need: 2 },
    polyline: { label: 'Polyline', open: true, min: 2 },
    arc: { label: 'Arc (3 points)', need: 3 },
    bezier: { label: 'Curve (Bézier, 4 points)', need: 4 },
    spline: { label: 'Spline (through points)', open: true, min: 2 },
    freehand: { label: 'Freehand', drag: true },
  },
  area: {
    polygon: { label: 'Polygon', open: true, min: 3 },
    rect: { label: 'Rectangle (2 corners)', need: 2 },
    rrect: { label: 'Rotated rectangle (3 points)', need: 3 },
    circle: { label: 'Circle (centre, edge)', need: 2 },
    ellipse: { label: 'Ellipse (2 corners)', need: 2 },
    sector: { label: 'Sector (centre, start, end)', need: 3 },
    segment: { label: 'Arc segment (3 points)', need: 3 },
    freehand: { label: 'Freehand', drag: true },
  },
  count: {
    markers: { label: 'Markers', open: true, min: 1 },
    spaced: { label: 'At spacing along a path', open: true, min: 2 },
  },
};
