// Loads libraries (a local copy in /vendor first, the CDN second), then the app scripts in order.
// Run `npm run vendor` once to download the local copies so the add-in works offline.
window.PDF_LOCAL = false;
const LIBS = [
  { local: 'vendor/dxf-parser.js', cdn: 'https://cdn.jsdelivr.net/npm/dxf-parser@1.1.2/dist/dxf-parser.js', ok: () => typeof DxfParser !== 'undefined' },
  { local: 'vendor/pdf.min.js', cdn: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js', ok: () => typeof pdfjsLib !== 'undefined', flag: 'PDF_LOCAL' },
];
const APP = ['core.js', 'shapes.js', 'dxf.js', 'scale.js', 'takeoff.js', 'symbols.js', 'pdf.js', 'main.js'];

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src;
    s.onload = resolve;
    s.onerror = () => reject(new Error(`Could not load ${src}`));
    document.head.append(s);
  });
}

async function loadLib(lib) {
  try { await loadScript(lib.local); if (lib.ok()) { if (lib.flag) window[lib.flag] = true; return; } } catch { /* fall through to CDN */ }
  try { await loadScript(lib.cdn); } catch (err) { console.warn(err.message); }
}

(async () => {
  for (const lib of LIBS) await loadLib(lib);
  for (const file of APP) await loadScript(file);
})().catch((err) => {
  document.getElementById('status').textContent = `Startup problem: ${err.message}`;
});
