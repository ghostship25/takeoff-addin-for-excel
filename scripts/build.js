// Production build for hosting / AppSource. Produces ./dist:
//   dist/public/       static files to upload to your HTTPS host (one app.js bundle, no CDN libraries)
//   dist/manifest.xml  the manifest pointing at your host
//
// Usage: npm run build -- https://addin.example.com --email support@example.com --provider "Your Company"
const fs = require('fs');
const path = require('path');

const args = process.argv.slice(2);
const opt = (name) => { const i = args.indexOf(`--${name}`); return i >= 0 ? args[i + 1] : undefined; };
const host = (args.find((a) => /^https?:\/\//.test(a)) || 'https://localhost:3000').replace(/\/$/, '');
const email = opt('email');
const provider = opt('provider');
if (!host.startsWith('https://')) { console.error('The host must start with https://'); process.exit(1); }

const root = path.join(__dirname, '..');
const pub = path.join(root, 'public');
const dist = path.join(root, 'dist');
const out = path.join(dist, 'public');
const read = (p) => fs.readFileSync(p, 'utf8');

// libraries must be bundled for a store build
const vendor = ['dxf-parser.js', 'pdf.min.js', 'pdf.worker.min.js'];
const missing = vendor.filter((f) => !fs.existsSync(path.join(pub, 'vendor', f)));
if (missing.length) {
  console.error(`Missing ${missing.join(', ')} in public/vendor.\nRun "npm run vendor" first (needs internet), then build again.`);
  process.exit(1);
}

fs.rmSync(dist, { recursive: true, force: true });
fs.mkdirSync(path.join(out, 'vendor'), { recursive: true });

// 1. one script: libraries + app files, in the same order the dev loader uses
const order = [...read(path.join(pub, 'loader.js')).match(/const APP = \[([^\]]*)\]/)[1].matchAll(/'([^']+)'/g)].map((m) => m[1]);
let app = order.map((f) => read(path.join(pub, f))).join('\n;\n');
let minified = false;
try {
  app = require('esbuild').transformSync(app, { minify: true, target: 'es2020' }).code;
  minified = true;
} catch { /* esbuild not installed: ship readable code */ }
const libs = ['dxf-parser.js', 'pdf.min.js'].map((f) => read(path.join(pub, 'vendor', f))).join('\n;\n');
fs.writeFileSync(path.join(out, 'app.js'), `window.PDF_LOCAL=true;\n${libs}\n;\n${app}\n`);
fs.copyFileSync(path.join(pub, 'vendor', 'pdf.worker.min.js'), path.join(out, 'vendor', 'pdf.worker.min.js'));

// 2. page, styles, icons, support and privacy pages
fs.writeFileSync(path.join(out, 'taskpane.html'), read(path.join(pub, 'taskpane.html')).replace('<script src="loader.js"></script>', '<script src="app.js"></script>'));
fs.copyFileSync(path.join(pub, 'taskpane.css'), path.join(out, 'taskpane.css'));
for (const f of fs.readdirSync(pub).filter((n) => /^icon-\d+\.png$/.test(n))) fs.copyFileSync(path.join(pub, f), path.join(out, f));
const today = new Date().toISOString().slice(0, 10);
for (const page of ['support.html', 'privacy.html']) {
  fs.writeFileSync(path.join(out, page), read(path.join(pub, page)).replaceAll('{{SUPPORT_EMAIL}}', email || '{{SUPPORT_EMAIL}}').replaceAll('{{DATE}}', today));
}

// 3. manifest
let manifest = read(path.join(root, 'manifest.xml')).replaceAll('https://localhost:3000', host);
if (provider) manifest = manifest.replace('<ProviderName>Takeoff</ProviderName>', `<ProviderName>${provider.replace(/[<&]/g, '')}</ProviderName>`);
fs.writeFileSync(path.join(dist, 'manifest.xml'), manifest);

const kb = (p) => `${Math.round(fs.statSync(p).size / 1024)} KB`;
console.log(`Built dist/ for ${host}`);
console.log(`  dist/public/app.js  ${kb(path.join(out, 'app.js'))}${minified ? ' (minified)' : ' (not minified: run "npm install" to add esbuild)'}`);
console.log('  dist/public         upload this folder to your HTTPS host');
console.log('  dist/manifest.xml   sideload or submit this');
if (!email) console.warn('\nNo --email given: support.html and privacy.html still contain {{SUPPORT_EMAIL}}.');
if (host.includes('localhost')) console.warn('\nHost is localhost: fine for testing, not for AppSource.');
