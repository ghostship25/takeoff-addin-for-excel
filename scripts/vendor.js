// Downloads pinned copies of the libraries into public/vendor so the add-in can run offline.
const https = require('https');
const fs = require('fs');
const path = require('path');

const FILES = {
  'dxf-parser.js': 'https://cdn.jsdelivr.net/npm/dxf-parser@1.1.2/dist/dxf-parser.js',
  'pdf.min.js': 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js',
  'pdf.worker.min.js': 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js',
};
const dir = path.join(__dirname, '..', 'public', 'vendor');
fs.mkdirSync(dir, { recursive: true });

function get(url, dest, hops = 0) {
  https.get(url, (res) => {
    if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location && hops < 5) {
      return get(new URL(res.headers.location, url).href, dest, hops + 1);
    }
    if (res.statusCode !== 200) { console.error(`FAILED ${res.statusCode} ${url}`); return; }
    res.pipe(fs.createWriteStream(dest)).on('finish', () => console.log(`saved ${path.basename(dest)}`));
  }).on('error', (e) => console.error(`FAILED ${url}: ${e.message}`));
}
for (const [name, url] of Object.entries(FILES)) get(url, path.join(dir, name));
