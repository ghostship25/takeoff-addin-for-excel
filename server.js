// Tiny HTTPS dev server. Office add-ins must load over HTTPS.
const https = require('https');
const fs = require('fs');
const path = require('path');
const os = require('os');

const certDir = path.join(os.homedir(), '.office-addin-dev-certs');
const opts = {
  key: fs.readFileSync(path.join(certDir, 'localhost.key')),
  cert: fs.readFileSync(path.join(certDir, 'localhost.crt')),
};
const root = path.join(__dirname, 'public');
const types = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png', '.json': 'application/json' };

https.createServer(opts, (req, res) => {
  const url = req.url.split('?')[0];
  const file = path.join(root, url === '/' ? 'taskpane.html' : url);
  if (!file.startsWith(root)) { res.writeHead(403); return res.end(); }
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404); return res.end('Not found'); }
    res.writeHead(200, { 'Content-Type': types[path.extname(file)] || 'application/octet-stream' });
    res.end(data);
  });
}).listen(3000, () => console.log('Takeoff add-in running at https://localhost:3000'));
