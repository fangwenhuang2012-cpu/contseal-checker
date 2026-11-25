const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');

const PORT = 8080;
const ROOT = __dirname;

const mimeTypes = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon'
};

function getLocalIp() {
  const nets = os.networkInterfaces();
  for (const name of Object.keys(nets)) {
    for (const net of nets[name]) {
      if (net.family === 'IPv4' && !net.internal) {
        return net.address;
      }
    }
  }
  return 'localhost';
}

const server = http.createServer((req, res) => {
  const urlPath = req.url.split('?')[0];

  // API Log
  if (urlPath === '/api/log' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => { body += chunk; });
    req.on('end', () => {
      console.log('[CLIENT LOG]', body);
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: true }));
    });
    return;
  }

  // API Archive
  if (urlPath === '/api/archive' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => { body += chunk; });
    req.on('end', () => {
      try {
        const parsed = JSON.parse(body);
        const stamp = new Date().toISOString().replace(/[:.]/g, '-');
        const rawVessel = parsed.vesselName || 'ChuyenTau';
        const safeVessel = rawVessel.replace(/\.(xlsx|xls|csv)$/i, '').replace(/[^A-Za-z0-9_-]/g, '_') || 'ChuyenTau';

        const archiveDir = path.join(ROOT, 'Archives', `Tau_${safeVessel}_${stamp}`);
        const imgDir = path.join(archiveDir, 'Images');
        fs.mkdirSync(imgDir, { recursive: true });

        fs.writeFileSync(path.join(archiveDir, `Archive_Tau_${safeVessel}_${stamp}.json`), body, 'utf8');
        if (parsed.htmlReport) {
          fs.writeFileSync(path.join(archiveDir, `BaoCao_Tau_${safeVessel}_${stamp}.html`), parsed.htmlReport, 'utf8');
        }

        let savedImgCount = 0;
        if (parsed.records && Array.isArray(parsed.records)) {
          parsed.records.forEach(rec => {
            const safeCont = (rec.contNo || 'cont').replace(/[^A-Za-z0-9_-]/g, '_');
            const safeId = (rec.id || 'id').toString().replace(/[^A-Za-z0-9_-]/g, '_');
            if (rec.contImage && rec.contImage.includes('base64,')) {
              const b64 = rec.contImage.split('base64,')[1];
              fs.writeFileSync(path.join(imgDir, `${safeCont}_Cont_${safeId}.jpg`), Buffer.from(b64, 'base64'));
              savedImgCount++;
            }
            if (rec.sealImage && rec.sealImage.includes('base64,')) {
              const b64 = rec.sealImage.split('base64,')[1];
              fs.writeFileSync(path.join(imgDir, `${safeCont}_Seal_${safeId}.jpg`), Buffer.from(b64, 'base64'));
              savedImgCount++;
            }
          });
        }
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ ok: true, folder: archiveDir, imageCount: savedImgCount }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  // Serve static files
  let safePath = urlPath === '/' ? '/index.html' : urlPath;
  const filePath = path.join(ROOT, safePath);

  fs.stat(filePath, (err, stats) => {
    if (err || !stats.isFile()) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('404 Not Found');
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = mimeTypes[ext] || 'application/octet-stream';

    res.writeHead(200, { 'Content-Type': contentType });
    fs.createReadStream(filePath).pipe(res);
  });
});

server.listen(PORT, '0.0.0.0', () => {
  const localIp = getLocalIp();
  console.log('==========================================================');
  console.log(`  ContSeal Checker v29 dang chay!`);
  console.log(`  - May tinh (Local):  http://localhost:${PORT}`);
  console.log(`  - Dien thoai (Wi-Fi): http://${localIp}:${PORT}`);
  console.log('==========================================================');
});
