const { app, BrowserWindow } = require('electron');
if (!app) {
  const env = { ...process.env };
  delete env.ELECTRON_RUN_AS_NODE;
  const result = require('node:child_process').spawnSync(require('electron'), [__filename], { env, stdio: 'inherit', windowsHide: true });
  process.exit(result.status ?? 1);
}
const fs = require('node:fs');
const path = require('node:path');
const http = require('node:http');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '../design-reference/kristall');
const output = path.resolve(__dirname, '../.kristall-check');
fs.mkdirSync(output, { recursive: true });
const errors = [];
let server;
app.whenReady().then(async () => {
  server = http.createServer((req, res) => {
    const scene = req.url === '/assets/background.jpg';
    const file = scene ? path.resolve(__dirname, '../public/kristall/assets/scene.svg') : path.resolve(root, '.' + decodeURIComponent(req.url === '/' ? '/index.html' : req.url));
    if (!scene && !file.startsWith(root + path.sep)) { res.writeHead(403).end(); return; }
    const type = { '.js': 'text/javascript', '.css': 'text/css', '.html': 'text/html', '.svg': 'image/svg+xml' }[path.extname(file)];
    fs.readFile(file, (err, data) => { res.writeHead(err ? 404 : 200, { 'Content-Type': type || 'application/octet-stream' }); res.end(err ? '' : data); });
  });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const win = new BrowserWindow({ show: false, width: 1280, height: 900, webPreferences: { offscreen: true, backgroundThrottling: false } });
  win.webContents.on('console-message', event => { if (event.level === 'error') errors.push(event.message); });
  await win.loadURL(`http://127.0.0.1:${server.address().port}`);
  await new Promise(resolve => setTimeout(resolve, 1200));
  const state = await win.webContents.executeJavaScript(`({ renderer: document.documentElement.dataset.renderer, ready: crystalReference.renderer.ready, material: crystalReference.material, width: crystalReference.renderer.width, canvas: document.querySelector('canvas').getBoundingClientRect().width })`);
  assert.equal(state.renderer, 'webgl');
  assert.equal(state.ready, true);
  assert.equal(state.width, state.canvas);
  assert.deepEqual(errors, []);
  fs.writeFileSync(path.join(output, 'reference-desktop.png'), (await win.webContents.capturePage()).toPNG());
  fs.writeFileSync(path.join(output, 'reference.json'), JSON.stringify(state, null, 2));
  console.log('KRISTALL_REFERENCE_PASS', state);
  server.close();
  win.destroy();
  app.quit();
}).catch(error => { console.error(error); server?.close(); app.exit(1); });
