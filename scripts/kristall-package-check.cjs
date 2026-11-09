// Validate the actual archive, then launch its own EXE with a disposable profile.
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const net = require('node:net');
const { spawn } = require('node:child_process');
const assert = require('node:assert/strict');
const asar = require('@electron/asar');
const WebSocket = require('ws');
const root = path.resolve(__dirname, '..');
const output = path.join(root, '.kristall-check');
const build = path.join(root, fs.readFileSync(path.join(output, 'build-folder.txt'), 'utf8').trim());
const archive = path.join(build, 'win-unpacked/resources/app.asar');
const pause = ms => new Promise(resolve => setTimeout(resolve, ms));
let child, socket;
async function main() {
  const files = ['public/index.html', 'public/aiRenderer.js', 'public/uiMotion.js', 'public/kristall/glass.js', 'public/kristall/motion.js', 'public/kristall/integration.js', 'public/kristall/kristall.css', 'public/kristall/assets/scene.svg'];
  for (const file of files) assert.deepEqual(asar.extractFile(archive, path.normalize(file)), fs.readFileSync(path.join(root, file)), 'Packaged bytes: ' + file);
  const source = fs.readFileSync(path.join(root, 'design-reference/kristall/glass.js'), 'utf8').split('export class GlassRenderer')[0];
  const ported = fs.readFileSync(path.join(root, 'public/kristall/glass.js'), 'utf8').split('class GlassRenderer')[0];
  assert.ok(ported.includes('bevelWidth * 0.34') && !ported.includes('col += exp(-dot(spot'), 'Crystal lens preserved; pointer spotlight disabled');
  const before = JSON.parse(fs.readFileSync(path.join(output, 'previous-builds.json'), 'utf8')).filter(file => /\.(exe|blockmap|zip)$/i.test(file.Name));
  for (const file of before) {
    const stat = fs.statSync(path.join(root, 'builds', file.Name));
    assert.equal(stat.size, file.Length, 'Previous build preserved: ' + file.Name);
    assert.ok(Math.abs(stat.mtimeMs - new Date(file.LastWriteTimeUtc).getTime()) <= 1, 'Previous timestamp preserved (NTFS rounding): ' + file.Name);
  }
  const server = net.createServer(); await new Promise(resolve => server.listen(0, '127.0.0.1', resolve)); const port = server.address().port; await new Promise(resolve => server.close(resolve));
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'repetierer-packaged-kristall-'));
  const env = { ...process.env }; delete env.ELECTRON_RUN_AS_NODE;
  child = spawn(path.join(build, 'win-unpacked/Repetierer.exe'), [`--remote-debugging-port=${port}`, `--user-data-dir=${profile}`], { env, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
  const logs = []; child.stdout.on('data', data => logs.push(data.toString())); child.stderr.on('data', data => logs.push(data.toString()));
  child.on('error', error => logs.push(error.message));
  let target;
  for (let i = 0; i < 180; i++) {
    try { const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json(); target = targets.find(item => item.url.endsWith('/public/index.html')); if (target) break; } catch {}
    await pause(150);
  }
  assert.ok(target, 'The packaged main window starts: ' + logs.join('').slice(-2000));
  socket = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((resolve,reject) => { socket.once('open', resolve); socket.once('error', reject); });
  let sequence = 0; const pending = new Map(); const errors = [];
  socket.on('message', raw => { const event = JSON.parse(raw); if (event.id) { const callback = pending.get(event.id); if (callback) { pending.delete(event.id); event.error ? callback.reject(Error(JSON.stringify(event.error))) : callback.resolve(event.result); } } else if (event.method === 'Runtime.exceptionThrown') errors.push(event.params.exceptionDetails.text); });
  const command = (method, params = {}) => new Promise((resolve,reject) => { const id=++sequence; pending.set(id,{resolve,reject}); socket.send(JSON.stringify({id,method,params})); });
  const evaluate = async expression => { const reply=await command('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true}); assert.ok(!reply.exceptionDetails, JSON.stringify(reply.exceptionDetails)); return reply.result.value; };
  await command('Runtime.enable');
  for (let i=0;i<100;i++) { if(await evaluate(`!!window.repetiererKristall?.renderer?.ready`)) break; await pause(100); }
  const state=await evaluate(`({renderer:document.documentElement.dataset.renderer,ready:repetiererKristall.renderer.ready,material:repetiererKristall.material,url:location.href,scene:getComputedStyle(document.querySelector('.crystal-scene')).backgroundImage,glError:repetiererKristall.renderer.gl.getError()})`);
  assert.equal(state.renderer,'webgl'); assert.equal(state.ready,true); assert.equal(state.glError,0); assert.match(state.scene,/scene\.svg/);
  await pause(400);
  fs.writeFileSync(path.join(output,'packaged-home.png'),Buffer.from((await command('Page.captureScreenshot',{format:'png'})).data,'base64'));
  await evaluate(`document.getElementById('settings-btn').click()`); await pause(700);
  assert.equal(await evaluate(`document.documentElement.dataset.crystalOverlay`),'true');
  assert.equal(await evaluate(`getComputedStyle(document.getElementById('content')).visibility`),'hidden');
  fs.writeFileSync(path.join(output,'packaged-settings.png'),Buffer.from((await command('Page.captureScreenshot',{format:'png'})).data,'base64'));
  assert.deepEqual(errors,[]);
  fs.writeFileSync(path.join(output,'package-tests.json'),JSON.stringify({files,oldBuildsPreserved:before.length,state,profile,errors},null,2));
  console.log('KRISTALL_PACKAGE_PASS',files.length,'files,',before.length,'previous build artifacts preserved');
  socket.close(); socket=null; child.kill(); child=null;
}
const watchdog=setTimeout(()=>{console.error('Package check timed out');socket?.close();child?.kill();process.exit(1)},60000);
main().then(()=>clearTimeout(watchdog)).catch(error=>{console.error(error);clearTimeout(watchdog);socket?.close();child?.kill();process.exitCode=1});
