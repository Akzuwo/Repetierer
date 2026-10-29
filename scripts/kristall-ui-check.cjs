// Material/coordinate checks complement the real product flows in test:ui.
const { app, BrowserWindow, ipcMain } = require('electron');
if (!app) {
  const env = { ...process.env }; delete env.ELECTRON_RUN_AS_NODE;
  const result = require('node:child_process').spawnSync(require('electron'), [__filename], { env, stdio: 'inherit', windowsHide: true });
  process.exit(result.status ?? 1);
}
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const output = path.resolve(__dirname, '../.kristall-check');
fs.mkdirSync(output, { recursive: true });
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'repetierer-kristall-')));
const pause = ms => new Promise(resolve => setTimeout(resolve, ms));
ipcMain.handle('ai:status', () => ({ state: 'ready', message: 'Bereit' }));
ipcMain.handle('assessment:config-status', () => ({ ok: true, data: { configured: false } }));
ipcMain.handle('assessment:brevo-status', () => ({ ok: true, data: { configured: false } }));
ipcMain.on('get-settings', event => event.reply('settings-data', { theme: 'midnight', probabilityDecreaseFactor: .8, neverSelectedBoostFactor: 1.5, wiggersRuleDurationMinutes: 30 }, {}, {}));
ipcMain.on('version', event => event.reply('version', '7.0.0'));
ipcMain.on('load-saved-file', event => event.reply('saved-file-missing', ''));
ipcMain.on('get-active-teaching-session', event => event.reply('active-teaching-session-data', null));
let win;
const errors = [], results = [];
const evaluate = async code => { try { return await win.webContents.executeJavaScript(code, true); } catch (error) { console.error('Evaluation failed:', code); throw error; } };
async function until(code) { for (let i = 0; i < 150; i++) { if (await evaluate(code)) return; await pause(30); } throw Error('Timeout: ' + code); }
async function shot(name) { await pause(120); win.webContents.invalidate(); await pause(60); fs.writeFileSync(path.join(output, name + '.png'), (await win.webContents.capturePage()).toPNG()); }
const record = (name, details) => results.push({ name, details });
async function main() {
  await app.whenReady();
  win = new BrowserWindow({ show: false, width: 1100, height: 800, frame: false, webPreferences: { nodeIntegration: true, contextIsolation: false, offscreen: true, backgroundThrottling: false } });
  win.webContents.on('console-message', event => { if (event.level === 'error') errors.push(event.message); });
  await win.loadFile(path.resolve(__dirname, '../public/index.html'));
  await until(`window.repetiererKristall?.renderer?.ready && document.getElementById('session-open-file-btn').dataset.glass`);
  assert.deepEqual(await evaluate(`repetiererKristall.material`), { refraction: 1, frost: 0, tint: .08, brightness: 1, contrast: 0, focus: 0, motion: false });
  assert.equal(await evaluate(`repetiererKristall.renderer.gl.getError()`), 0);
  assert.equal(await evaluate(`document.querySelectorAll('#glass-canvas').length`), 1);
  assert.match(await evaluate(`getComputedStyle(document.querySelector('.crystal-scene')).backgroundImage`), /scene\.svg/);
  record('shader / one canvas / abstract background', true);
  await shot('app-home-desktop');
  assert.equal(await evaluate(`document.querySelectorAll('#title-bar button[data-glass], #drawer button[data-glass]').length`), 0);
  await evaluate(`document.getElementById('drawer-toggle').click()`); await pause(650);
  await shot('app-drawer-desktop');
  const drawer = await evaluate(`(() => { const d=document.getElementById('drawer').getBoundingClientRect(); const t=document.getElementById('drawer-title').getBoundingClientRect(); const i=document.querySelector('#drawer-main-view .nav-icon').getBoundingClientRect(); return {left:d.left,top:d.top,alignment:Math.abs(t.left-i.left)}; })()`);
  assert.ok(Math.abs(drawer.left)<1 && Math.abs(drawer.top-72)<1 && drawer.alignment<1, JSON.stringify(drawer));
  await evaluate(`document.getElementById('drawer-toggle').click()`); await pause(350);
  await evaluate(`document.getElementById('settings-btn').click()`);
  await until(`!document.getElementById('settings-view').classList.contains('update-hidden')`);
  await pause(600);
  await shot('app-settings-desktop');
  // Native keyboard activation must change both the checkbox and the lens.
  const start = await evaluate(`document.getElementById('extra-joker-setting').checked`);
  await evaluate(`document.getElementById('extra-joker-setting').focus()`);
  win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Space' });
  win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Space' });
  await until(`document.getElementById('extra-joker-setting').checked === ${!start}`);
  await until(`repetiererKristall.clock.active.size === 0`);
  assert.equal(await evaluate(`document.getElementById('extra-joker-setting').checked`), !start);
  const toggle = await evaluate(`(() => { const input = document.getElementById('extra-joker-setting'); const c = repetiererKristall.controls.get(input).controller; return { x: c.left.value, expected: input.checked ? 24 : 2, active: repetiererKristall.clock.active.size }; })()`);
  assert.ok(Math.abs(toggle.x - toggle.expected) < .02);
  assert.equal(toggle.active, 0);
  record('native toggle keyboard and spring settling', toggle);
  // Shared clock: press outside/release, then change preference mid-animation.
  await evaluate(`(() => { const c = repetiererKristall.controls.get(document.getElementById('extra-joker-setting')).controller; c.press(); window.dispatchEvent(new Event('blur')); })()`);
  await until(`repetiererKristall.clock.active.size === 0`);
  assert.equal(await evaluate(`repetiererKristall.controls.get(document.getElementById('extra-joker-setting')).controller.engagement.value`), 0);
  win.webContents.debugger.attach('1.3');
  await win.webContents.debugger.sendCommand('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-reduced-motion', value: 'no-preference' }] });
  await evaluate(`repetiererKristall.controls.get(document.getElementById('extra-joker-setting')).controller.press()`);
  await pause(25);
  await win.webContents.debugger.sendCommand('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-reduced-motion', value: 'reduce' }] });
  await pause(100);
  assert.deepEqual(await evaluate(`({ active: repetiererKristall.clock.active.size, frame: repetiererKristall.clock.frame })`), { active: 0, frame: 0 });
  record('outside release / live reduced motion', true);
  await win.webContents.debugger.sendCommand('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-reduced-motion', value: 'no-preference' }] });
  await evaluate(`window.closeRepetiererSettings()`);
  await pause(400);
  // Exercise the exact production nav/range controllers on a disposable fixture.
  await evaluate(`(() => {
    const fixture = document.createElement('section'); fixture.id = 'crystal-fixture';
    fixture.style.cssText = 'position:fixed;z-index:50;left:30.25px;top:120.5px;width:360.5px';
    fixture.innerHTML = '<nav class="session-tabs"><button class="btn-1">Eins</button><button class="btn-2">Zwei</button><button class="btn-2">Drei</button></nav><label style="display:block;margin-top:20px">Testregler<input type="range" id="crystal-range" min="0" max="100" value="0"></label>';
    document.body.append(fixture); repetiererKristall.refresh();
  })()`);
  await pause(100);
  await evaluate(`(() => { const nav = document.querySelector('#crystal-fixture nav'); const b = [...nav.querySelectorAll('button')]; b[0].className='btn-2'; b[2].className='btn-1'; repetiererKristall.refresh(); })()`);
  await pause(35);
  const intermediate = await evaluate(`(() => { const nav = document.querySelector('#crystal-fixture nav'); const m = repetiererKristall.navigations.get(nav).marker; return { active: repetiererKristall.clock.active.has(m), left: m.left.value, target: m.left.target }; })()`);
  assert.ok(intermediate.active && Math.abs(intermediate.left - intermediate.target) > .02);
  await evaluate(`(() => { const nav = document.querySelector('#crystal-fixture nav'); const b = [...nav.querySelectorAll('button')]; b[2].className='btn-2'; b[1].className='btn-1'; repetiererKristall.refresh(); })()`);
  await pause(1200);
  const alignment = await evaluate(`(() => { const nav = document.querySelector('#crystal-fixture nav'); const a = nav.querySelector('.liquid-marker').getBoundingClientRect(); const b = nav.querySelector('.btn-1').getBoundingClientRect(); return Math.max(Math.abs(a.left-b.left), Math.abs(a.right-b.right), Math.abs(a.top-b.top), Math.abs(a.bottom-b.bottom)); })()`);
  assert.ok(alignment < .02, 'Fractional nav geometry: ' + alignment);
  for (const value of [0, 50, 100]) {
    await evaluate(`(() => { const input = document.getElementById('crystal-range'); input.value = ${value}; input.dispatchEvent(new Event('input', {bubbles:true})); })()`);
    await pause(1000);
    const range = await evaluate(`(() => { const input = document.getElementById('crystal-range'); const c = repetiererKristall.controls.get(input).controller; return { left: c.left.value, target: c.left.target, fill: +input.parentElement.dataset.fill, center: (c.currentGeometry.left+c.currentGeometry.width/2)/parseFloat(getComputedStyle(input.parentElement).width) }; })()`);
    assert.ok(Math.abs(range.left - range.target) < .02);
    assert.ok(Math.abs(range.fill - range.center) < .00002);
  }
  await evaluate(`document.getElementById('crystal-range').focus()`);
  win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Left' }); win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Left' });
  await pause(100);
  assert.equal(await evaluate(`document.getElementById('crystal-range').value`), '99');
  record('interruptible navigation / fractional alignment / ranges 0,50,100 / keyboard', { intermediate, alignment });
  // The app hides scrollbars. Force one only in this disposable raster fixture
  // to retain coverage of optical alignment with a reduced layout viewport.
  assert.equal(await evaluate(`getComputedStyle(document.documentElement).scrollbarWidth`),'none');
  await evaluate(`(()=>{const s=document.createElement('style');s.id='raster-scrollbar-fixture';s.textContent='html{scrollbar-width:auto!important}html::-webkit-scrollbar{display:block!important;width:14px!important}';document.head.append(s);document.getElementById('crystal-fixture').remove();document.getElementById('container').style.visibility='hidden';document.documentElement.style.cssText='overflow-y:scroll!important;height:100%';document.body.style.cssText='overflow:visible!important;height:1600px';})()`);
  const raster = [];
  for (const dpr of [1, 1.25, 2]) {
    await win.webContents.debugger.sendCommand('Emulation.setDeviceMetricsOverride', { width: 1100, height: 800, deviceScaleFactor: dpr, mobile: false });
    await pause(150);
    const measurement = await evaluate(`(() => {
      const r = repetiererKristall.renderer, gl = r.gl;
      const neutral = document.createElement('canvas'); neutral.width=neutral.height=16; const cx=neutral.getContext('2d'); cx.fillStyle='#244b47'; cx.fillRect(0,0,16,16);
      gl.bindTexture(gl.TEXTURE_2D,r.texture); gl.texImage2D(gl.TEXTURE_2D,0,gl.RGB,gl.RGB,gl.UNSIGNED_BYTE,neutral); r.imageSize=[16,16]; r.resize(); r.draw();
      const bounds = r.canvas.getBoundingClientRect();
      const shape=document.createElement('div'); shape.style.cssText='position:fixed;left:90.25px;top:130.5px;width:140.5px;height:80.25px;border-radius:40px'; document.body.append(shape);
      const box=shape.getBoundingClientRect(), [px,py]=r.pixelRatio;
      const y=Math.floor((r.height-box.top-box.height/2)*py); const baseline=new Uint8Array(r.canvas.width*4); gl.readPixels(0,y,r.canvas.width,1,gl.RGBA,gl.UNSIGNED_BYTE,baseline);
      shape.dataset.glass='pill'; r.draw(); const line=new Uint8Array(r.canvas.width*4); gl.readPixels(0,y,r.canvas.width,1,gl.RGBA,gl.UNSIGNED_BYTE,line);
      const found=[]; for(let x=0;x<r.canvas.width;x++) if(Math.max(...[0,1,2].map(i=>Math.abs(line[x*4+i]-baseline[x*4+i])))>3) found.push(x);
      shape.remove();
      return {dpr:devicePixelRatio,cssWidth:r.width,innerWidth,scrollbar:innerWidth-document.documentElement.clientWidth,ratio:r.pixelRatio, pixels:found.length,offset:found.length?Math.abs((found[0]+found[found.length-1]+1)/(2*px)-(box.left+box.width/2)):999,glError:gl.getError()};
    })()`);
    assert.ok(Math.abs(measurement.dpr - dpr) < .00001);
    assert.ok(measurement.scrollbar > 0, 'Classic scrollbar is visible');
    assert.ok(measurement.pixels > 0);
    assert.ok(measurement.offset < 1.1, JSON.stringify(measurement));
    assert.equal(measurement.glError, 0);
    raster.push(measurement);
  }
  record('neutral texture raster / visible scrollbar / DPR 1,1.25,2', raster);
  await evaluate(`document.getElementById('raster-scrollbar-fixture').remove();document.documentElement.style.cssText='';document.body.style.cssText='';document.getElementById('container').style.visibility=''; repetiererKristall.renderer.loadScene('abstract')`);
  assert.equal(await evaluate(`getComputedStyle(document.documentElement).scrollbarWidth`),'none');
  for (const [width, height, name] of [[680,560,'compact'],[390,844,'mobile']]) {
    await win.webContents.debugger.sendCommand('Emulation.setDeviceMetricsOverride', { width, height, deviceScaleFactor: 1, mobile: false });
    await pause(250);
    const bounds = await evaluate(`(() => { const a=document.getElementById('session-empty-view').getBoundingClientRect(),h=document.getElementById('title-bar').getBoundingClientRect();return {inside:a.left>=0&&a.right<=innerWidth+1, header:document.getElementById('quit').getBoundingClientRect().right<=innerWidth+1};})()`);
    assert.ok(bounds.inside && bounds.header, JSON.stringify(bounds));
    await shot('app-home-' + name);
    await evaluate(`document.getElementById('settings-btn').click()`); await pause(650);
    const settings = await evaluate(`(() => { const e=document.querySelector('#settings-view .settings-content').getBoundingClientRect(); return {inside:e.left>=0&&e.right<=innerWidth+1&&e.top>=0&&e.bottom<=innerHeight+1};})()`);
    assert.ok(settings.inside);
    await shot('app-settings-' + name);
    await evaluate(`window.closeRepetiererSettings()`); await pause(250);
  }
  record('compact 680x560 / mobile 390x844 / settings overflow', true);
  // Native inputs and focus still work during loss and restoration of WebGL.
  await evaluate(`window.testContext = repetiererKristall.renderer.gl.getExtension('WEBGL_lose_context'); testContext.loseContext()`);
  await until(`document.documentElement.dataset.renderer==='css'`);
  await shot('app-fallback-mobile');
  await evaluate(`document.getElementById('settings-btn').click()`); await pause(500);
  const before = await evaluate(`document.getElementById('extra-joker-setting').checked`);
  await evaluate(`document.getElementById('extra-joker-setting').click()`);
  assert.equal(await evaluate(`document.getElementById('extra-joker-setting').checked`), !before);
  await evaluate(`testContext.restoreContext()`);
  await until(`document.documentElement.dataset.renderer==='webgl'`);
  record('context loss / usable CSS fallback / restoration', true);
  // No permanent render loop: once the clock settles, draws stop.
  await pause(1200);
  const idle = await evaluate(`new Promise(resolve=>{const r=repetiererKristall.renderer, original=r.draw;let draws=0;r.draw=function(){draws++;return original.call(this)};setTimeout(()=>{r.draw=original;resolve({draws,active:repetiererKristall.clock.active.size,frame:repetiererKristall.clock.frame})},250)})`);
  assert.equal(idle.active, 0); assert.equal(idle.frame, 0); assert.ok(idle.draws < 3, JSON.stringify(idle));
  record('idle clock / no permanent drawing loop', idle);
  assert.deepEqual(errors, []);
  fs.writeFileSync(path.join(output, 'material-tests.json'), JSON.stringify(results, null, 2));
  console.log('KRISTALL_UI_PASS', results.length, 'groups');
  win.webContents.debugger.detach(); win.destroy(); app.quit();
}
const watchdog=setTimeout(()=>{console.error('Kristall tests timed out');app.exit(1)},90000);
main().then(()=>clearTimeout(watchdog)).catch(error=>{console.error(error);clearTimeout(watchdog);app.exit(1)});
