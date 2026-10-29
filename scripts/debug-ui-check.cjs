const {app, BrowserWindow, ipcMain} = require('electron');
if (!app) {
  const env = {...process.env}; delete env.ELECTRON_RUN_AS_NODE;
  const result = require('node:child_process').spawnSync(require('electron'), [__filename], {env, stdio:'inherit', windowsHide:true});
  process.exit(result.status ?? 1);
}
const assert = require('node:assert/strict'), fs = require('node:fs'), path = require('node:path'), os = require('node:os');
const {isDebugMode} = require('../src/debugFlags');
const pause = ms => new Promise(resolve => setTimeout(resolve,ms));
const output = path.join(__dirname,'..','.ui-polish-smoke'); fs.mkdirSync(output,{recursive:true});
app.setPath('userData',fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-debug-ui-')));
let win, errors = [], writes = 0;
const evaluate = code => win.webContents.executeJavaScript(code,true);
async function waitFor(code) {for(let i=0;i<150;i++){try {if(await evaluate(code))return;}catch{} await pause(40);} console.error(await evaluate(`({phase:document.getElementById('content').dataset.pageMotion,context:window.getRepetiererAiContext?.(),header:document.getElementById('session-header').className,end:document.getElementById('end-session-btn').disabled,modal:document.getElementById('session-end-modal').className})`));throw new Error('Timeout: '+code);}
async function shot(name) {console.log('DEBUG capture:',name);await pause(650);win.webContents.invalidate();await pause(100);fs.writeFileSync(path.join(output,name+'.png'),(await win.webContents.capturePage()).toPNG());}
async function choose(id) {
  console.log('DEBUG choose:', id);
  await evaluate(`sessionStorage.setItem('repetierer-debug-scenario',${JSON.stringify(id)})`);
  const loaded = new Promise(resolve=>win.webContents.once('did-finish-load',resolve));
  win.reload();
  await loaded; await waitFor(`document.getElementById('debug-launcher')?.textContent === 'Debug · Simulation'`);
  console.log('DEBUG fixture ready:',id);
}
async function main() {
  for(const flag of ['--debug','--debug-mode','--dev-mode']) assert.equal(isDebugMode(['app',flag]),true);
  assert.equal(isDebugMode(['app']),false);
  await app.whenReady();
  // A forged saved scenario cannot enable simulation or the menu without flag.
  const onMode = event=>event.reply('debug-mode',false);
  ipcMain.on('get-debug-mode',onMode); ipcMain.handle('ai:status',()=>({state:'unavailable'}));
  win = new BrowserWindow({show:false,webPreferences:{nodeIntegration:true,contextIsolation:false}});
  await win.loadFile(path.join(__dirname,'..','public','index.html')); await pause(350);
  assert.equal(await evaluate(`!!document.getElementById('debug-launcher')`),false);
  await evaluate(`sessionStorage.setItem('repetierer-debug-scenario','migration-preview')`);
  await win.reload(); await pause(350);
  assert.equal(await evaluate(`document.documentElement.dataset.debugSimulation`),undefined);
  await evaluate(`sessionStorage.removeItem('repetierer-debug-scenario')`);
  const normalWindow = win; ipcMain.removeListener('get-debug-mode',onMode); ipcMain.removeHandler('ai:status');
  // Boot the real entry point, using its real flag handling and window prefs.
  process.argv.push('--debug');
  process.env.REPETIERER_UI_TEST = '1';
  app.on('browser-window-created',(_event,window)=>{window.once('ready-to-show',()=>setTimeout(()=>window.hide(),10));});
  require('../app.js');
  for(let i=0;i<150;i++) {win=BrowserWindow.getAllWindows().find(w=>w!==normalWindow && w.webContents.getURL().endsWith('/public/index.html'));if(win)break;await pause(50);}
  assert.ok(win,'Actual app main window');
  win.webContents.setBackgroundThrottling(false);
  normalWindow.destroy();
  win.webContents.on('console-message',event=>{if(event.level==='error' || event.level===3) {errors.push(event.message);console.error('Renderer:',event.message);}});
  await waitFor(`!!document.getElementById('debug-launcher')`);
  console.log('DEBUG: real startup and launcher ready');
  assert.equal(await evaluate(`process.argv.includes('--repetierer-debug')`),true);
  await evaluate(`document.getElementById('debug-launcher').click()`); await pause(350);
  console.log('DEBUG: menu opened');
  assert.equal(await evaluate(`document.getElementById('problem-report-btn').checkVisibility()`),false,'Problem reports are temporarily absent from the drawer');
  const disabledReply=await evaluate(`new Promise(resolve=>{const ipc=require('electron').ipcRenderer;ipc.once('problem-report-result',(_event,result)=>resolve(result));ipc.send('send-problem-report',{comment:'must not send'});})`);
  assert.deepEqual(disabledReply,{ok:false,error:'feature-disabled'},'The real main process rejects problem reports before sending');
  const modalIds=await evaluate(`[...document.querySelectorAll('.modal:not(#debug-panel)')].filter(m=>m.id!=='problem-report-modal').map(m=>m.id)`);
  assert.equal(await evaluate(`document.querySelectorAll('[data-debug-scenario^="modal:"]').length`),modalIds.length);
  console.log('DEBUG: modal catalog verified');
  await shot('debug-scenario-menu');
  console.log('DEBUG: scenario menu captured');
  await evaluate(`(()=>{const s=document.querySelector('#debug-panel input');s.value='Änderungsbestätigung';s.dispatchEvent(new Event('input'));})()`);
  assert.equal(await evaluate(`document.querySelectorAll('[data-debug-scenario]').length`),1);
  for(const channel of ['run-joker-date-migration','update-download-approved','add-participation-entry','__debug-test-write']) ipcMain.on(channel,()=>writes++);
  await choose('ai-confirmation'); await waitFor(`!!document.querySelector('.ai-confirm-card')`);
  await evaluate(`document.querySelector('.ai-confirm-actions button:last-child').click()`);
  await waitFor(`!document.querySelector('.ai-confirm-card') && !!document.querySelector('.ai-assistant')`);
  await choose('ai-loading'); await waitFor(`document.getElementById('ai-activity').dataset.busy === 'true'`);
  await choose('ai-error'); await waitFor(`!!document.querySelector('.ai-error')`);
  await choose('migration-locked'); await waitFor(`document.getElementById('joker-date-migration-status').textContent.includes('gesperrt')`);
  await evaluate(`document.getElementById('joker-date-migration-run').click()`); await pause(120);
  assert.match(await evaluate(`document.getElementById('joker-date-migration-status').textContent`),/gesperrt/);
  await choose('migration-preview'); await evaluate(`document.getElementById('joker-date-migration-run').click()`);
  await waitFor(`document.getElementById('joker-date-migration-notice').textContent.includes('18')`);
  await choose('update-available'); await waitFor(`!document.getElementById('update-panel').classList.contains('update-hidden')`);
  await evaluate(`document.getElementById('update-primary-btn').click()`);
  await waitFor(`document.getElementById('update-progress-bar').style.width === '62%'`);
  await evaluate(`require('electron').ipcRenderer.send('__debug-test-write')`); await pause(100); assert.equal(writes,0,'Simulated writes never reach main');
  for(const id of modalIds) {
    console.log('DEBUG modal:',id);
    await choose('modal:'+id);
    await waitFor(`document.getElementById(${JSON.stringify(id)}).classList.contains('modal-open')`);
    assert.equal(await evaluate(`!!document.getElementById(${JSON.stringify(id)}).querySelector('.modal-close-action')`),false,'No X close button: '+id);
    await evaluate(`document.getElementById(${JSON.stringify(id)}).click()`);
    await waitFor(`getComputedStyle(document.getElementById(${JSON.stringify(id)})).display==='none'`);
  }
  await choose('settings'); await waitFor(`!document.getElementById('settings-view').classList.contains('update-hidden')`);
  assert.equal(await evaluate(`document.getElementById('settings-view').closest('.modal')`),null);
  await evaluate(`document.getElementById('settings-back-btn').click()`);
  await waitFor(`document.getElementById('settings-view').classList.contains('update-hidden')`);
  await choose('attendance'); await waitFor(`document.querySelectorAll('.attendance-person').length===25`);
  await evaluate(`document.querySelector('.attendance-person').click()`);
  assert.equal(await evaluate(`document.querySelector('.attendance-person').getAttribute('aria-pressed')`),'true');
  await choose('participation'); await waitFor(`document.querySelectorAll('.participation-student').length===25`);
  win.showInactive();
  await waitFor(`document.visibilityState==='visible'`);
  await pause(650);
  await evaluate(`document.getElementById('session-repetition-tab').click()`);
  await waitFor(`document.getElementById('content').dataset.pageMotion==='enter'`);
  await pause(600);
  await evaluate(`document.getElementById('session-participation-tab').click()`); await pause(650);
  // Interrupt the real dock navigation, then inspect settled styles/lenses.
  await evaluate(`document.getElementById('session-repetition-tab').click()`); await pause(70);
  await evaluate(`document.getElementById('session-participation-tab').click()`); await pause(70);
  await evaluate(`document.getElementById('session-repetition-tab').click()`);
  await waitFor(`document.getElementById('session-repetition-tab').getAttribute('aria-pressed')==='true'`);
  await pause(700);
  assert.equal(await evaluate(`document.getElementById('repetition').classList.contains('update-hidden')`),false);
  assert.equal(await evaluate(`require('./vendor/ui-motion.js').gsap.getTweensOf(document.getElementById('content')).length`),0);
  await evaluate(`document.getElementById('session-participation-tab').click();document.getElementById('session-repetition-tab').click();document.getElementById('session-participation-tab').click()`);
  await waitFor(`document.getElementById('session-participation-tab').getAttribute('aria-pressed')==='true'`);
  await pause(650);
  assert.equal(await evaluate(`document.getElementById('participation-view').classList.contains('update-hidden')`),false,'The last same-tick navigation wins');

  win.webContents.debugger.attach('1.3');
  await win.webContents.debugger.sendCommand('Emulation.setEmulatedMedia',{features:[{name:'prefers-reduced-motion',value:'reduce'}]});
  await evaluate(`document.getElementById('session-participation-tab').click()`);
  assert.equal(await evaluate(`document.getElementById('session-participation-tab').getAttribute('aria-pressed')`),'true');
  await evaluate(`document.getElementById('debug-launcher').click()`);await pause(100);
  const loaded = new Promise(resolve=>win.webContents.once('did-finish-load',resolve));
  await evaluate(`document.getElementById('debug-reset').click()`).catch(()=>{});await loaded;
  await waitFor(`document.getElementById('debug-launcher')?.textContent==='Debug'`);
  assert.equal(await evaluate(`document.documentElement.dataset.debugSimulation`),undefined);
  assert.deepEqual(errors,[],'No renderer errors');
  console.log(`DEBUG_UI_PASS: real --debug startup, normal-mode gate, ${modalIds.length} modals, AI confirmation/loading/error, migration lock/retry/success, update download, isolated writes, search/reset, interrupted navigation and reduced motion`);
}
const watchdog=setTimeout(()=>{console.error('Debug UI check timed out');app.exit(1);},120000);
main().then(()=>{clearTimeout(watchdog);app.exit(0);}).catch(error=>{console.error(error);clearTimeout(watchdog);app.exit(1);});
