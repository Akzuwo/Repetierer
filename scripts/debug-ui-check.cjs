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
const sentMails=[];let failMail=false;
require('nodemailer').createTransport=()=>({sendMail:async message=>{if(failMail){const error=new Error('not exposed');error.code='EAUTH';throw error;}sentMails.push(message);return {accepted:[message.to]};}});
const evaluate = code => win.webContents.executeJavaScript(code,true);
async function waitFor(code) {for(let i=0;i<150;i++){try {if(await evaluate(code))return;}catch{} await pause(40);} console.error(await evaluate(`({phase:document.getElementById('content').dataset.pageMotion,context:window.getRepetiererAiContext?.(),header:document.getElementById('session-header').className,end:document.getElementById('end-session-btn').disabled,modal:document.getElementById('session-end-modal').className})`));throw new Error('Timeout: '+code);}
async function shot(name) {console.log('DEBUG capture:',name);await pause(650);win.webContents.invalidate();await pause(100);fs.writeFileSync(path.join(output,name+'.png'),(await win.webContents.capturePage()).toPNG());}
async function mouseClick(selector) {
  const point = await evaluate(`(()=>{const button=document.querySelector(${JSON.stringify(selector)});button.scrollIntoView({block:'center'});const r=button.getBoundingClientRect();return {x:Math.round(r.x+r.width/2),y:Math.round(r.y+r.height/2)};})()`);
  win.webContents.sendInputEvent({type:'mouseDown',button:'left',clickCount:1,...point});
  win.webContents.sendInputEvent({type:'mouseUp',button:'left',clickCount:1,...point});
}
async function choose(id) {
  await evaluate("notify.clear();document.getElementById('debug-launcher').click()");
  await waitFor("document.getElementById('debug-panel').classList.contains('modal-open')");await pause(450);
  await mouseClick('[data-debug-scenario="'+id+'"]');
  await waitFor("!document.getElementById('debug-panel').classList.contains('modal-open') && notify.getActive().length===1");
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
  if (process.env.REPETIERER_DEBUG_VISIBLE !== '1') {
    process.env.REPETIERER_UI_TEST = '1';
    app.on('browser-window-created',(_event,window)=>{window.once('ready-to-show',()=>setTimeout(()=>window.hide(),10));});
  }
  require('../app.js');
  for(let i=0;i<150;i++) {win=BrowserWindow.getAllWindows().find(w=>w!==normalWindow && w.webContents.getURL().endsWith('/public/index.html'));if(win)break;await pause(50);}
  assert.ok(win,'Actual app main window');
  win.webContents.setBackgroundThrottling(false);
  normalWindow.destroy();
  win.webContents.on('console-message',event=>{if(event.level==='error' || event.level===3) {errors.push(event.message);console.error('Renderer:',event.message);}});
  await waitFor(`!!document.getElementById('debug-launcher')`);
  console.log('DEBUG: real startup and launcher ready');
  assert.equal(await evaluate(`process.argv.includes('--repetierer-debug')`),true);
  // Exercise the real editor IPC before exercising the message-only debug menu.
  const nativeWorkbook=path.join(app.getPath('userData'),'editor-ipc.xlsx');
  const nativeBook=new (require('exceljs').Workbook)();
  for(const name of ['Editor A','Editor B']) {const sheet=nativeBook.addWorksheet(name);sheet.getCell('A1').value='repetierer';sheet.getCell('A6').value=name;sheet.getCell('B6').value=5;sheet.getCell('H6').value=1;}
  await nativeBook.xlsx.writeFile(nativeWorkbook);
  const nativeProgram=require('../src/program');await new Promise(resolve=>nativeProgram.setFile(nativeWorkbook,resolve));
  const nativeData=await evaluate(`new Promise(resolve=>{const ipc=require('electron').ipcRenderer;ipc.once('editor-persons-data',(_event,people,meta)=>resolve({people,meta}));ipc.send('get-editor-persons',{className:'Editor B',requestId:741});})`);
  assert.equal(nativeData.meta.className,'Editor B');assert.equal(nativeData.meta.requestId,741);assert.equal(nativeData.people[0].name,'Editor B');
  const nativeSave=await evaluate(`new Promise(resolve=>{const ipc=require('electron').ipcRenderer;ipc.once('editor-persons-saved',(_event,result,meta)=>resolve({result,meta}));ipc.send('save-editor-persons',[{id:0,name:'Editor B geändert',joker:3}],{className:'Editor B',requestId:742,filePath:${JSON.stringify(nativeWorkbook)}});})`);
  assert.equal(nativeSave.result.success,true);assert.equal(nativeSave.result.persons[0].name,'Editor B geändert');assert.equal(nativeSave.meta.requestId,742);
  assert.equal(nativeProgram.getCurrentClass(),undefined);assert.equal(nativeProgram.getClassPersons('Editor A')[0].name,'Editor A');
  console.log('EDITOR_MAIN_IPC_PASS: independent class read/save, real workbook, request metadata, no repetition selection');
  // Three message actions plus the requested update and mail tools.
  // Full app/dialog scenarios remain covered by test:ui and focused checks.
  await evaluate("window.debugDocumentSentinel='same-document';document.getElementById('debug-launcher').click()");await pause(500);
  assert.deepEqual(await evaluate("[...document.querySelectorAll('#debug-panel button')].map(b=>b.textContent)"),['Nachricht','Warnung','Fehler','Fake Update','Update mit Neustart','Testmail senden']);
  assert.equal(await evaluate("document.querySelectorAll('#debug-actions button').length"),1);
  assert.equal(await evaluate("document.querySelectorAll('#debug-panel input').length"),1);
  await shot('debug-three-actions');
  await evaluate("document.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}))");
  await waitFor("getComputedStyle(document.getElementById('debug-panel')).display==='none'");
  for(const type of ['info','warning','error']) {
    await choose('notification-'+type);
    assert.equal(await evaluate("notify.getActive()[0].type"),type);
    assert.equal(await evaluate("window.debugDocumentSentinel"),'same-document','No reload or demo page');
    assert.equal(await evaluate("document.documentElement.dataset.debugSimulation"),undefined);
    assert.equal(await evaluate("document.getElementById('title-bar').checkVisibility({visibilityProperty:true})"),true);
  }
  assert.equal(nativeProgram.getClassPersons('Editor B')[0].name,'Editor B geändert');
  await evaluate("document.getElementById('debug-launcher').click()");await pause(500);
  await evaluate("document.getElementById('debug-mail-send').focus()");
  win.webContents.sendInputEvent({type:'keyDown',keyCode:'Tab'});win.webContents.sendInputEvent({type:'keyUp',keyCode:'Tab'});
  await waitFor("document.activeElement===document.querySelector('#debug-panel button')");
  // The form goes through the real IPC, config validation and production template.
  fs.writeFileSync(require('../src/storage').getPaths().integrationEnvPath,[
    'BREVO_SMTP_HOST=smtp.example.ch','BREVO_SMTP_PORT=587','BREVO_SMTP_USER=fixture-user',
    'BREVO_SMTP_PASSWORD=fixture-only','BREVO_FROM_EMAIL=no-reply@example.ch','BREVO_FROM_NAME=Fixture'
  ].join('\n'));
  await evaluate(`notify.clear();document.getElementById('debug-mail-recipient').value='demo@example.ch'`);
  await mouseClick('#debug-mail-send');await waitFor(`notify.getActive().some(n=>n.type==='success')`);
  assert.equal(sentMails.length,1);assert.equal(sentMails[0].to,'demo@example.ch');assert.match(sentMails[0].html,/<table/);
  failMail=true;await evaluate(`notify.clear()`);await mouseClick('#debug-mail-send');await waitFor(`notify.getActive().some(n=>n.type==='error')`);
  assert.equal(sentMails.length,1);assert.equal(await evaluate(`document.getElementById('debug-mail-send').disabled`),false);
  await shot('debug-menu-mail');
  await evaluate(`document.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}))`);await pause(500);
  // Preview approval completes locally; cancellation cannot leave an armed restart.
  await evaluate(`document.getElementById('debug-launcher').click()`);await pause(450);
  await mouseClick('[data-debug-scenario="update-preview"]');await waitFor(`document.getElementById('update-panel').classList.contains('modal-open')`);await pause(450);
  assert.equal(await evaluate(`document.getElementById('update-title').innerText`),'Fake Update');
  await mouseClick('#update-primary-btn');await waitFor(`notify.getActive().some(n=>n.title==='Fake Update')`);
  assert.equal(nativeProgram.getClassPersons('Editor B')[0].name,'Editor B geändert');
  const debugFlagIndex=process.argv.indexOf('--debug');process.argv.splice(debugFlagIndex,1);
  const blocked=await evaluate(`require('electron').ipcRenderer.invoke('debug:send-test-mail',{recipient:'demo@example.ch'})`);
  assert.equal(blocked.ok,false);assert.match(blocked.error.message,/nur im Debug/);assert.equal(sentMails.length,1);
  process.argv.push('--debug');
  assert.deepEqual(errors,[],'No renderer errors');
  console.log('DEBUG_UI_PASS: real startup, normal-mode gate, three message actions, update preview, testmail success/error with local SMTP interception, no reload/demo state, workbook unchanged, keyboard focus');
}
const watchdog=setTimeout(()=>{console.error('Debug UI check timed out');app.exit(1);},120000);
main().then(()=>{clearTimeout(watchdog);app.exit(0);}).catch(error=>{console.error(error);clearTimeout(watchdog);app.exit(1);});
