const {app,BrowserWindow,dialog}=require('electron');
if(!app){const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;const r=require('node:child_process').spawnSync(require('electron'),[__filename],{env,windowsHide:true,stdio:'inherit',timeout:90000});process.exit(r.status??1);}
const fs=require('node:fs'),path=require('node:path'),os=require('node:os'),assert=require('node:assert/strict');
const profile=fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-doc-ui-'));app.setPath('userData',profile);
let win,mode='save',calls=0,release;const errors=[],pause=ms=>new Promise(r=>setTimeout(r,ms));
const evaluate=code=>win.webContents.executeJavaScript(code);
async function until(code){for(let i=0;i<150;i++){if(await evaluate(code))return;await pause(50);}throw new Error('Timeout: '+code);}
async function click(selector){await evaluate(`document.querySelector(${JSON.stringify(selector)}).scrollIntoView({block:'center'})`);await pause(180);const p=await evaluate(`(()=>{const r=document.querySelector(${JSON.stringify(selector)}).getBoundingClientRect();return {x:Math.round(r.x+r.width/2),y:Math.round(r.y+r.height/2)}})()`);win.webContents.sendInputEvent({type:'mouseDown',button:'left',clickCount:1,...p});win.webContents.sendInputEvent({type:'mouseUp',button:'left',clickCount:1,...p});await pause(80);}
const settle=()=>until(`!document.getElementById('drawer-content').dataset.viewMotion && !document.getElementById('drawer').dataset.motionState`);
async function shot(name){const dir=path.join(__dirname,'../.ui-polish-smoke');fs.mkdirSync(dir,{recursive:true});fs.writeFileSync(path.join(dir,name+'.png'),(await win.webContents.capturePage()).toPNG());}
app.whenReady().then(async()=>{
 require('../src/mainProcess');
 // Only the native save picker is substituted; renderer, IPC and file writes are real.
 const root=app.getAppPath.bind(app);app.getAppPath=()=>path.resolve(__dirname,'..');
 dialog.showSaveDialog=async(_parent,options)=>{
  calls++;assert.deepEqual(options.filters[0].extensions,['md']);
  if(mode==='wait')return new Promise(resolve=>release=resolve);
  if(mode==='cancel')return {canceled:true};
  return {filePath:mode==='fail'?path.join(profile,'missing','api.md'):path.join(profile,'API.md')};
 };
 win=new BrowserWindow({width:1100,height:800,show:false,frame:false,webPreferences:{nodeIntegration:true,contextIsolation:false,offscreen:true}});
 win.webContents.on('console-message',event=>{if(event.level==='error'||event.level===3)errors.push(event.message);});
 await win.loadFile(path.join(__dirname,'../public/index.html'));win.showInactive();win.webContents.focus();await pause(300);
 await click('#drawer-toggle');await settle();await click('#drawer-documentation-view-btn');await settle();
 assert.equal(await evaluate(`document.getElementById('drawer-title').textContent`),'Dokumentationen');
 assert.equal(await evaluate(`document.getElementById('drawer-documentation-view').parentElement.id`),'drawer-content');
 assert.equal(await evaluate(`document.getElementById('documentation-manual-btn').disabled`),true);
 assert.equal(await evaluate(`document.getElementById('documentation-manual-status').checkVisibility()`),true);
 await click('#documentation-manual-btn');assert.equal(calls,0,'Unavailable manual does not create an empty document');
 mode='wait';await click('#documentation-api-btn');assert.equal(calls,1);assert.equal(await evaluate(`document.getElementById('documentation-api-btn').disabled`),true);
 await click('#documentation-api-btn');assert.equal(calls,1,'Double click cannot open another picker');release({canceled:true});await until(`!document.getElementById('documentation-api-btn').disabled`);
 assert.equal(fs.existsSync(path.join(profile,'API.md')),false);
 mode='save';win.show();win.focus();win.webContents.focus();await pause(150);await evaluate(`document.getElementById('documentation-api-btn').focus()`);
 assert.equal(await evaluate(`document.activeElement.id`),'documentation-api-btn','The export button accepts keyboard focus');
 win.webContents.sendInputEvent({type:'keyDown',keyCode:'Enter'});win.webContents.sendInputEvent({type:'char',keyCode:'\r'});win.webContents.sendInputEvent({type:'keyUp',keyCode:'Enter'});
 for(let i=0;i<100&&calls<2;i++)await pause(50);assert.equal(calls,2,'Native Enter opens the save picker');
 await until(`!document.getElementById('documentation-api-btn').disabled`);
 for(let i=0;i<100&&!fs.existsSync(path.join(profile,'API.md'));i++)await pause(50);
 const markdown=fs.readFileSync(path.join(profile,'API.md'),'utf8');assert.match(markdown,/^# Repetierer – API-Dokumentation/);assert.match(markdown,/\/v1\/snapshot/);assert.match(markdown,/\/teacher\/mail\/send/);assert.match(markdown,/\/teacher\/rounds/);
 assert.doesNotMatch(markdown,/\.cloudflare-private|\brt_[A-Za-z0-9_-]{43}\b/);
 mode='fail';await click('#documentation-api-btn');await until(`!document.getElementById('documentation-api-btn').disabled`);await until(`document.body.textContent.includes('konnte nicht gespeichert werden')`);
 await evaluate(`notify.clear()`);
 for(const size of [[1100,800],[1440,900]]){win.setSize(...size);await pause(200);assert.equal(await evaluate(`(()=>{const b=document.getElementById('documentation-api-btn').getBoundingClientRect(),d=document.getElementById('drawer').getBoundingClientRect();return b.left>=d.left&&b.right<=d.right&&b.top>=d.top&&b.bottom<d.bottom;})()`),true);await shot('documentation-drawer-'+size[0]);}
 await click('#drawer-back-btn');await settle();assert.equal(await evaluate(`document.getElementById('drawer-main-view').checkVisibility()`),true);
 await click('#drawer-data-view-btn');await settle();assert.equal(await evaluate(`document.getElementById('drawer-data-view').checkVisibility()`),true);assert.equal(await evaluate(`document.getElementById('drawer-documentation-view').checkVisibility()`),false);
 await click('#drawer-back-btn');await settle();await win.webContents.debugger.attach('1.3');await win.webContents.debugger.sendCommand('Emulation.setEmulatedMedia',{features:[{name:'prefers-reduced-motion',value:'reduce'}]});
 await click('#drawer-documentation-view-btn');await settle();await click('#drawer-back-btn');await settle();assert.equal(await evaluate(`document.getElementById('drawer-title').textContent`),'Navigation');
 assert.deepEqual(errors,[]);app.getAppPath=root;
 console.log('DOCUMENTATION_UI_PASS: native menu input and keyboard export, actual IPC/Markdown write, unavailable manual, cancel, failure, duplicate guard, back/file submenu, desktop sizes and reduced motion; temporary profile');win.destroy();app.quit();
}).catch(error=>{console.error(error.stack);win?.destroy();app.exit(1);});
