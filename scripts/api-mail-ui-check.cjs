const {app,BrowserWindow,net,safeStorage,dialog}=require('electron');
if(!app){const env={...process.env,NODE_USE_SYSTEM_CA:'1'};delete env.ELECTRON_RUN_AS_NODE;const r=require('node:child_process').spawnSync(require('electron'),[__filename],{env,windowsHide:true,stdio:'inherit',timeout:90000});process.exit(r.status??1);}
const fs=require('node:fs'),path=require('node:path'),os=require('node:os'),assert=require('node:assert/strict');
const profile=fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-api-mail-ui-'));app.setPath('userData',profile);
const pause=ms=>new Promise(r=>setTimeout(r,ms));let win,mode='live',connectionChecks=0;
const originalFetch=net.fetch.bind(net);
const schedule=global.setInterval;let healthTick;
global.setInterval=(fn,ms,...args)=>{if(ms===300000)healthTick=fn;return schedule(fn,ms,...args);};
net.fetch=async(url,options)=>{
 if(String(url).endsWith('/teacher/me'))connectionChecks++;
 if(mode==='offline'&&String(url).includes('/teacher/'))throw new Error('synthetic unavailable');
 if(String(url).endsWith('/teacher/mail/capabilities') && mode!=='live'){
  if(mode==='offline')throw new Error('synthetic unavailable');
  return new Response(JSON.stringify({success:true,mail:{supported:true,configured:false,dailyLimit:200}}),{headers:{'Content-Type':'application/json'}});
 }
 return originalFetch(url,options);
};
async function evaluate(code){return win.webContents.executeJavaScript(code);}
async function until(code){for(let i=0;i<100;i++){if(await evaluate(code))return;await pause(80);}throw new Error('Timeout: '+code);}
async function click(selector){await evaluate(`document.querySelector(${JSON.stringify(selector)}).scrollIntoView({block:'center'})`);await pause(400);const p=await evaluate(`(()=>{const r=document.querySelector(${JSON.stringify(selector)}).getBoundingClientRect();return {x:Math.round(r.x+r.width/2),y:Math.round(r.y+r.height/2)}})()`);win.webContents.sendInputEvent({type:'mouseDown',button:'left',clickCount:1,...p});win.webContents.sendInputEvent({type:'mouseUp',button:'left',clickCount:1,...p});await pause(150);}
async function open(){await click('#settings-btn');await until(`!document.getElementById('content').dataset.pageMotion && document.getElementById('settings-view').checkVisibility()`);}
async function shot(name){const dir=path.join(__dirname,'../.ui-polish-smoke');fs.mkdirSync(dir,{recursive:true});fs.writeFileSync(path.join(dir,name+'.png'),(await win.webContents.capturePage()).toPNG());}
app.whenReady().then(async()=>{
 require('../src/mainProcess');
 const source='C:/Users/timow/Downloads/Repetierer-Konfiguration.repetierer',before=fs.readFileSync(source),values=JSON.parse(before).values;
 values.ASSESSMENT_API_URL='https://akzuwo.ch/app/services/repetierer/api';
 values.ASSESSMENT_TEACHER_TOKEN=require('../src/assessment/teacherCredentials').protectTeacherToken(values.ASSESSMENT_TEACHER_TOKEN,safeStorage);
 fs.writeFileSync(path.join(profile,'repetierer.env'),Object.entries(values).map(([k,v])=>k+'='+JSON.stringify(v)).join('\n'));
 win=new BrowserWindow({show:false,width:1100,height:800,webPreferences:{nodeIntegration:true,contextIsolation:false,offscreen:true}});
 await win.loadFile(path.join(__dirname,'../public/index.html'));
 await until(`document.getElementById('assessment-api-status').textContent.includes('Erreichbar')`);
 assert.ok(connectionChecks>0,'Connection is authenticated at app startup before opening settings');
 const initialChecks=connectionChecks;assert.equal(typeof healthTick,'function','Five-minute health check runs in the main process');
 await open();await until(`document.getElementById('brevo-status').textContent.includes('Cloudversand')`);
 assert.equal(connectionChecks,initialChecks,'Opening settings uses the shared cached health result');
 assert.equal(await evaluate(`document.querySelector('#assessment-api-test-btn,#brevo-test-btn')`),null);
 assert.equal(await evaluate(`document.querySelector('#save-settings-btn,#close-settings-modal')`),null);
 assert.equal(await evaluate(`document.getElementById('smtp-settings-actions').checkVisibility()`),false);
 assert.equal(await evaluate(`document.getElementById('smtp-mail-limit-setting').disabled`),true);
 assert.equal(await evaluate(`document.getElementById('smtp-mail-limit-row').checkVisibility()`),false);
 await click('#local-api-setting');assert.equal(await evaluate(`document.getElementById('local-api-setting').checked`),false);await shot('api-mail-local-api');
 await evaluate(`document.getElementById('brevo-status').scrollIntoView({block:'center'})`);await pause(500);await shot('api-mail-cloud');
 assert.equal(JSON.parse(fs.readFileSync(path.join(profile,'settings.json'))).localApiEnabled,false);
 assert.equal(await evaluate(`document.getElementById('local-api-status').textContent.startsWith('Deaktiviert')`),true);
 assert.equal(await evaluate(`document.getElementById('settings-view').checkVisibility()`),true,'Live save leaves settings open');
 await click('#settings-back-btn');await open();assert.equal(await evaluate(`document.getElementById('local-api-setting').checked`),false);
 await evaluate(`document.getElementById('assessment-config-remove-btn').scrollIntoView({block:'center'})`);await pause(400);
 const removal=await evaluate(`(()=>{const b=document.getElementById('assessment-config-remove-btn'),r=b.getBoundingClientRect(),p=b.parentElement.getBoundingClientRect();return {width:r.width,parent:p.width,last:b.parentElement===document.getElementById('settings-global-fields').lastElementChild};})()`);
 assert.equal(removal.last,true);assert.ok(Math.abs(removal.width-removal.parent)<2,'Removal button spans the full settings width');
 await shot('api-mail-cloud-bottom');
 const exported=path.join(profile,'cloud.repetierer');dialog.showSaveDialogSync=()=>exported;
 await click('#integration-package-export-btn');await until(`document.getElementById('integration-package-export-btn').disabled===false`);
 for(let i=0;i<50&&!fs.existsSync(exported);i++)await pause(100);
 assert.ok(fs.existsSync(exported));assert.equal(JSON.parse(fs.readFileSync(exported)).values.BREVO_SMTP_PASSWORD,undefined);
 assert.ok(fs.readFileSync(path.join(profile,'repetierer.env'),'utf8').includes('BREVO_SMTP_PASSWORD='));assert.deepEqual(fs.readFileSync(source),before);
 // A failed atomic settings write rolls the control back without reopening/closing the page.
 const rename=fs.renameSync;let failures=0;
 fs.renameSync=(from,to)=>{if(to===path.join(profile,'settings.json')){failures++;throw new Error('synthetic denied');}return rename(from,to);};
 await evaluate(`document.getElementById('notification-duration-setting').scrollIntoView({block:'center'});document.getElementById('notification-duration-setting').focus()`);await pause(300);
 win.webContents.sendInputEvent({type:'keyDown',keyCode:'Home'});win.webContents.sendInputEvent({type:'keyUp',keyCode:'Home'});
 await until(`document.getElementById('notification-duration-setting').value==='6'`);assert.equal(failures,1);fs.renameSync=rename;
 assert.equal(await evaluate(`document.getElementById('notification-duration-setting-value').value`),'6 Sek.');
 await click('#settings-back-btn');mode='smtp';await healthTick();await open();await until(`document.getElementById('brevo-status').textContent.includes('Lokaler SMTP')`);
 assert.equal(await evaluate(`document.getElementById('smtp-mail-limit-setting').disabled`),false);await click('#smtp-mail-limit-setting');await shot('api-mail-smtp');
 await until(`document.getElementById('brevo-status').textContent.includes(document.getElementById('smtp-mail-limit-setting').value)`);
 mode='offline';await healthTick();await until(`document.getElementById('brevo-status').textContent.includes('nicht erreichbar')`);
 assert.equal(await evaluate(`document.getElementById('smtp-mail-limit-setting').disabled`),true);await shot('api-mail-offline');
 mode='live';await healthTick();await until(`document.getElementById('brevo-status').textContent.includes('Cloudversand')`);
 assert.equal(await evaluate(`document.getElementById('smtp-settings-actions').checkVisibility()`),false);
 global.setInterval=schedule;
 console.log('API_MAIL_UI_PASS: authenticated startup check, cached settings reads, five-minute main-process health check, pushed offline/recovery status, no test buttons, hidden SMTP for cloud, live saves/failure rollback and cloud-only export; 1100x800 temporary profile');
 win.destroy();app.quit();
}).catch(error=>{console.error(error.stack);win?.destroy();app.exit(1);});
