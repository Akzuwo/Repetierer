const {app,BrowserWindow,ipcMain}=require('electron');
if(!app){const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;const r=require('node:child_process').spawnSync(require('electron'),[__filename],{env,stdio:'inherit',windowsHide:true});process.exit(r.status??1);}
const fs=require('node:fs'),path=require('node:path'),os=require('node:os'),assert=require('node:assert/strict');
const profile=fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-integration-ui-'));app.setPath('userData',profile);
let setups=0,creates=0,reopens=0,win;const errors=[];
// No installer, Ollama process, real mail or external API is used by this test.
require.cache[require.resolve('../src/ai/ollamaService')]={id:require.resolve('../src/ai/ollamaService'),filename:require.resolve('../src/ai/ollamaService'),loaded:true,exports:{createOllamaService:()=>({status:async()=>({state:'ready',message:'Bereit'}),setup:async()=>{setups++;return {state:'ready',message:'Bereit'};},chat:async()=>({content:'Fixture'})})}};
const pause=ms=>new Promise(r=>setTimeout(r,ms));
const evaluate=code=>win.webContents.executeJavaScript(code,true);
async function until(code){for(let i=0;i<180;i++){if(await evaluate(code))return;await pause(35);}throw new Error('Timeout: '+code);}
async function click(selector){const point=await evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();return {x:Math.round(r.x+r.width/2),y:Math.round(r.y+r.height/2)};})()`);win.webContents.sendInputEvent({type:'mouseDown',button:'left',clickCount:1,...point});win.webContents.sendInputEvent({type:'mouseUp',button:'left',clickCount:1,...point});}
async function shot(name){console.log('UI capture:',name);await pause(500);win.webContents.invalidate();await pause(100);const dir=path.join(__dirname,'../.ui-polish-smoke');fs.mkdirSync(dir,{recursive:true});fs.writeFileSync(path.join(dir,name+'.png'),(await win.webContents.capturePage()).toPNG());}
async function main(){
 await app.whenReady();require('../src/mainProcess');
 const {AssessmentStore}=require('../src/assessment/assessmentStore'),{presentClass,summarizeRound}=require('../src/assessment/mainAssessment');
 const store=new AssessmentStore(path.join(profile,'fixture-assessments.json')),workbook=path.join(profile,'fixture.xlsx'),people=[{id:6,name:'Anna Beispiel'},{id:7,name:'Max Beispiel'}];
 const round=store.createRound(workbook,'Demo A',people,{title:'Test-Runde'});store.setCloudStatus(workbook,'Demo A',people,round.id,'closed');
 const view=()=>{const {classRecord,round:current}=store.getRound(workbook,'Demo A',people,round.id);return {filePath:workbook,classRecord:presentClass(classRecord),round:current,students:classRecord.students,summary:summarizeRound(current,classRecord.students),assessmentConfigured:true,invitation:null,deliveryValidation:{valid:false,incomplete:people}};};
 const handlers={
  'assessment:get-overview':()=>presentClass(store.getClassContext(workbook,'Demo A',people)),
  'assessment:get-round':()=>view(),
  'assessment:create-round':()=>{creates++;return presentClass(store.getClassContext(workbook,'Demo A',people));},
  'assessment:reopen-round':args=>{reopens++;assert.ok(args.deadline);store.setDeadline(workbook,'Demo A',people,round.id,args.deadline);store.setCloudStatus(workbook,'Demo A',people,round.id,'open');return view();}
 };
 for(const [channel,handler] of Object.entries(handlers)){ipcMain.removeHandler(channel);ipcMain.handle(channel,(_e,args)=>({ok:true,data:handler(args)}));}
 win=new BrowserWindow({show:true,width:1100,height:800,frame:false,webPreferences:{nodeIntegration:true,contextIsolation:false,offscreen:true,backgroundThrottling:false}});
 win.webContents.on('console-message',e=>{if(e.level==='error')errors.push(e.message);});
 await win.loadFile(path.join(__dirname,'../public/index.html'));
 console.log('UI: loaded');
 await click('#settings-btn');await until(`document.getElementById('ai-setup-btn').textContent==='Assistent öffnen' && !document.getElementById('settings-view').classList.contains('update-hidden')`);await pause(650);
 assert.equal(await evaluate(`document.getElementById('ai-info-btn').hasAttribute('data-glass')`),false,'Help is a bare question mark, without a shader pill');
 assert.equal(await evaluate(`getComputedStyle(document.getElementById('ai-info-btn')).boxShadow`),'none');
 assert.equal(await evaluate(`getComputedStyle(document.getElementById('ai-info-btn')).backgroundColor`),'rgba(0, 0, 0, 0)');
 assert.match(await evaluate(`document.getElementById('ai-info-btn').title`),/Erklärung/);
 const helpPoint=await evaluate(`(()=>{const e=document.getElementById('ai-info-btn');e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();return {x:Math.round(r.x+r.width/2),y:Math.round(r.y+r.height/2)};})()`);
 win.webContents.sendInputEvent({type:'mouseMove',...helpPoint});await pause(250);
 assert.notEqual(await evaluate(`getComputedStyle(document.getElementById('ai-info-btn')).textShadow`),'none','Bare question mark has a hover glow');
 await shot('ai-help-bare-question-mark');
 await click('#ai-info-btn');await until(`document.getElementById('ai-info-modal').classList.contains('modal-open')`);
 assert.equal(await evaluate(`document.getElementById('ai-info-confirm-btn').checkVisibility()`),false);assert.match(await evaluate(`document.getElementById('ai-info-modal').textContent`),/ausdrücklichen Bestätigung/);await shot('ai-integration-explanation');
 await click('#ai-info-close-btn');await until(`getComputedStyle(document.getElementById('ai-info-modal')).display==='none'`);
 await click('#ai-remove-btn');await until(`document.getElementById('ai-remove-modal').classList.contains('modal-open')`);await pause(500);await click('#ai-remove-confirm-btn');
 await until(`document.getElementById('ai-setup-status').textContent==='KI-Integration entfernt' && getComputedStyle(document.getElementById('ai-remove-modal')).display==='none'`);
 assert.equal(require('../src/storage').getAppSettings().aiIntegrationEnabled,false);
 console.log('UI: AI removed');
 require('../src/storage').saveAppSettings(require('../src/storage').getAppSettings());assert.equal(require('../src/storage').getAppSettings().aiIntegrationEnabled,false);
 const loaded=new Promise(resolve=>win.webContents.once('did-finish-load',resolve));win.reload();await loaded;await until(`document.getElementById('ai-setup-status').textContent==='KI-Integration entfernt'`);
 await click('#settings-btn');await until(`!document.getElementById('settings-view').classList.contains('update-hidden')`);await pause(650);
 await click('#ai-setup-btn');await until(`document.getElementById('ai-info-confirm-btn').checkVisibility()`);assert.equal(setups,0);
 await pause(500);await click('#ai-info-confirm-btn');await until(`document.getElementById('ai-setup-btn').textContent==='Assistent öffnen'`);assert.equal(setups,1);assert.equal(require('../src/storage').getAppSettings().aiIntegrationEnabled,true);
 console.log('UI: AI re-enabled after consent');
 await click('#settings-back-btn');await pause(650);win.webContents.send('classes',['Demo A'],workbook);await pause(200);
 await evaluate(`document.getElementById('assessment-nav-btn').click()`);await until(`!document.getElementById('assessment-view').classList.contains('update-hidden')`);await pause(650);
 await evaluate(`const select=document.getElementById('assessment-class-select');select.value='Demo A';select.dispatchEvent(new Event('change'));`);await until(`!!document.getElementById('assessment-create-round-btn')`);await pause(650);
 await evaluate(`document.getElementById('assessment-round-title').value='Ungültige Runde';document.getElementById('assessment-round-deadline').value='2099-01-01';document.getElementById('assessment-create-round-btn').click()`);
 await until(`notify.getActive().some(n=>n.text.includes('90 Tage'))`);assert.equal(creates,0);
 await evaluate(`document.querySelector('[data-round-id]').click()`);await until(`!!document.getElementById('assessment-reopen-round-btn')`);await pause(650);await evaluate(`notify.clear()`);await shot('assessment-reopen-round');
 await click('#assessment-reopen-round-btn');await until(`!!document.getElementById('assessment-close-round-btn')`);assert.equal(reopens,1);
 await evaluate(`document.getElementById('assessment-teacher-btn').click()`);await until(`!!document.getElementById('teacher-comment')`);await pause(650);
 await evaluate(`document.getElementById('teacher-comment').value='Ungespeicherter Entwurf';document.getElementById('teacher-comment').focus()`);
 win.webContents.send('assessment:auto-synced',{filePath:workbook,className:'Demo A',roundId:round.id,view:view()});await pause(200);
 assert.equal(await evaluate(`document.getElementById('teacher-comment').value`),'Ungespeicherter Entwurf');
 assert.equal(await evaluate(`document.activeElement.id`),'teacher-comment');
 assert.deepEqual(errors,[]);console.log('INTEGRATION_LIFECYCLE_UI_PASS: native KI help/remove/reinstall, persisted opt-out, deadline validation, reopen, unsaved teacher draft preserved');
}
const timer=setTimeout(()=>{console.error('Integration UI timeout');app.exit(1);},60000);
main().then(()=>{clearTimeout(timer);app.exit(0);}).catch(e=>{console.error(e);clearTimeout(timer);app.exit(1);});
