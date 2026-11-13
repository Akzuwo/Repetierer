const {app,BrowserWindow,net,safeStorage,clipboard}=require('electron');
if(!app){const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;const run=require('node:child_process').spawnSync(require('electron'),[__filename],{env,stdio:'inherit',windowsHide:true});process.exit(run.status??1);}
const fs=require('node:fs'),os=require('node:os'),path=require('node:path'),assert=require('node:assert/strict'),crypto=require('node:crypto'),Excel=require('exceljs');
const profile=fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-assessment-workflow-'));app.setPath('userData',profile);
const online=new Map();let win,failReopen=false,holdSync=false,releaseSync;const errors=[];
// The HTTPS transport is isolated; production IPC, Excel and AssessmentStore run unchanged.
net.fetch=async(url,options={})=>{
 const resource=new URL(url).pathname.split('/teacher/')[1],method=options.method||'GET';
 const ok=data=>new Response(JSON.stringify({success:true,...data}));
 if(resource==='me')return ok({teacherId:'fixture'});
 const match=resource?.match(/^rounds\/([^/]+)(?:\/(responses|close|reopen))?$/);assert.ok(match,'Unexpected API request');
 const [,id,action]=match,body=options.body?JSON.parse(options.body):{};
 if(method==='PUT'){online.set(id,{...body,id,closed:false,responses:[]});return ok({roundId:id});}
 const round=online.get(id);assert.ok(round);
 if(action==='responses'){if(holdSync)await new Promise(resolve=>releaseSync=resolve);return ok({responses:round.responses,closed:round.closed,nextCursor:null});}
 if(action==='close'){round.closed=true;return ok({closed:true});}
 if(action==='reopen'){
  if(failReopen)return new Response(JSON.stringify({success:false,error:'not_found'}),{status:404});
  round.closed=false;round.deadline=body.deadline;return ok({closed:false});
 }
 throw new Error('Unexpected API operation');
};
const pause=ms=>new Promise(resolve=>setTimeout(resolve,ms)),evaluate=code=>win.webContents.executeJavaScript(code,true);
async function until(code){for(let i=0;i<220;i++){if(await evaluate(code))return;await pause(35);}console.error(errors);throw new Error('Timeout: '+code);}
async function settle(){await until(`!document.getElementById('content').dataset.pageMotion && document.getElementById('assessment-body').getAttribute('aria-busy')!=='true'`);await pause(250);}
async function click(selector){await evaluate(`document.querySelector(${JSON.stringify(selector)}).scrollIntoView({block:'center'})`);await pause(180);const point=await evaluate(`(()=>{const r=document.querySelector(${JSON.stringify(selector)}).getBoundingClientRect();return {x:Math.round(r.x+r.width/2),y:Math.round(r.y+r.height/2)}})()`);win.webContents.sendInputEvent({type:'mouseDown',button:'left',clickCount:1,...point});win.webContents.sendInputEvent({type:'mouseUp',button:'left',clickCount:1,...point});}
async function chooseClass(index){await click('#assessment-class-select + .liquid-select-trigger');await until(`!!document.querySelector('.liquid-select-menu:popover-open')`);await click(`.liquid-select-menu:popover-open [data-index="${index}"]`);await settle();}
async function shot(name){win.webContents.invalidate();await pause(150);const out=path.join(__dirname,'../.ui-polish-smoke');fs.mkdirSync(out,{recursive:true});fs.writeFileSync(path.join(out,name+'.png'),(await win.webContents.capturePage()).toPNG());}
async function main(){
 await app.whenReady();require('../src/mainProcess');
 const storage=require('../src/storage'),book=new Excel.Workbook(),file=path.join(profile,'fixture.xlsx');
 for(const name of ['Klasse A','Klasse B']){const sheet=book.addWorksheet(name);sheet.getCell('A1').value='repetierer';sheet.getCell('A6').value='Anna Beispiel';sheet.getCell('A7').value='Max Beispiel';sheet.getCell('H6').value=1;sheet.getCell('H7').value=1;require('../src/jokerDates').migrateWorksheet(sheet);}
 await book.xlsx.writeFile(file);const workbookBytes=fs.readFileSync(file);storage.saveExcelFilePath(file);
 const token=require('../src/assessment/teacherCredentials').protectTeacherToken('rt_'+'a'.repeat(43),safeStorage);
 fs.writeFileSync(storage.getPaths().integrationEnvPath,`ASSESSMENT_API_URL=https://assessment-ui.invalid/app/services/repetierer/api\nASSESSMENT_TEACHER_TOKEN=${token}\n`);
 win=new BrowserWindow({width:1100,height:800,show:true,frame:false,webPreferences:{nodeIntegration:true,contextIsolation:false,offscreen:true,backgroundThrottling:false}});
 win.webContents.on('console-message',e=>{if(e.level==='error')errors.push(e.message);});
 await win.loadFile(path.join(__dirname,'../public/index.html'));await until(`document.querySelectorAll('#class-list button').length===2`);await pause(650);
 await click('#drawer-toggle');await pause(650);await click('#assessment-nav-btn');await until(`!document.getElementById('assessment-view').classList.contains('update-hidden')`);await settle();
 await chooseClass(1);await until(`!!document.getElementById('assessment-create-round-btn')`);
 await evaluate(`document.getElementById('assessment-round-title').value='Runde A'`);await click('#assessment-create-round-btn');await until(`!!document.querySelector('[data-round-id]')`);await settle();
 assert.equal(online.size,1);const remote=[...online.values()][0];
 await click('[data-round-id]');await until(`!!document.getElementById('assessment-invitation-btn')`);await settle();
 await click('#assessment-invitation-btn');await until(`notify.getActive().some(n=>n.text==='Einladungslink kopiert.')`);const link=`https://assessment-ui.invalid/app/services/repetierer/?round=${remote.id}`;
 assert.equal(await clipboard.readText(),link,'Clipboard contains only the invitation URL');
 const logCopy=await evaluate(`require('electron').ipcRenderer.invoke('copy-log-text','Fixture log')`);assert.equal(logCopy.copied,true);assert.equal(await clipboard.readText(),'Fixture log','Log copying also uses the supported main-process clipboard');
assert.equal(await evaluate(`!!document.querySelector('[data-copy],.assessment-invitation textarea')`),false);
 await evaluate(`notify.clear()`);await click('[data-teacher-index="1"]');await until(`!!document.querySelector('.teacher-assessment')`);await settle();
 assert.match(await evaluate(`document.getElementById('assessment-heading').textContent`),/Max Beispiel/);
 await click('#teacher-previous-btn');await settle();assert.match(await evaluate(`document.getElementById('assessment-heading').textContent`),/Anna Beispiel/);
 await evaluate(`const first=document.querySelector('[data-teacher-criterion]');first.scrollIntoView({block:'center'});first.focus()`);win.webContents.sendInputEvent({type:'keyDown',keyCode:'End'});win.webContents.sendInputEvent({type:'keyUp',keyCode:'End'});
 await until(`document.querySelector('[data-teacher-criterion]').dataset.rated==='true'`);
 await click('.teacher-criterion-comment summary');await evaluate(`document.querySelector('[data-teacher-comment]').focus()`);await win.webContents.insertText('Entwurf bleibt erhalten');
 await click('#assessment-back-btn');await settle();await click('[data-teacher-index="0"]');await settle();
 assert.equal(await evaluate(`document.querySelector('[data-teacher-criterion]').value`),'5');assert.equal(await evaluate(`document.querySelector('[data-teacher-comment]').value`),'Entwurf bleibt erhalten','Back/overview navigation preserves drafts');
 const count=await evaluate(`document.querySelectorAll('[data-teacher-criterion]').length`);
 for(let i=1;i<count;i++){await evaluate(`(()=>{const e=document.querySelectorAll('[data-teacher-criterion]')[${i}];e.scrollIntoView({block:'center'});e.focus()})()`);await pause(60);win.webContents.sendInputEvent({type:'keyDown',keyCode:'End'});win.webContents.sendInputEvent({type:'keyUp',keyCode:'End'});}
 await until(`[...document.querySelectorAll('[data-teacher-criterion]')].every(e=>e.dataset.rated==='true')`);
 await click('#teacher-save-btn');await settle();await until(`notify.getActive().some(n=>n.text.includes('gespeichert'))`);await click('#assessment-back-btn');await settle();
 const persisted=()=>Object.values(Object.values(JSON.parse(fs.readFileSync(storage.getPaths().assessmentPath,'utf8')).workbooks)[0].classes).find(c=>c.name==='Klasse A').rounds[0];
 assert.equal(Object.values(persisted().assessments).filter(a=>a.teacherAnswers).length,1);
 const response=name=>({responseId:crypto.randomUUID(),roundId:remote.id,timestamp:new Date().toISOString(),name,email:name.startsWith('Anna')?'anna@example.ch':'max@example.ch',answers:Object.fromEntries(remote.criteria.map(c=>[c.id,4]))});
 remote.responses.push(response('Anna Beispiel'));await evaluate(`notify.clear()`);await click('#assessment-sync-btn');await settle();assert.equal(persisted().externalResponses.length,1);
 await click('#assessment-close-round-btn');await until(`!!document.getElementById('assessment-reopen-round-btn')`);await settle();assert.equal(remote.closed,true);
 failReopen=true;await click('#assessment-reopen-round-btn');await until(`notify.getActive().some(n=>n.text.includes('unterstützt'))`);assert.equal(await evaluate(`document.getElementById('assessment-body').inert`),false,'API failures leave the page usable');assert.equal(remote.closed,true);
 failReopen=false;await evaluate(`notify.clear()`);await click('#assessment-reopen-round-btn');await until(`!!document.getElementById('assessment-close-round-btn')`);await settle();assert.equal(remote.closed,false);assert.equal(persisted().externalResponses.length,1);
 remote.responses.push(response('Max Beispiel'));await click('#assessment-sync-btn');await settle();assert.equal(persisted().externalResponses.length,2);
 await evaluate(`notify.clear()`);await click('#assessment-invitation-btn');await until(`notify.getActive().some(n=>n.text==='Einladungslink kopiert.')`);assert.equal(await clipboard.readText(),link,'Reopening keeps the same invitation URL');
 holdSync=true;await click('#assessment-sync-btn');await until(`document.getElementById('assessment-body').inert`);
 await click('#drawer-toggle');await pause(650);await click('#repetition-nav-btn');await until(`document.getElementById('assessment-view').classList.contains('update-hidden') && !document.getElementById('content').dataset.pageMotion`);
 assert.equal(typeof releaseSync,'function');holdSync=false;releaseSync();await pause(600);
 assert.equal(await evaluate(`document.getElementById('session-empty-view').checkVisibility()`),true,'A late sync result cannot replace the page after navigation');
 await click('#drawer-toggle');await pause(650);await click('#assessment-nav-btn');await until(`!!document.querySelector('[data-round-id]') && !document.getElementById('assessment-view').classList.contains('update-hidden')`);await settle();
 await click('[data-round-id]');await settle();
 await evaluate(`notify.clear()`);
 for(const size of [[1100,800],[1440,900]]){win.setSize(...size);await pause(250);await evaluate(`document.getElementById('assessment-body').scrollTop=0`);await pause(180);assert.equal(await evaluate(`[...document.querySelectorAll('.assessment-metric')].every(e=>e.textContent.trim().length>0&&getComputedStyle(e).opacity!=='0')`),true,'Metrics remain readable');await shot('assessment-workflow-'+size[0]);}
 await chooseClass(2);await until(`!!document.getElementById('assessment-create-round-btn')`);assert.equal(await evaluate(`document.querySelectorAll('[data-round-id]').length`),0,'Rounds do not leak into another class');
 await chooseClass(1);await click('[data-round-id]');await settle();assert.equal(await evaluate(`document.querySelectorAll('[data-teacher-index]').length`),2);
 assert.deepEqual(fs.readFileSync(file),workbookBytes,'Assessments do not modify Excel');assert.deepEqual(errors,[]);
 console.log('ASSESSMENT_WORKFLOW_UI_PASS: production IPC/store, native class selection/create/open, URL-only clipboard, direct student navigation, draft/save, sync, close/reopen/error recovery, late response and desktop metrics');
}
const watchdog=setTimeout(()=>{console.error('Assessment workflow timeout');app.exit(1);},150000);
main().then(()=>{clearTimeout(watchdog);app.exit(0);}).catch(error=>{console.error(error);clearTimeout(watchdog);app.exit(1);});
