const {app,BrowserWindow}=require('electron');
if(!app){const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;const result=require('node:child_process').spawnSync(require('electron'),[__filename,...process.argv.slice(2)],{env,stdio:'inherit',windowsHide:true});process.exit(result.status??1);}
const fs=require('node:fs'),os=require('node:os'),path=require('node:path'),assert=require('node:assert/strict'),Excel=require('exceljs');
const profile=fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-homework-ui-'));app.setPath('userData',profile);
let win;const errors=[],pause=ms=>new Promise(r=>setTimeout(r,ms));
const evaluate=code=>win.webContents.executeJavaScript(code,true);
async function until(code){for(let i=0;i<180;i++){if(await evaluate(code))return;await pause(40);}throw new Error('Timeout: '+code);}
async function click(selector){
 await evaluate(`document.querySelector(${JSON.stringify(selector)}).scrollIntoView({block:'center'})`);await pause(220);
 const point=await evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});const r=e.getBoundingClientRect();return {x:Math.round(r.x+r.width/2),y:Math.round(r.y+r.height/2)};})()`);

 win.webContents.sendInputEvent({type:'mouseDown',button:'left',clickCount:1,...point});win.webContents.sendInputEvent({type:'mouseUp',button:'left',clickCount:1,...point});
}
async function settle(){await until(`!document.getElementById('content').dataset.pageMotion`);await pause(450);}
async function shot(name){const out=path.join(__dirname,'../.ui-polish-smoke');fs.mkdirSync(out,{recursive:true});fs.writeFileSync(path.join(out,name+'.png'),(await win.webContents.capturePage()).toPNG());}
async function main(){
 await app.whenReady();require('../src/mainProcess');
 const storage=require('../src/storage'),program=require('../src/program'),homework=require('../src/homework');
 const file=path.join(profile,'fixture.xlsx'),book=new Excel.Workbook();
 for(const name of ['Klasse A','Klasse B']){const s=book.addWorksheet(name);s.getCell('A1').value='repetierer';for(let i=0;i<25;i++){s.getCell(i+6,1).value='Beispiel '+String(i+1).padStart(2,'0');s.getCell(i+6,8).value=1;}require('../src/jokerDates').migrateWorksheet(s);}
 await book.xlsx.writeFile(file);storage.saveExcelFilePath(file);const original=fs.readFileSync(file);
 win=new BrowserWindow({show:true,width:1100,height:800,frame:false,webPreferences:{nodeIntegration:true,contextIsolation:false,offscreen:true,backgroundThrottling:false}});
 win.webContents.on('console-message',event=>{if(event.level==='error')errors.push(event.message);});
 await win.loadFile(path.join(__dirname,'../public/index.html'));
 await until(`document.querySelectorAll('#class-list button').length===2`);await settle();
 // A loaded/selected workbook class is not an active session.
 await new Promise(r=>program.setClass('Klasse B',r));
 await click('#settings-btn');await until(`!document.getElementById('settings-view').classList.contains('update-hidden')`);await settle();
 assert.equal(await evaluate(`document.getElementById('settings-class-select').value`),'');
 assert.equal(await evaluate(`document.getElementById('settings-class-select').disabled`),false);
 await click('#settings-class-tab');await pause(500);
 assert.equal(await evaluate(`document.getElementById('settings-class-tab').getAttribute('aria-selected')`),'true');
 assert.equal(await evaluate(`getComputedStyle(document.getElementById('settings-navdock')).webkitAppRegion`),'no-drag','Dock must accept native clicks instead of dragging the window');
 assert.equal(await evaluate(`document.getElementById('settings-class-rules').disabled`),true);
 assert.equal(await evaluate(`document.getElementById('settings-class-picker').checkVisibility()`),true);
 for(const size of [[1100,800],[1440,900]]){
  win.setSize(...size);await pause(300);
  const dock=await evaluate(`(()=>{const d=document.getElementById('settings-navdock'),r=d.getBoundingClientRect(),b=document.getElementById('save-settings-btn').getBoundingClientRect();return {parent:d.parentElement.id,glass:d.dataset.glass,marker:!!d.querySelector('.nav-selection'),top:r.top,bottom:r.bottom,center:r.x+r.width/2,width:innerWidth,height:innerHeight,saveBottom:b.bottom};})()`);
  assert.equal(dock.parent,'container');assert.ok(dock.glass && dock.marker);assert.ok(dock.top>dock.height*.8 && dock.bottom<=dock.height && Math.abs(dock.center-dock.width/2)<2);assert.ok(dock.saveBottom<dock.top,'Save actions do not overlap the floating dock');
  await shot('settings-dock-empty-'+size[0]);
 }
 win.setSize(1100,800);await pause(200);
 await click('#settings-back-btn');await until(`document.getElementById('settings-view').classList.contains('update-hidden')`);await settle();
 if(!process.argv.includes('--settings-only')) {await click('#class-list button');await until(`!document.getElementById('session-attendance-view').classList.contains('update-hidden')`);await settle();
 await click('#session-attendance-list button:last-child');await click('#confirm-session-start-btn');
 await until(`!document.getElementById('participation-view').classList.contains('update-hidden')`);await settle();
 assert.deepEqual(await evaluate(`[...document.querySelectorAll('#session-chrome .session-tabs button')].map(e=>e.textContent)`),['Mitarbeit','Hausaufgaben','Repetition']);
 await click('#session-homework-tab');await until(`document.getElementById('homework-migration-modal').classList.contains('modal-open')`);await pause(500);
 assert.deepEqual(fs.readFileSync(file),original,'Preview does not modify workbook');await shot('homework-migration-preview');
 await click('#homework-migration-run');await until(`!document.getElementById('homework-migration-modal').classList.contains('modal-open') && document.querySelectorAll('#homework-student-list button').length===24`);await settle();
 assert.equal(program.getHomeworkStatus().needsMigration,false);await shot('homework-students');
 const before=await evaluate(`(()=>{const r=document.querySelector('#homework-student-list button:last-child').getBoundingClientRect();return r.y;})()`);
 await evaluate(`document.getElementById('homework-student-list').scrollTop=9999`);await pause(500);
 assert.ok((await evaluate(`document.getElementById('homework-student-list').scrollTop`))>0,'Student list scrolls');
 assert.ok(before>await evaluate(`document.querySelector('#homework-student-list button:last-child').getBoundingClientRect().y`));
 await evaluate(`document.getElementById('homework-student-list').scrollTop=0`);await pause(500);
 for(const grade of ['5.5','6']){
  await click('#homework-student-list button');await until(`!!document.getElementById('homework-grade-input')`);await settle();
  assert.equal(await evaluate(`document.getElementById('homework-grade-input').hasAttribute('data-glass')`),false);
  await evaluate(`document.getElementById('homework-grade-input').value=${JSON.stringify(grade)}`);await shot('homework-grade');
  await click('#homework-save-btn');await until(`!!document.getElementById('homework-student-list')`);await settle();
 }
 assert.equal(program.getHomeworkPersons()[0].homework.length,2);assert.equal(program.getHomeworkPersons()[0].joker,2);
 await click('#homework-undo-btn');await until(`document.querySelector('#homework-student-list button span').textContent==='1 Vorführung'`);assert.equal(program.getHomeworkPersons()[0].joker,1);
 await click('#homework-redo-btn');await until(`document.querySelector('#homework-student-list button span').textContent==='2 Vorführungen'`);assert.equal(program.getHomeworkPersons()[0].joker,2);
 }
 await evaluate(`notify.clear()`);await click('#settings-btn');await until(`!document.getElementById('settings-view').classList.contains('update-hidden')`);await settle();
 await click('#settings-global-tab');await pause(500);assert.equal(await evaluate(`document.getElementById('settings-global-fields').classList.contains('update-hidden')`),false);await shot('settings-global-scope');
 await click('#settings-class-tab');await pause(500);assert.equal(await evaluate(`document.getElementById('settings-class-select').value`),process.argv.includes('--settings-only')?'':'Klasse A');await shot('settings-class-scope');
 await evaluate(`document.getElementById('settings-class-select').value='Klasse B';document.getElementById('settings-class-select').dispatchEvent(new Event('change',{bubbles:true}))`);
 await until(`window.getSettingsScopeContext().className==='Klasse B'`);
 for(const id of ['homework-joker-setting','homework-weight-setting']){await evaluate(`document.getElementById('${id}').scrollIntoView({block:'center'});document.getElementById('${id}').focus()`);await pause(150);win.webContents.sendInputEvent({type:'keyDown',keyCode:'Right'});win.webContents.sendInputEvent({type:'keyUp',keyCode:'Right'});}
 await pause(400);await click('#homework-enabled-setting');await until(`document.getElementById('homework-enabled-setting').checked===false`);assert.equal(await evaluate(`document.getElementById('homework-enabled-setting').checked`),false);
 assert.equal(await evaluate(`document.getElementById('settings-class-select').disabled`),true,'Class changes cannot discard a draft');
 await click('#save-settings-btn');await until(`document.getElementById('settings-view').classList.contains('update-hidden')`);await settle();
 assert.equal(storage.getClassSettings(file,'Klasse B').homeworkJokerThreshold,3);assert.equal(storage.getClassSettings(file,'Klasse B').homeworkWeight,1.3);
 assert.equal(storage.getClassSettings(file,'Klasse B').homeworkEnabled,false);assert.equal(storage.getClassSettings(file,'Klasse A').homeworkEnabled,true);
 if(!process.argv.includes('--settings-only'))assert.equal(await evaluate(`document.getElementById('session-homework-tab').classList.contains('update-hidden')`),false,'Editing another class does not change the current session');
 if(process.argv.includes('--settings-only')){console.log('SETTINGS_ONLY_PASS');return;}
 await click('#end-session-btn');await pause(450);await click('#confirm-session-end-btn');await until(`!document.getElementById('session-empty-view').classList.contains('update-hidden')`);await settle();
 await click('#class-list button:nth-child(2)');await until(`!document.getElementById('session-attendance-view').classList.contains('update-hidden')`);await settle();await click('#confirm-session-start-btn');await until(`!document.getElementById('participation-view').classList.contains('update-hidden')`);await settle();
 assert.equal(await evaluate(`document.getElementById('session-homework-tab').classList.contains('update-hidden')`),true,'New class applies its own setting');
 const saved=new Excel.Workbook();await saved.xlsx.readFile(file);assert.equal(homework.layout(saved.getWorksheet('Klasse B')).slots.length,0,'Disabled class stays untouched');
 const active=require('../src/teachingSessions').createTeachingSessionStore(storage.getPaths().teachingSessionsPath).getActiveSession();assert.equal(active.filePath,file);
 const otherFile=path.join(profile,'other.xlsx');fs.copyFileSync(file,otherFile);const otherBytes=fs.readFileSync(otherFile);await new Promise(r=>program.setFile(otherFile,r));await new Promise(r=>program.setClass('Klasse B',r));
 const wrongFile=await evaluate(`require('electron').ipcRenderer.invoke('homework:save',{sessionId:${JSON.stringify(active.id)},personId:0,grade:5})`);assert.equal(wrongFile.success,false);assert.match(wrongFile.error,/anderen Exceldatei/);assert.deepEqual(fs.readFileSync(otherFile),otherBytes);
 assert.deepEqual(errors,[],'No renderer errors');console.log('HOMEWORK_UI_PASS: native session, migration/backup, scroll, two grades/joker, undo/redo, global/class settings and isolated class switching');
}
const watchdog=setTimeout(()=>{console.error('Homework UI timeout');app.exit(1);},150000);
main().then(()=>{clearTimeout(watchdog);app.exit(0);}).catch(error=>{console.error(error);clearTimeout(watchdog);app.exit(1);});
