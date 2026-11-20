const {app,BrowserWindow,session}=require('electron');
const fs=require('node:fs'),path=require('node:path'),os=require('node:os'),assert=require('node:assert/strict');
if(!app){
 for(const mode of ['grade','empty','direct']){
  const profile=fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-close-ui-'));
  const env={...process.env,REPETIERER_CLOSE_TEST_PROFILE:profile,REPETIERER_CLOSE_TEST_MODE:mode};delete env.ELECTRON_RUN_AS_NODE;
  const run=require('node:child_process').spawnSync(require('electron'),[__filename],{env,stdio:'inherit',windowsHide:true,timeout:150000});
  assert.equal(run.status,0,`App workflow ${mode}`);
  assert.equal(JSON.parse(fs.readFileSync(path.join(profile,'close-proof.json'),'utf8')).passed,true);
  console.log('SESSION_CLOSE_UI_PASS '+mode);
 }
 process.exit(0);
}
const profile=process.env.REPETIERER_CLOSE_TEST_PROFILE,mode=process.env.REPETIERER_CLOSE_TEST_MODE;
app.setPath('userData',profile);
const pause=ms=>new Promise(resolve=>setTimeout(resolve,ms));let win,store,program,paths;
const evaluate=async code=>{try{return await win.webContents.executeJavaScript(code,true);}catch(error){console.error('Renderer check failed:',code);throw error;}};
async function until(code){for(let i=0;i<350;i++){if(await evaluate(code))return;await pause(50);}throw new Error('Timeout: '+code);}
async function click(selector){
 const point=await evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});if(!e||!e.checkVisibility())throw new Error('Not visible: '+${JSON.stringify(selector)});const r=e.getBoundingClientRect();return {x:Math.round(r.x+r.width/2),y:Math.round(r.y+r.height/2)}})()`);
 win.webContents.sendInputEvent({type:'mouseDown',button:'left',clickCount:1,...point});win.webContents.sendInputEvent({type:'mouseUp',button:'left',clickCount:1,...point});await pause(90);
}
async function question(text){await until(`document.getElementById('question-modal').classList.contains('modal-open')&&document.getElementById('question-title').textContent===${JSON.stringify(text)}`);await pause(450);}
async function answer(index){await click('#question-button'+index);await pause(350);}
async function key(keyCode,modifiers=[]){win.webContents.sendInputEvent({type:'keyDown',keyCode,modifiers});if(keyCode==='Enter')win.webContents.sendInputEvent({type:'char',keyCode:'\r'});win.webContents.sendInputEvent({type:'keyUp',keyCode,modifiers});await pause(100);}
async function typeGrade(value){await click('#grade-input');await key('A',['control']);win.webContents.insertText(value);await pause(70);}
async function start(){await click('#class-list button');await until(`!document.getElementById('session-attendance-view').classList.contains('update-hidden')`);await pause(450);await click('#confirm-session-start-btn');await until(`!document.getElementById('participation-view').classList.contains('update-hidden')`);await pause(450);}
async function select(){await click('#session-repetition-tab');await pause(500);await click('#start-btn');await until(`!document.getElementById('grade-input').classList.contains('disabled')`);await pause(500);}
async function shot(name){const dir=path.join(__dirname,'../.ui-polish-smoke');fs.mkdirSync(dir,{recursive:true});fs.writeFileSync(path.join(dir,name+'.png'),(await win.webContents.capturePage()).toPNG());}
function exitProof(check){app.once('will-quit',()=>{check();fs.writeFileSync(path.join(profile,'close-proof.json'),JSON.stringify({passed:true,mode}));});}
async function main(){
 const Excel=require('exceljs'),book=new Excel.Workbook(),sheet=book.addWorksheet('Fixture Klasse');sheet.getCell('A1').value='repetierer';
 for(let i=0;i<2;i++){sheet.getCell(i+6,1).value=['Lea Beispiel','Ben Beispiel'][i];sheet.getCell(i+6,8).value=1;}
 require('../src/jokerDates').migrateWorksheet(sheet);
 const file=path.join(profile,'fixture.xlsx');await book.xlsx.writeFile(file);
 const storage=require('../src/storage');paths=storage.getPaths();storage.saveExcelFilePath(file);storage.saveAppSettings({visualEffectsEnabled:false});
 require('../app');await app.whenReady();
 session.defaultSession.webRequest.onBeforeRequest((details,callback)=>callback({cancel:/^https?:/.test(details.url)}));
 for(let i=0;i<300;i++){win=BrowserWindow.getAllWindows().find(w=>w.webContents.getURL().endsWith('/public/index.html'));if(win)break;await pause(50);}
 assert(win,'Main window');win.show();win.focus();
 await until(`document.querySelectorAll('#class-list button').length===1&&getComputedStyle(document.getElementById('container')).opacity==='1'`);await pause(700);
 program=require('../src/program');store=require('../src/teachingSessions').createTeachingSessionStore(paths.teachingSessionsPath);
 if(mode==='direct'){exitProof(()=>assert.equal(store.getActiveSession(),null));win.close();return;}
 await start();
 if(mode==='empty'){
  await click('#quit');await question('Session beenden und schliessen?');await answer(2);await question('Session beenden?');
  const id=store.getActiveSession().id;
  exitProof(()=>{assert.equal(store.getActiveSession(),null);assert(store.readData().sessions.find(s=>s.id===id).endedAt);});await answer(2);return;
 }
 const bytes=fs.readFileSync(paths.teachingSessionsPath);
 for(const size of [[1100,800],[1440,900]]){
  win.setSize(...size);await pause(180);await click('#end-session-btn');await question('Session beenden?');
  assert.match(await evaluate(`document.getElementById('question-content').textContent`),/keine Repetition/);
  assert.equal(await evaluate('document.activeElement.id'),'question-button1');
  assert.equal(await evaluate(`(()=>{const a=document.getElementById('question-button1').getBoundingClientRect(),b=document.getElementById('question-button2').getBoundingClientRect(),c=document.querySelector('#question-modal .modal-content').getBoundingClientRect();return Math.abs(a.width-b.width)<1&&c.left>=0&&c.right<=innerWidth&&c.top>=0&&c.bottom<=innerHeight})()`),true);
  await shot('balanced-session-empty-'+size[0]);await key('Tab');assert.equal(await evaluate('document.activeElement.id'),'question-button2');await key('Escape');await pause(350);
  assert.deepEqual(fs.readFileSync(paths.teachingSessionsPath),bytes);
 }
 win.setSize(1100,800);await pause(200);
 await evaluate(`window.questionResults=[];window.askQuestion({title:'<b>Frage eins</b>',content:'Nur Plaintext.',buttons:['Zurück','Weiter'],highlight:2,focus:1}).then(value=>questionResults.push(value));window.askQuestion({title:'Frage zwei',content:'Wartet auf die erste Entscheidung.',buttons:['Zurück','Weiter']}).then(value=>questionResults.push(value));undefined`);
 await question('<b>Frage eins</b>');assert.equal(await evaluate(`!!document.querySelector('#question-title b')`),false);
 assert.equal(await evaluate('document.activeElement.id'),'question-button1');assert.equal(await evaluate(`document.getElementById('question-button2').classList.contains('question-highlight')`),true);
 await answer(2);await question('Frage zwei');
 const doublePoint=await evaluate(`(()=>{const r=document.getElementById('question-button2').getBoundingClientRect();return {x:Math.round(r.x+r.width/2),y:Math.round(r.y+r.height/2)}})()`);
 win.webContents.sendInputEvent({type:'mouseDown',button:'left',clickCount:2,...doublePoint});win.webContents.sendInputEvent({type:'mouseUp',button:'left',clickCount:2,...doublePoint});await pause(300);
 assert.equal(await evaluate('questionResults.length'),1,'Double-click cannot accept the queued follow-up question');
 await key('Escape');await pause(350);assert.deepEqual(await evaluate('questionResults.map(value=>value.action)'),['button2','cancel']);
 console.log('QUESTIONS_UI_PASS queue, plain text, independent focus, double-click guard and cancellation');
 win.close();await question('Session beenden und schliessen?');await answer(2);await question('Session beenden?');await answer(1);assert.equal(win.isDestroyed(),false);
 await select();await typeGrade('5');
 const originalJoker=program.setJoker;let delayedJoker;
 program.setJoker=callback=>{delayedJoker=callback;};await click('#res > :nth-child(4)');
 assert.equal(require('../src/mainProcess').getSessionFinishState().saving,true);
 win.close();await pause(650);assert(store.getActiveSession());assert.equal(win.isDestroyed(),false);
 assert.equal(await evaluate(`document.getElementById('question-modal').classList.contains('modal-open')`),false);
 delayedJoker('fixture-cancel');program.setJoker=originalJoker;await pause(150);
 assert.equal(require('../src/mainProcess').getSessionFinishState().saving,false);
 const originalHomework=program.saveHomework;let delayedHomework;
 program.saveHomework=(_id,_grade,callback)=>{delayedHomework=callback;};
 const homeworkArgs={sessionId:store.getActiveSession().id,personId:store.getActiveSession().presentStudents[0].id,grade:5};
 await evaluate(`window.homeworkSavePromise=require('electron').ipcRenderer.invoke('homework:save',${JSON.stringify(homeworkArgs)});undefined`);
 for(let i=0;!delayedHomework&&i<100;i++)await pause(20);assert(delayedHomework);
 win.close();await pause(650);assert(store.getActiveSession());assert.equal(win.isDestroyed(),false);
 assert.equal(require('../src/mainProcess').getSessionFinishState().saving,true);
 delayedHomework(null);program.saveHomework=originalHomework;await evaluate('homeworkSavePromise');
 assert.equal(require('../src/mainProcess').getSessionFinishState().saving,false);
 console.log('SESSION_WRITE_GUARD_PASS delayed joker and homework block native quit; failed callbacks release guard');
 await click('#end-session-btn');await question('Noch eine Note speichern?');await shot('balanced-open-grade');await answer(1);
 assert.equal(await evaluate('document.activeElement.id'),'grade-input');
 const originalSave=program.saveGrade;program.saveGrade=(_grade,callback)=>callback('fixture-save-failure');
 await click('#res > :nth-child(3)');await pause(350);program.saveGrade=originalSave;
 assert(store.getActiveSession());assert.equal(program.getSelectedPerson().selectionToken!==null,true);assert.equal(await evaluate(`document.getElementById('grade-input').value`),'5');
 await click('#end-session-btn');await question('Noch eine Note speichern?');await answer(1);await click('#res > :nth-child(3)');
 await until(`document.getElementById('session-empty-view').classList.contains('update-hidden')===false`);assert.equal(store.getActiveSession(),null);
 assert.equal(store.readData().sessions[0].repetitionEntries[0].grade,5);
 assert(storage.getBackup().entries.some(e=>e.type==='grade'&&e.grade===5&&e.excelWriteSucceeded));
 await pause(500);await start();await select();await typeGrade('4.5');
 win.close();await question('Session beenden und schliessen?');await answer(2);await question('Noch eine Note speichern?');await answer(2);await question('Session beenden?');await answer(1);
 assert(store.getActiveSession());assert.equal(await evaluate(`document.getElementById('grade-input').value`),'4.5');
 win.close();await question('Session beenden und schliessen?');await answer(2);await question('Noch eine Note speichern?');await answer(1);
 exitProof(()=>{assert.equal(store.getActiveSession(),null);assert.equal(store.readData().sessions.length,2);assert.equal(store.readData().sessions[1].repetitionEntries[0].grade,4.5);});
 await click('#res > :nth-child(3)');
}
main().catch(error=>{console.error(error);app.exit(1);});
