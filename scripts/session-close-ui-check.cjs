const {app,BrowserWindow}=require('electron');
const fs=require('node:fs'),path=require('node:path'),os=require('node:os'),assert=require('node:assert/strict');
if(!app){
 const profile=fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-close-ui-'));
 const env={...process.env,REPETIERER_CLOSE_TEST_PROFILE:profile};delete env.ELECTRON_RUN_AS_NODE;
 const run=require('node:child_process').spawnSync(require('electron'),[__filename],{env,stdio:'inherit',windowsHide:true,timeout:90000});
 assert.equal(run.status,0);assert.equal(JSON.parse(fs.readFileSync(path.join(profile,'close-proof.json'),'utf8')).sessionPreserved,true);
 console.log('SESSION_CLOSE_UI_PASS: real app, close button, native window close, cancel/Escape, focus, two desktop sizes and confirmed quit preserving session');process.exit(0);
}
const profile=process.env.REPETIERER_CLOSE_TEST_PROFILE;app.setPath('userData',profile);
const pause=ms=>new Promise(resolve=>setTimeout(resolve,ms));let win;
async function until(fn){for(let i=0;i<450;i++){try{if(await fn())return;}catch{}await pause(60);}throw new Error('Session close UI timeout');}
const evaluate=code=>win.webContents.executeJavaScript(code,true);
async function click(selector){const point=await evaluate(`(()=>{const r=document.querySelector(${JSON.stringify(selector)}).getBoundingClientRect();return {x:Math.round(r.x+r.width/2),y:Math.round(r.y+r.height/2)}})()`);win.webContents.sendInputEvent({type:'mouseDown',button:'left',clickCount:1,...point});win.webContents.sendInputEvent({type:'mouseUp',button:'left',clickCount:1,...point});}
async function main(){
 require('../app');await app.whenReady();
 await until(()=>{win=BrowserWindow.getAllWindows().find(w=>w.webContents.getURL().endsWith('/public/index.html'));return !!win;});
 await until(()=>evaluate(`!!window.getSettingsScopeContext && getComputedStyle(document.getElementById('container')).opacity==='1'`));await pause(800);
 const store=require('../src/teachingSessions').createTeachingSessionStore(require('../src/storage').getPaths().teachingSessionsPath);
 const session=store.startSession('Fixture Klasse',[{id:0,name:'Beispielperson'}],[],path.join(profile,'fixture.xlsx')).session;
 const bytes=fs.readFileSync(require('../src/storage').getPaths().teachingSessionsPath);
 for(const size of [[1100,800],[1440,900]]){
  win.setSize(...size);await pause(250);await click('#quit');
  await until(()=>evaluate(`document.getElementById('session-close-modal').classList.contains('modal-open')`));await pause(450);
  assert.match(await evaluate(`document.getElementById('session-close-text').textContent`),/Fixture Klasse/);
  assert.equal(await evaluate(`document.activeElement.id`),'session-close-cancel-btn');
  assert.equal(await evaluate(`(()=>{const r=document.querySelector('#session-close-modal .modal-content').getBoundingClientRect();return r.left>=0&&r.right<=innerWidth&&r.top>=0&&r.bottom<=innerHeight})()`),true);
  const out=path.join(__dirname,'../.ui-polish-smoke');fs.mkdirSync(out,{recursive:true});fs.writeFileSync(path.join(out,'session-close-'+size[0]+'.png'),(await win.webContents.capturePage()).toPNG());
  await click('#session-close-cancel-btn');await until(()=>evaluate(`getComputedStyle(document.getElementById('session-close-modal')).display==='none'`));
  assert.equal(win.isDestroyed(),false);assert.equal(store.getActiveSession().id,session.id);
 }
 await evaluate(`require('./uiMotion').openDialog(document.getElementById('session-end-modal'))`);await pause(500);
 win.close();await until(()=>evaluate(`document.getElementById('session-close-modal').classList.contains('modal-open')`));await pause(500);
 assert.equal(await evaluate(`(()=>{const b=document.getElementById('session-close-cancel-btn'),r=b.getBoundingClientRect();return document.elementFromPoint(r.x+r.width/2,r.y+r.height/2).closest('button')===b})()`),true,'Quit confirmation remains above another modal');
 win.webContents.sendInputEvent({type:'keyDown',keyCode:'Escape'});win.webContents.sendInputEvent({type:'keyUp',keyCode:'Escape'});
 await until(()=>evaluate(`getComputedStyle(document.getElementById('session-close-modal')).display==='none'`));assert.equal(win.isDestroyed(),false);
 await evaluate(`require('./uiMotion').closeDialog(document.getElementById('session-end-modal'))`);await pause(400);
 win.close();await until(()=>evaluate(`document.getElementById('session-close-modal').classList.contains('modal-open')`));await pause(450);
 app.once('will-quit',()=>{
  assert.deepEqual(fs.readFileSync(require('../src/storage').getPaths().teachingSessionsPath),bytes);
  assert.equal(store.getActiveSession().id,session.id);
  fs.writeFileSync(path.join(profile,'close-proof.json'),JSON.stringify({sessionPreserved:true}));
 });
 await click('#session-close-confirm-btn');
}
main().catch(error=>{console.error(error);app.exit(1);});
