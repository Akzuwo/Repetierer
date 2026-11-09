// Test a disposable copy of the shipped Windows runtime, never the installation.
const fs = require('node:fs'), path = require('node:path'), os = require('node:os');
const assert = require('node:assert/strict'), {spawn} = require('node:child_process');
const asar = require('@electron/asar');
const pause = ms => new Promise(r => setTimeout(r, ms));
const root = path.resolve(__dirname, '..');
const baseline = process.argv.includes('--baseline');
const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'repetierer-exe-debug-'));
const runtime = path.join(temp, 'runtime'), profile = path.join(temp, 'profile');
fs.cpSync(path.join(root, 'builds/win-unpacked'), runtime, {recursive:true});
const resources = path.join(runtime, 'resources'), appDir = path.join(resources, 'app');
asar.extractAll(path.join(resources, 'app.asar'), appDir);
fs.renameSync(path.join(resources, 'app.asar'), path.join(resources, 'app.original.asar'));
if (!baseline) {
  for (const name of ['app.js','src','public','package.json']) fs.cpSync(path.join(root,name), path.join(appDir,name), {recursive:true});
}
const entry = path.join(appDir, 'app.js');
fs.mkdirSync(profile);
fs.writeFileSync(entry, `require('electron').app.setPath('userData',${JSON.stringify(profile)});\n` + fs.readFileSync(entry,'utf8'));
let child, socket, port, serial = 0;
const pending = new Map(), errors = [];
function command(method, params = {}) {
  const id = ++serial;
  return new Promise((resolve,reject) => {
    const timer = setTimeout(()=>{pending.delete(id);reject(new Error('CDP timeout: '+method));},10000);
    pending.set(id,{resolve:value=>{clearTimeout(timer);resolve(value);},reject});
    socket.send(JSON.stringify({id,method,params}));
  });
}
async function evaluate(expression) {
  const result = await command('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true});
  if (result.exceptionDetails) throw new Error(JSON.stringify(result.exceptionDetails));
  return result.result.value;
}
async function until(expression) {
  for(let i=0;i<160;i++){try{if(await evaluate(expression))return;}catch{}await pause(50);}
  throw new Error('Timeout: '+expression);
}
async function click(selector) {
  const point = await evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2};})()`);
  await command('Input.dispatchMouseEvent',{type:'mousePressed',button:'left',clickCount:1,...point});
  await command('Input.dispatchMouseEvent',{type:'mouseReleased',button:'left',clickCount:1,...point});
}
async function shot(name) {
  const result = await command('Page.captureScreenshot');
  const out=path.join(root,'.ui-polish-smoke');fs.mkdirSync(out,{recursive:true});
  fs.writeFileSync(path.join(out,`debug-exe-${baseline?'baseline':'current'}-${name}.png`),Buffer.from(result.data,'base64'));
}
async function connect() {
  let page;
  for(let i=0;i<300;i++){try{page=(await(await fetch(`http://127.0.0.1:${port}/json`)).json()).find(p=>p.url.endsWith('/public/index.html'));if(page)break;}catch{}await pause(100);}
  assert.ok(page,'Packaged main window');
  socket=new WebSocket(page.webSocketDebuggerUrl);
  await new Promise(r=>socket.addEventListener('open',r,{once:true}));
  socket.addEventListener('message',event=>{const data=JSON.parse(event.data);if(data.id){const p=pending.get(data.id);pending.delete(data.id);if(data.error)p?.reject(new Error(JSON.stringify(data.error)));else p?.resolve(data.result);}else if(data.method==='Runtime.exceptionThrown')errors.push(data.params.exceptionDetails);});
  await command('Runtime.enable');
}
async function main() {
  await asar.createPackage(appDir,path.join(resources,'app.asar'));
  const net=require('node:net'),server=net.createServer();
  await new Promise(r=>server.listen(0,'127.0.0.1',r));port=server.address().port;await new Promise(r=>server.close(r));
  const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
  child=spawn(path.join(runtime,'Repetierer.exe'),['--debug-mode',`--remote-debugging-port=${port}`],{env,windowsHide:true,stdio:['ignore','pipe','pipe']});
  let output='';child.stdout.on('data',data=>{output+=data;});child.stderr.on('data',data=>{output+=data;});
  child.on('error',e=>console.error(e));
  await connect();
  await until(`!!document.getElementById('debug-launcher')`);
  for (const id of baseline ? ['notification','update-error'] : ['notification-info','notification-warning','notification-error']) {
    if(!baseline)await evaluate(`notify.clear()`);
    await click('#debug-launcher');await pause(500);
    await click(`[data-debug-scenario="${id}"]`);
    if(baseline)await until(`sessionStorage.getItem('repetierer-debug-scenario')===${JSON.stringify(id)} && document.getElementById('debug-launcher')?.textContent==='Debug · Simulation'`);
    else await until(`notify.getActive().some(n=>n.type===${JSON.stringify(id.slice(13))}) && !document.getElementById('debug-panel').classList.contains('modal-open')`);
    await pause(800);await shot(id);
    const state=await evaluate(`({scenario:sessionStorage.getItem('repetierer-debug-scenario'),errors:window.notify?.getActive(),legacyPanel:{display:getComputedStyle(document.getElementById('update-panel')).display,position:getComputedStyle(document.getElementById('update-panel')).position,top:document.getElementById('update-panel').getBoundingClientRect().top,height:innerHeight,text:document.getElementById('update-panel').innerText},headerVisible:document.getElementById('title-bar').checkVisibility({visibilityProperty:true}),contentVisible:document.getElementById('content').checkVisibility({visibilityProperty:true}),repetitionVisible:document.getElementById('repetition').checkVisibility({visibilityProperty:true})})`);
    console.log(JSON.stringify(state));
    if(!baseline){assert.ok(state.errors.length>0,'A visible notification is produced');assert.ok(state.headerVisible && state.contentVisible,'App remains visible');assert.equal(state.repetitionVisible,false,'No repetition without an active session');await until(`document.querySelector('.notification-card')?.dataset.motion==='idle'`);const geometry=await evaluate(`(()=>{const r=document.querySelector('.notification-card').getBoundingClientRect();return {top:r.top,bottom:r.bottom,height:innerHeight};})()`);assert.ok(geometry.top>=0 && geometry.bottom<=geometry.height,'Notification is inside the window');}
  }
  if(!baseline) {
    await command('Page.bringToFront');await until(`!document.hidden`);await pause(700);
    assert.equal(await evaluate(`document.getElementById('version').textContent`),'Version 7.0.0');
    assert.equal(await evaluate(`document.getElementById('build-number').textContent`),'Build 28.09112026');
    for(let cycle=0;cycle<2;cycle++) {
      for(const open of [true,false]) {
        await evaluate(`window.barFrames=[];window.barFrameEnd=performance.now()+450;(function sample(){const e=document.querySelector('#drawer-toggle span');const m=new DOMMatrix(getComputedStyle(e).transform);window.barFrames.push(Math.atan2(m.b,m.a)*180/Math.PI);if(performance.now()<window.barFrameEnd)requestAnimationFrame(sample);})()`);
        await click('#drawer-toggle');await pause(650);
        const frames=await evaluate('window.barFrames');
        assert.ok(frames.some(angle=>angle>1 && angle<44),'Hamburger animates in both directions: '+JSON.stringify(frames));
        assert.ok(Math.abs(frames.at(-1)-(open?45:0))<.1);
        if(open) {
          const centers=await evaluate(`Array.from(document.querySelectorAll('#drawer-toggle span')).map(e=>{const r=e.getBoundingClientRect();return r.y+r.height/2})`);
          assert.ok(Math.abs(centers[0]-centers[2])<.5,'Cross bars share their vertical center');
        }
      }
    }
    // Reverse before completion and finish with reduced motion enabled.
    await click('#drawer-toggle');await pause(70);await click('#drawer-toggle');await pause(650);
    assert.equal(await evaluate(`document.getElementById('drawer-toggle').getAttribute('aria-expanded')`),'false');
    await command('Emulation.setEmulatedMedia',{features:[{name:'prefers-reduced-motion',value:'reduce'}]});
    await click('#drawer-toggle');await pause(100);
    assert.ok(Math.abs(await evaluate(`require('./vendor/ui-motion').gsap.getProperty(document.querySelector('#drawer-toggle span'),'rotation')`)-45)<.1);
    await click('#drawer-toggle');await pause(100);
    await command('Emulation.setEmulatedMedia',{features:[]});
    await click('#debug-launcher');await pause(500);await shot('menu');
    await click('[data-debug-scenario="update-restart"]');await until(`document.getElementById('update-panel').classList.contains('modal-open')`);await pause(500);
    assert.match(await evaluate(`document.getElementById('update-detail').textContent`),/wirklich neu/);
    await shot('update-restart-confirm');
    const oldPid=child.pid;
    await click('#update-primary-btn');
    const completed=path.join(profile,'debug-update-completed.json');
    for(let i=0;i<400 && !fs.existsSync(completed);i++)await pause(100);
    assert.ok(fs.existsSync(completed),'App relaunch completed');
    const proof=JSON.parse(fs.readFileSync(completed,'utf8'));
    assert.equal(proof.previousPid,oldPid);assert.notEqual(proof.pid,oldPid,'Native app process really restarted');
    socket.close();await connect();
    await until(`!!document.getElementById('debug-launcher')`);
    assert.equal(await evaluate(`document.getElementById('build-number').textContent`),'Build 28.09112026');
    assert.equal(fs.existsSync(path.join(profile,'update-state.json')),false,'Real updater state untouched');
    console.log('RESTART_PASS: new native PID, same profile and build, no download or installer');
    await shot('after-restart');
  }
  assert.deepEqual(errors,[],'No renderer exceptions');
  console.log(`DEBUG_EXE_${baseline?'BASELINE':'PASS'}: packaged Windows runtime, isolated profile, native scenario clicks`);
}
main().catch(e=>{console.error(e);process.exitCode=1;}).finally(async()=>{try{if(socket?.readyState===1)await evaluate("require('electron').ipcRenderer.send('quit')");}catch{}socket?.close();child?.kill();});
