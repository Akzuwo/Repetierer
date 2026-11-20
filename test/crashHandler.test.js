const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),os=require('node:os'),path=require('node:path');
const {createCrashJournal,sanitize}=require('../src/crashJournal'),{monitor}=require('../src/crashWatchdog');
const temp=()=>fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-crash-test-'));
test('clean quit creates no crash notice; forced exit survives a restart and sends one sanitized report',async()=>{
  const dir=temp(),first=createCrashJournal(dir);first.begin('7');first.clean();const second=createCrashJournal(dir);second.begin('7');assert.equal((await second.deliver()).crashed,false);
  second.append(second.run.id,'error',new Error('token=SECRET password=HIDDEN anna@example.ch C:\\Users\\Private\\app.js'));
  const third=createCrashJournal(dir);third.begin('7');const sent=[];const result=await third.deliver(report=>{sent.push(report);});assert.equal(result.crashed,true);assert.equal(result.sent,1);assert.equal(sent.length,1);assert.doesNotMatch(sent[0].log,/SECRET|HIDDEN|anna@example|Private/);assert.ok(fs.existsSync(path.join(dir,`${second.run.id}.log`)));
  third.clean();const fourth=createCrashJournal(dir);fourth.begin('7');const again=await fourth.deliver(report=>sent.push(report));assert.equal(again.crashed,false);assert.equal(sent.length,1);
});
test('renderer faults persist after clean quit; unavailable Brevo and failed SMTP keep reports for later retry',async()=>{
  const dir=temp(),a=createCrashJournal(dir);a.begin('7');a.incident(a.run.id,'renderer-crashed','exit 1');a.clean();const b=createCrashJournal(dir);b.begin('7');assert.equal((await b.deliver()).configured,false);b.clean();const c=createCrashJournal(dir);c.begin('7');assert.equal((await c.deliver(()=>{throw new Error('smtp password=never expose');})).failed,true);c.clean();
  const d=createCrashJournal(dir);d.begin('7');let count=0;assert.equal((await d.deliver(()=>{count++;})).sent,1);assert.equal(count,1);assert.doesNotMatch(fs.readFileSync(path.join(dir,`${a.run.id}.incident.json`),'utf8'),/never expose/);
});
test('separate watchdog records a real stale heartbeat independently and does not repeatedly report it',async()=>{
  const dir=temp(),j=createCrashJournal(dir);j.begin('7');const heartbeat=path.join(dir,`${j.run.id}.heartbeat`);fs.writeFileSync(heartbeat,'tick');fs.utimesSync(heartbeat,new Date(0),new Date(Date.now()-1000));
  const timer=monitor(dir,j.run.id,process.pid,{pollMs:25,stallMs:100});
  try{await new Promise(resolve=>setTimeout(resolve,140));assert.equal(j.pending()[0].kind,'main-unresponsive');assert.equal(j.pending().length,1);fs.writeFileSync(heartbeat,'alive');}finally{clearInterval(timer);}
});
test('diagnostics are bounded and remove bearer tokens and URL credentials',()=>{
  const j=createCrashJournal(temp());j.begin('7');for(let n=0;n<30;n++)j.append(j.run.id,'error','x'.repeat(16000));assert.ok(fs.statSync(path.join(j.dir,`${j.run.id}.log`)).size<220000);assert.doesNotMatch(sanitize('https://login:secret@example.ch/path?token=secret'),/secret/);
});
test('complete authorization schemes and quoted multiword secrets are removed before mailing',()=>{
  assert.doesNotMatch(sanitize({password:'part one " part two secret'}),/part one|part two|secret/);
  assert.doesNotMatch(sanitize('password: "part one \\" part two secret"'),/part one|part two|secret/);
  for(const secret of ['Authorization: Bearer very-secret-opaque-token','Authorization: Basic dXNlcjpwYXNz','password: "two words secret"',"token='several words secret'",'BREVO_SMTP_PASSWORD="line one\nline two"']){
    assert.doesNotMatch(sanitize(secret),/very-secret|dXNlc|two words|words secret|line one|line two/);
  }
});
test('watchdog child records a hang while its parent event loop is actually blocked',async()=>{
  const dir=temp(),journal=createCrashJournal(dir);journal.begin('7');const heartbeat=path.join(dir,`${journal.run.id}.heartbeat`),ready=path.join(dir,'watchdog-ready');fs.writeFileSync(heartbeat,'alive');
  const code=`require(${JSON.stringify(path.resolve(__dirname,'../src/crashWatchdog'))}).monitor(${JSON.stringify(dir)},${JSON.stringify(journal.run.id)},${process.pid},{pollMs:25,stallMs:100});require('node:fs').writeFileSync(${JSON.stringify(ready)},'ready');`;
  const child=require('node:child_process').spawn(process.execPath,['-e',code],{stdio:'ignore',windowsHide:true});
  try{
    for(let i=0;i<100&&!fs.existsSync(ready);i++)await new Promise(r=>setTimeout(r,10));assert.ok(fs.existsSync(ready));
    const end=Date.now()+350;while(Date.now()<end){/* Deliberately block the parent, not the monitor. */}
    assert.equal(journal.pending()[0]?.kind,'main-unresponsive');
  }finally{child.kill();}
});
