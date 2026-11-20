const test=require('node:test'),assert=require('node:assert/strict');
const {createDossierMail}=require('../src/studentDossierMail');
function fixture(){
 const data={filePath:'fixture.xlsx',fingerprint:'original',classes:[{name:'A',students:[{id:0,name:'Anna',className:'A',email:'anna@example.ch'},{id:1,name:'Max',className:'A',email:'max@example.ch'}]},{name:'B',students:[{id:0,name:'Eva',className:'B',email:'eva@example.ch'}]}]};
 let busy=false,fail='',duringPdf=()=>{},config=true;const mails=[];
 const service=createDossierMail({load:async()=>structuredClone(data),requireContext:(d,args)=>{if(args.filePath!==d.filePath||args.fingerprint!==d.fingerprint)throw new Error('Datei geändert');},readConfig:()=>{if(!config)throw new Error('Brevo fehlt');return {fromName:'Repetierer',fromEmail:'sender@example.ch'};},createMailer:()=>({send:async message=>{if(message.to===fail)throw {code:'EAUTH',message:'secret credential'};mails.push(message);return {accepted:[message.to]};}}),renderPdf:async d=>{await duringPdf(d);return Buffer.from('%PDF '+d.name);},safeName:s=>s,isBusy:()=>busy,setBusy:v=>{busy=v;}});
 return {data,mails,service,args:{filePath:data.filePath,fingerprint:data.fingerprint,scope:'all',className:'A',studentId:0,studentName:'Anna'},setFail:v=>fail=v,setDuringPdf:fn=>duringPdf=fn,setConfig:v=>config=v};
}
test('preview never sends; all scopes create private single-recipient PDFs; tokens cannot replay',async()=>{
 for(const [scope,count] of [['student',1],['class',2],['all',3]]){
  const f=fixture(),preview=await f.service.prepare({...f.args,scope});assert.equal(preview.recipients.length,count);assert.equal(f.mails.length,0);
  const result=await f.service.send(preview.token);assert.equal(result.sent,count);assert.equal(result.remaining,0);
  for(const mail of f.mails){assert.equal(typeof mail.to,'string');assert.equal(mail.cc,undefined);assert.equal(mail.bcc,undefined);assert.equal(mail.attachments.length,1);const person=f.data.classes.flatMap(c=>c.students).find(s=>s.email===mail.to);assert.equal(mail.attachments[0].content.toString(),'%PDF '+person.name);}
  await assert.rejects(f.service.send(preview.token),/abgelaufen/);
 }
});
test('missing addresses block the batch; missing config and wrong context cannot send',async()=>{
 const f=fixture();f.data.classes[0].students[1].email='';const p=await f.service.prepare(f.args);assert.equal(p.token,null);assert.deepEqual(p.missing,[{name:'Max',className:'A'}]);assert.equal(f.mails.length,0);
 f.setConfig(false);await assert.rejects(f.service.prepare(f.args),/Brevo/);f.setConfig(true);await assert.rejects(f.service.prepare({...f.args,fingerprint:'wrong'}),/Datei/);
});
test('changed address after preview or during PDF rendering prevents sending to stale recipients',async()=>{
 const f=fixture(),p=await f.service.prepare(f.args);f.data.classes[0].students[0].email='new@example.ch';const r=await f.service.send(p.token);assert.equal(r.sent,0);assert.match(r.error,/geändert/);assert.equal(f.mails.length,0);
 const g=fixture(),q=await g.service.prepare(g.args);g.setDuringPdf(()=>{g.data.fingerprint='changed';});const s=await g.service.send(q.token);assert.equal(s.sent,0);assert.equal(s.remaining,3);assert.equal(g.mails.length,0);
});
test('partial failure retries only unsent pupils, sanitizes SMTP errors and serializes concurrent sends',async()=>{
 const f=fixture();f.setFail('max@example.ch');const p=await f.service.prepare(f.args),r=await f.service.send(p.token);assert.equal(r.sent,2);assert.equal(r.failed,1);assert.equal(r.remaining,1);assert.doesNotMatch(JSON.stringify(r),/secret credential/);
 f.setFail('');const retry=await f.service.prepare({...f.args,retryToken:r.retryToken});assert.deepEqual(retry.recipients.map(p=>p.name),['Max']);assert.equal((await f.service.send(retry.token)).sent,1);assert.equal(f.mails.length,3);
 const g=fixture();let release,started;const waiting=new Promise(r=>started=r);g.setDuringPdf(()=>{started();return new Promise(r=>release=r);});const q=await g.service.prepare({...g.args,scope:'student'}),sending=g.service.send(q.token);await waiting;await assert.rejects(g.service.send(q.token),/läuft bereits/);await assert.rejects(g.service.prepare(g.args),/läuft bereits/);release();assert.equal((await sending).sent,1);
});

test('asynchronous mail configuration is awaited for preview and confirmed delivery',async()=>{
 const f=fixture();let busy=false,checks=0,sent=0;
 const service=createDossierMail({load:async()=>structuredClone(f.data),requireContext:()=>{},readConfig:async()=>{await Promise.resolve();checks++;return {fromName:'Repetierer',fromEmail:'sender@example.ch'};},createMailer:config=>{assert.equal(config.fromEmail,'sender@example.ch');return {send:async message=>{assert.equal(message.from.address,'sender@example.ch');sent++;}};},renderPdf:async()=>Buffer.from('%PDF'),safeName:s=>s,isBusy:()=>busy,setBusy:v=>busy=v});
 const preview=await service.prepare(f.args);assert.equal(checks,1);assert.equal(sent,0);
 assert.equal((await service.send(preview.token)).sent,3);assert.equal(checks,2);assert.equal(sent,3);
 const unavailable=createDossierMail({load:async()=>f.data,requireContext:()=>{},readConfig:async()=>{throw new Error('Cloud nicht erreichbar');},isBusy:()=>false});
 await assert.rejects(unavailable.prepare(f.args),/Cloud nicht erreichbar/);
});
