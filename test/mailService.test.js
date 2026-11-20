const test=require('node:test'),assert=require('node:assert/strict'),os=require('node:os'),fs=require('node:fs'),path=require('node:path');
const {resolveMailService}=require('../src/assessment/mailService');
const {AssessmentApiClient}=require('../src/assessment/assessmentApiClient');
test('Cloud gewinnt, SMTP wird nicht gelesen; Fehler und unklares Ergebnis fallen nie auf SMTP zurück',async()=>{
 let smtp=0,sends=0;
 const base={profile:fs.mkdtempSync(path.join(os.tmpdir(),'mail-')),getSmtpConfig:()=>{smtp++;return {};},createSmtpMailer:()=>({send:()=>smtp++})};
 const cloud=await resolveMailService({...base,client:{mailCapabilities:async()=>({configured:true,dailyLimit:200}),sendMail:async()=>{sends++;throw Object.assign(new Error('Limit'),{code:'MAIL_DAILY_LIMIT'});}}});
 assert.equal(cloud.mode,'cloud');assert.equal(smtp,0);await assert.rejects(cloud.send({}),/Limit/);assert.equal(smtp,0);assert.equal(sends,1);
 for(const client of [{mailCapabilities:async()=>{throw new Error('offline');}},{mailCapabilities:async()=>({})}])await assert.rejects(resolveMailService({...base,client}));
 assert.equal(smtp,0);
 for(const configured of [{configured:false},{configured:false,legacy:true}])assert.equal((await resolveMailService({...base,client:{mailCapabilities:async()=>configured}})).mode,'smtp');
});
test('Capabilities: eindeutiger alter Server, Authfehler, Offline und ungültiges JSON',async()=>{
 const config={apiUrl:'https://example.invalid/api',teacherToken:'fixture'};
 const legacy=new AssessmentApiClient(config,{fetch:async()=>new Response(JSON.stringify({success:false,error:'not_found'}),{status:404})});assert.equal((await legacy.mailCapabilities()).legacy,true);
 for(const response of [new Response('{}'),new Response('bad',{status:404}),new Response('{"error":"unauthorized"}',{status:401})]){
  const client=new AssessmentApiClient(config,{fetch:async()=>response});await assert.rejects(client.mailCapabilities());
 }
});
test('Eigenes SMTP-Limit bleibt über neue Mailerinstanzen und Neustart erhalten',async()=>{
 const profile=fs.mkdtempSync(path.join(os.tmpdir(),'smtp-limit-'));let sends=0;
 const options={profile,smtpConfig:{},localDailyLimit:2,createSmtpMailer:()=>({send:async()=>{sends++;}})};
 const a=await resolveMailService(options),b=await resolveMailService(options);
 const result=await Promise.allSettled([a.send({}),b.send({}),a.send({})]);assert.equal(result.filter(r=>r.status==='fulfilled').length,2);assert.equal(sends,2);
 await assert.rejects((await resolveMailService(options)).send({}),e=>e.code==='MAIL_DAILY_LIMIT');
});
test('Cloudclient sendet PDF-Inhalt als Base64 und dieselbe Auftrags-ID bei Domain-Fallback',async()=>{
 const calls=[];const config={apiUrl:'https://primary.invalid/api',fallbackApiUrl:'https://fallback.invalid/api',teacherToken:'fixture'};
 const client=new AssessmentApiClient(config,{fetch:async(url,request)=>{calls.push(JSON.parse(request.body));if(url.includes('primary'))throw new Error('offline');return Response.json({success:true,sent:true});}});
 const id=require('crypto').randomUUID(),content=Buffer.from('%PDF-1.7 synthetic');
 await client.sendMail({to:'synthetic@example.invalid',subject:'Akte',text:'Test',attachments:[{filename:'Akte.pdf',content}]},id);
 assert.equal(calls.length,2);assert.equal(calls[0].idempotencyKey,id);assert.deepEqual(calls[0],calls[1]);assert.equal(calls[0].message.attachments[0].content,content.toString('base64'));
});
