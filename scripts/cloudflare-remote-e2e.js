const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const assert = require('node:assert/strict');
const {spawnSync} = require('node:child_process');
const {parseEnv,validateAssessmentConfig} = require('../src/assessment/integrationConfig');
const {AssessmentApiClient} = require('../src/assessment/assessmentApiClient');
const {DEFAULT_CRITERIA} = require('../src/assessment/assessmentStore');
const root=path.join(__dirname,'..');
async function run() {
  const source=path.resolve(process.argv[2] || '');
  const config=validateAssessmentConfig(parseEnv(fs.readFileSync(source,'utf8')));
  const directory=fs.mkdtempSync(path.join(root,'.cloudflare-test','remote-'));
  const credentialFile=path.join(directory,'test-credential.env');
  fs.copyFileSync(source,credentialFile);
  const controlFile=path.join(directory,'control.json');
  fs.writeFileSync(controlFile,JSON.stringify({directory,credentialFile,apiUrl:config.apiUrl}));
  const env={...process.env,NODE_USE_SYSTEM_CA:'1'};delete env.ELECTRON_RUN_AS_NODE;
  const result=spawnSync(require('electron'),[path.join(__dirname,'cloudflare-electron-e2e.js'),controlFile],{cwd:root,env,encoding:'utf8',windowsHide:true,timeout:90000});
  assert.equal(result.status,0,result.stderr+'\n'+result.stdout);
  console.log(result.stdout.trim());
  const client=new AssessmentApiClient(config);
  const connectivity=await client.checkStatus();
  const apiUrl=connectivity.apiUrl;
  const id=crypto.randomUUID();
  const round={externalRoundId:id,title:'E2E atomare Duplikatsperre',criteria:DEFAULT_CRITERIA};
  await client.createRound(round,20);
  try {
    const send=i=>fetch(apiUrl+'/public/rounds/'+id+'/responses',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name:i%2?' E2E   TESTPERSON ':'E2E Testperson',email:'e2e@example.ch',browserToken:String(i).repeat(48),answers:Object.fromEntries(DEFAULT_CRITERIA.map(c=>[c.id,4])),comment:'Synthetischer Test'})});
    const responses=await Promise.all(Array.from({length:8},(_,i)=>send(i)));
    assert.equal(responses.filter(r=>r.status===201).length,1);
    assert.equal(responses.filter(r=>r.status===409).length,7);
    const sync=await client.fetchResponses({externalRoundId:id,criteria:DEFAULT_CRITERIA});
    assert.equal(sync.responses.length,1);
    assert.equal((await client.fetchResponses({externalRoundId:id,criteria:DEFAULT_CRITERIA,knownResponseIds:[sync.responses[0].externalResponseId]})).responses.length,0);
    await client.closeRound(id);
    assert.equal((await send(9)).status,410);
  } finally {await client.deleteRound(id);}
  const health=await fetch(apiUrl+'/health');assert.equal(health.status,200);
  const publicResponse=await fetch(client.getFormUrl());assert.equal(publicResponse.status,200);
  const noAuth=await fetch(apiUrl+'/teacher/me');assert.equal(noAuth.status,401);
  console.log(JSON.stringify({ok:true,apiUrl,fallback:connectivity.fallback,concurrentDuplicates:true,https:true,unauthorizedDenied:true,syntheticDataDeleted:true}));
}
run().catch(error=>{console.error(error.message);process.exitCode=1;});
