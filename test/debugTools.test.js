const test=require('node:test'),assert=require('node:assert/strict');
const {createDebugUpdate}=require('../src/debugUpdate');
const {getBuildInfo,createBuildMetadata,validBuildDate}=require('../src/buildInfo');
const {buildArguments}=require('../scripts/release-with-icon');
const {buildDebugMessage}=require('../src/assessment/debugMail');

test('version stays semver; stored build identity survives later app starts',()=>{
  const metadata={version:'7.0.0',buildNumber:28,buildDate:'09112026'};
  assert.equal(getBuildInfo(metadata).build,'28.09112026');
  const plan=buildArguments(metadata,{},new Date('2026-11-09T23:30:00Z'));
  assert.equal(plan.info.version,'7.0.0');
  assert.equal(plan.info.build,'28.10112026','Build time uses Zurich date');
  assert.deepEqual(plan.metadataArgs,['--config.extraMetadata.buildNumber=28','--config.extraMetadata.buildDate=10112026']);
  assert.match(plan.artifactArg,/7|\$\{version\}/);
  assert.match(plan.artifactArg,/Build-28\.10112026/);
  const builder=require('electron-builder/out/builder');
  const parsed=builder.configureBuildCommand(builder.createYargs()).parse(buildArguments(metadata,{LOCAL_BUILD_DATE:'09112026'}).metadataArgs);
  assert.equal(parsed.config.extraMetadata.buildDate,'09112026','CLI preserves leading zero');
  assert.equal(createBuildMetadata(metadata,{LOCAL_BUILD_NUMBER:'29',LOCAL_BUILD_DATE:'09112026'}).build,'29.09112026');
  for(const value of ['31112026','00002026',9112026,null])assert.equal(validBuildDate(value),false);
  assert.throws(()=>createBuildMetadata(metadata,{LOCAL_BUILD_NUMBER:'oops'}),/Ungültig/);
});

test('preview, cancellation and normal-mode guard never restart',async()=>{
  const events=[];let restarts=0;
  const make=enabled=>createDebugUpdate({enabled,version:'7.0.0',send:(status,data)=>events.push({status,...data}),restart:()=>restarts++,delay:async()=>{}});
  const normal=make(false);assert.equal(normal.show('restart'),false);assert.equal(await normal.approve(),false);assert.deepEqual(events,[]);
  const debug=make(true);assert.equal(debug.show('preview'),true);assert.equal(await debug.approve(),true);assert.equal(restarts,0);assert.equal(events.at(-1).status,'simulated-preview');
  debug.show('restart');debug.cancel();assert.equal(await debug.approve(),false);assert.equal(restarts,0);
});

test('full simulation produces progress then one real restart even on repeated approval',async()=>{
  const events=[];let restarts=0,release;
  const firstDelay=new Promise(resolve=>release=resolve);let waits=0;
  const debug=createDebugUpdate({enabled:true,version:'7.0.0',send:(status,data)=>events.push({status,...data}),restart:()=>restarts++,delay:()=>waits++===0?firstDelay:Promise.resolve()});
  debug.show('restart');const first=debug.approve();assert.equal(await debug.approve(),true);assert.equal(debug.show('preview'),false);release();await first;
  assert.equal(restarts,1);assert.deepEqual(events.map(e=>e.status),['available','downloading','progress','progress','progress','downloaded','installing']);
  assert.ok(events.every(e=>e.simulated));assert.equal(await debug.approve(),false);
});

test('debug mail uses real result template and fixed sender, one explicit recipient',()=>{
  const message=buildDebugMessage(' demo@example.ch ');
  assert.equal(message.to,'demo@example.ch');assert.equal(message.from.address,'no-reply@akzuwo.ch');
  assert.equal(message.subject,'Mitmachnote – Beispielrunde – Alex Beispiel');
  assert.match(message.html,/<table/);assert.match(message.html,/Beteiligung mit Äusserungen/);assert.match(message.text,/Selbstbeurteilung/);
  for(const address of ['', 'a@example.ch,b@example.ch','a@example.ch\r\nBcc: x@example.ch'])assert.throws(()=>buildDebugMessage(address),/gültige/);
});
