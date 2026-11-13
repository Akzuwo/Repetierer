const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs'),os=require('node:os'),path=require('node:path');
const {spawnSync}=require('node:child_process');
const {assertSupportedBuildNode}=require('../scripts/release-with-icon');

test('build Node versions support the system CA environment option',()=>{
 for(const version of ['22.19.0','22.23.0','24.6.0','24.14.0','26.0.0'])assert.doesNotThrow(()=>assertSupportedBuildNode(version));
 for(const version of ['20.19.0','22.12.0','23.11.0','24.5.0'])assert.throws(()=>assertSupportedBuildNode(version),/Node aktualisieren/);
});

for(const scenario of [{initial:null,fail:false},{initial:'0',fail:false},{initial:'0',fail:true}]){
 test(`local PowerShell build scopes CA trust and preserves state: ${JSON.stringify(scenario)}`,{skip:process.platform!=='win32'},()=>{
  const fixture=fs.mkdtempSync(path.join(os.tmpdir(),'repetierer build script '));
  fs.mkdirSync(path.join(fixture,'scripts'));fs.mkdirSync(path.join(fixture,'src'));fs.mkdirSync(path.join(fixture,'builds'));
  for(const file of ['build-local.ps1','scripts/release-with-icon.js','src/buildInfo.js'])fs.copyFileSync(path.join(__dirname,'..',file),path.join(fixture,file));
  fs.writeFileSync(path.join(fixture,'package.json'),JSON.stringify({version:'7.0.0',buildNumber:28}));
  fs.writeFileSync(path.join(fixture,'.local-build-number'),'30');
  fs.writeFileSync(path.join(fixture,'builds','Repetierer-old.exe'),'preserve-original');
  const runner=path.join(fixture,'verify.ps1');
  fs.writeFileSync(runner,`
$ErrorActionPreference='Stop'
$initial=${scenario.initial===null?'$null':"'0'"}
if ($null -eq $initial) { Remove-Item Env:NODE_USE_SYSTEM_CA -ErrorAction SilentlyContinue } else { $env:NODE_USE_SYSTEM_CA=$initial }
$env:PUBLISH_MODE='original-mode'
$env:LOCAL_BUILD_NUMBER='original-number'
$env:LOCAL_BUILD_DATE='original-date'
$global:observed=[System.Collections.Generic.List[object]]::new()
function npm {
 $global:observed.Add(@{command='npm';arguments=@($args);ca=$env:NODE_USE_SYSTEM_CA})
 $global:LASTEXITCODE=0
}
function node {
 if ($args[0] -eq '-e') { & $env:REPETIERER_TEST_NODE @args; $global:LASTEXITCODE=$LASTEXITCODE; return }
 if ($args[0] -ne '.\\scripts\\release-with-icon.js') { throw 'Unexpected executable call' }
 $global:observed.Add(@{command='release';ca=$env:NODE_USE_SYSTEM_CA;publish=$env:PUBLISH_MODE;number=$env:LOCAL_BUILD_NUMBER;date=$env:LOCAL_BUILD_DATE})
 $global:LASTEXITCODE=${scenario.fail?1:0}
}
$failed=$false
try { & (Join-Path $PSScriptRoot 'build-local.ps1') } catch { $failed=$true }
$result=@{failed=$failed;ca=$env:NODE_USE_SYSTEM_CA;publish=$env:PUBLISH_MODE;number=$env:LOCAL_BUILD_NUMBER;date=$env:LOCAL_BUILD_DATE;counter=([IO.File]::ReadAllText((Join-Path $PSScriptRoot '.local-build-number')));calls=@($global:observed.ToArray())}
Write-Output ('BUILD_TEST_RESULT:'+($result|ConvertTo-Json -Depth 5 -Compress))
`);
  const run=spawnSync('powershell.exe',['-NoProfile','-ExecutionPolicy','Bypass','-File',runner],{encoding:'utf8',windowsHide:true,env:{...process.env,REPETIERER_TEST_NODE:process.execPath},timeout:30000});
  assert.equal(run.status,0,(run.stdout+'\n'+run.stderr).slice(-6000));
  const line=run.stdout.split(/\r?\n/).find(line=>line.startsWith('BUILD_TEST_RESULT:'));
  assert.ok(line,(run.stdout+'\n'+run.stderr).slice(-6000));
  const result=JSON.parse(line.slice('BUILD_TEST_RESULT:'.length));
  assert.equal(result.failed,scenario.fail);
  assert.equal(result.ca||null,scenario.initial);
  assert.equal(result.publish,'original-mode');assert.equal(result.number,'original-number');assert.equal(result.date,'original-date');
  assert.equal(result.counter,scenario.fail?'30':'31');
  assert.ok(result.calls.some(call=>call.command==='npm'&&call.arguments[0]==='ci'),'Fresh dependency installation uses scoped trust');
  assert.ok(result.calls.every(call=>call.ca==='1'));
  const release=result.calls.find(call=>call.command==='release');assert.ok(release);
  assert.equal(release.publish,'never');assert.equal(release.number,'31');assert.match(release.date,/^\d{8}$/);
  assert.equal(fs.readFileSync(path.join(fixture,'builds','Repetierer-old.exe'),'utf8'),'preserve-original');
 });
}
