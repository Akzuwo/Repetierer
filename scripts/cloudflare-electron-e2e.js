// Real Electron IPC + real browser survey, against a local/staging/production Worker.
// Only synthetic pupil data. Pass a JSON control-file path, never a secret on the CLI.
const {app, BrowserWindow, dialog, ipcMain, net, safeStorage} = require('electron');
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const control = JSON.parse(fs.readFileSync(process.argv[2],'utf8'));
const directory = control.directory;
app.setPath('userData',path.join(directory,'electron-profile'));
const pause = ms=>new Promise(r=>setTimeout(r,ms));
let teacherWindow, pupilWindow;
let roundId;
const errors=[];
async function waitFor(win,expr) {
  for(let i=0;i<200;i++){if(await win.webContents.executeJavaScript(expr,true))return;await pause(100);}
  throw new Error('Browser-Zustand nicht erreicht: '+expr);
}
async function run() {
  assert.ok(safeStorage.isEncryptionAvailable(),'OS-Schlüsselspeicher muss aktiv sein');
  const paths={assessmentPath:path.join(directory,'mitmachnoten-ipc.json'),integrationEnvPath:path.join(directory,'repetierer.env'),brevoEnvPath:path.join(directory,'brevo-smtp.env')};
  const students=[{id:0,name:'E2E Testperson'}];
  const program={getCurrentFilePath:()=>path.join(directory,'class.xlsx'),getClassPersons:()=>students};
  // The picker is the only replaced UI boundary. Production IPC handlers are unchanged.
  dialog.showOpenDialogSync=()=>[control.credentialFile];
  require('../src/assessment/mainAssessment').registerAssessmentHandlers({getProgram:()=>program,getPaths:()=>paths});
  teacherWindow=new BrowserWindow({show:false,webPreferences:{nodeIntegration:true,contextIsolation:false}});
  await teacherWindow.loadURL('data:text/html,<html lang="de"><body>Electron E2E</body></html>');
  const invoke=async(channel,args={})=>{
    const result=await teacherWindow.webContents.executeJavaScript(`require('electron').ipcRenderer.invoke(${JSON.stringify(channel)},${JSON.stringify(args)})`,true);
    assert.ok(result.ok,result.error?.message);return result.data;
  };
  // HTTPS is required for real imports; local HTTP tests prepare encrypted config directly.
  if(control.apiUrl.startsWith('https:')) await invoke('assessment:config-import');
  else {
    const {parseEnv}=require('../src/assessment/integrationConfig');
    const values=parseEnv(fs.readFileSync(control.credentialFile,'utf8'));
    const {protectTeacherToken}=require('../src/assessment/teacherCredentials');
    fs.writeFileSync(paths.integrationEnvPath,`ASSESSMENT_API_URL=https://local-test.invalid/app/services/repetierer/api\nASSESSMENT_TEACHER_TOKEN=${protectTeacherToken(values.ASSESSMENT_TEACHER_TOKEN,safeStorage)}\n`);
    const actualFetch=net.fetch.bind(net);
    net.fetch=(url,options)=>actualFetch(String(url).replace('https://local-test.invalid/app/services/repetierer/api',control.apiUrl),options);
  }
  const stored=fs.readFileSync(paths.integrationEnvPath,'utf8');
  assert.match(stored,/ASSESSMENT_TEACHER_TOKEN="?encrypted:/);
  assert.equal(stored.includes('rt_'),false);
  const status=await invoke('assessment:config-status');
  assert.equal(status.configured,true);assert.equal(Object.hasOwn(status,'teacherToken'),false);
  await invoke('assessment:api-test');
  const originalFetch=net.fetch;
  net.fetch=(url,options)=>{
    if(options?.method==='PUT') return Promise.reject(new Error('Simulated interrupted publish'));
    return originalFetch(url,options);
  };
  const interrupted=await teacherWindow.webContents.executeJavaScript(`require('electron').ipcRenderer.invoke('assessment:create-round',${JSON.stringify({className:'E2E',title:'Electron E2E '+Date.now()})})`,true);
  assert.equal(interrupted.ok,false);
  net.fetch=originalFetch;
  const overview=await invoke('assessment:get-overview',{className:'E2E'});
  assert.equal(overview.rounds.length,1);assert.equal(overview.rounds[0].cloudStatus,'pending');
  roundId=overview.rounds[0].id;
  let view=await invoke('assessment:get-round',{className:'E2E',roundId});
  assert.equal(view.round.cloudStatus,'open');assert.match(view.invitation.text,/round=/);
  assert.equal(view.round.cloudFormUrl,view.invitation.text.match(/https:\/\/\S+/)[0].split('?')[0]);
  const persisted=JSON.parse(fs.readFileSync(paths.assessmentPath,'utf8'));
  const persistedRound=Object.values(Object.values(persisted.workbooks)[0].classes)[0].rounds[0];
  assert.equal(persistedRound.cloudFormUrl,view.round.cloudFormUrl);
  const url=control.apiUrl.startsWith('https:') ? view.invitation.text.match(/https:\/\/\S+/)[0] : control.apiUrl.replace(/api$/,'')+'?round='+view.round.externalRoundId;
  pupilWindow=new BrowserWindow({show:false,width:1100,height:1000,webPreferences:{nodeIntegration:false,contextIsolation:true,sandbox:true}});
  pupilWindow.webContents.on('console-message',event=>{if(event.level==='error') errors.push(event.message);});
  await pupilWindow.loadURL(url);
  await waitFor(pupilWindow,"!document.getElementById('survey').hidden");
  assert.equal(await pupilWindow.webContents.executeJavaScript('document.querySelectorAll("fieldset").length'),7);
  for(const [name,width,height] of [['desktop',1100,1000],['mobile',390,844]]) {
    pupilWindow.setSize(width,height);await pause(150);
    assert.equal(await pupilWindow.webContents.executeJavaScript('document.documentElement.scrollWidth <= innerWidth'),true);
    assert.equal(await pupilWindow.webContents.executeJavaScript('Array.from(document.querySelectorAll(".choice span")).every(e=>e.getBoundingClientRect().height>=44)'),true);
    assert.equal(await pupilWindow.webContents.executeJavaScript('Math.abs(document.getElementById("submit").getBoundingClientRect().right-document.querySelector(".submit-actions").getBoundingClientRect().right)<2'),true);
    assert.equal(await pupilWindow.webContents.executeJavaScript('getComputedStyle(document.querySelector("footer")).textAlign'), 'center');
    if (name==='desktop') assert.ok(await pupilWindow.webContents.executeJavaScript('document.querySelector(".choice span").getBoundingClientRect().height < 85'));
    fs.writeFileSync(path.join(directory,`survey-${name}.png`),(await pupilWindow.webContents.capturePage()).toPNG());
  }
  await pupilWindow.webContents.executeJavaScript(`document.querySelector('[name=name]').value='E2E Testperson';document.querySelector('[name=email]').value='e2e@example.ch';document.querySelectorAll('fieldset').forEach(f=>f.querySelector('[value="4"]').checked=true);document.getElementById('survey').requestSubmit();`,true);
  await waitFor(pupilWindow,"document.getElementById('status').textContent.includes('Vielen Dank')");
  await pupilWindow.reload();
  await waitFor(pupilWindow,"document.getElementById('status').textContent.includes('bereits abgegeben')");
  view=await invoke('assessment:sync-responses',{className:'E2E',roundId});
  assert.equal(view.round.externalResponses.length,1);
  view=await invoke('assessment:sync-responses',{className:'E2E',roundId});
  assert.equal(view.round.externalResponses.length,1);assert.equal(view.round.lastSyncResult.added,0);
  const evaluated=await invoke('assessment:save-teacher',{className:'E2E',roundId,studentId:view.students[0].id,answers:Object.fromEntries(view.round.criteria.map(c=>[c.id,3])),comment:'E2E Lehrerkommentar',criterionComments:{participation_frequency:'E2E Kriteriumskommentar'}});
  assert.equal(evaluated.round.assessments[view.students[0].id].teacherCriterionComments.participation_frequency,'E2E Kriteriumskommentar');
  view=await invoke('assessment:close-round',{className:'E2E',roundId});
  assert.equal(view.round.cloudStatus,'closed');
  await pupilWindow.reload();
  await waitFor(pupilWindow,"document.getElementById('status').textContent.includes('geschlossen')");
  assert.equal(errors.length,0,errors.join('\n'));
  await invoke('assessment:delete-round',{className:'E2E',roundId});roundId=null;
  assert.equal((await invoke('assessment:get-overview',{className:'E2E'})).rounds.length,0);
  console.log(JSON.stringify({ok:true,ipc:true,browser:true,encryptedCredentials:true,responsive:true,syntheticDataDeleted:true,screenshots:directory}));
}
app.whenReady().then(run).then(()=>app.quit()).catch(async error=>{
  // Clean up an online synthetic test round even if a later assertion fails.
  if(roundId&&teacherWindow&&!teacherWindow.isDestroyed()) {
    try{await teacherWindow.webContents.executeJavaScript(`require('electron').ipcRenderer.invoke('assessment:delete-round',${JSON.stringify({className:'E2E',roundId})})`,true);}catch{}
  }
  console.error(error.stack);app.exit(1);
});
