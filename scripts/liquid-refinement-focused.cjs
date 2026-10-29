const {app,BrowserWindow,ipcMain}=require('electron');
if(!app){const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;const r=require('node:child_process').spawnSync(require('electron'),[__filename,...process.argv.slice(2)],{env,stdio:'inherit',windowsHide:true});process.exit(r.status??1);}
const fs=require('node:fs'),path=require('node:path'),os=require('node:os'),assert=require('node:assert/strict');
app.setPath('userData',fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-liquid-')));
const ok=data=>({ok:true,data}),pause=ms=>new Promise(r=>setTimeout(r,ms));let win;
const bundleFixture=path.join(app.getPath('userData'),'shared.repetierer');
const values={ASSESSMENT_API_URL:'https://example.ch/api',ASSESSMENT_TEACHER_TOKEN:'rt_'+'a'.repeat(43),BREVO_SMTP_HOST:'smtp.example.ch',BREVO_SMTP_PORT:'587',BREVO_SMTP_USER:'teacher@example.ch',BREVO_SMTP_PASSWORD:'fixture-only',BREVO_FROM_EMAIL:'teacher@example.ch',BREVO_FROM_NAME:'Demo'};
ipcMain.handle('ai:status',()=>({state:'ready',message:'Bereit'}));
ipcMain.handle('assessment:config-status',()=>ok({configured:false}));ipcMain.handle('assessment:brevo-status',()=>ok({configured:false}));
if(process.argv.includes('--assessment-polish')) {
 const {AssessmentStore}=require('../src/assessment/assessmentStore'),{presentClass,summarizeRound}=require('../src/assessment/mainAssessment');
 const store=new AssessmentStore(path.join(app.getPath('userData'),'assessment.json')),workbook=path.join(app.getPath('userData'),'class.xlsx');
 const people=Array.from({length:28},(_,i)=>({id:i+6,name:`Schüler ${i+1} Beispiel`}));
 const round=store.createRound(workbook,'Demo A',people,{title:'Herbstsemester 2026'});
 const view=()=>{const {classRecord,round:current}=store.getRound(workbook,'Demo A',people,round.id);return {classRecord:presentClass(classRecord),round:current,students:classRecord.students,summary:summarizeRound(current,classRecord.students),invitation:null,assessmentConfigured:false,deliveryValidation:{valid:false,incomplete:people}};};
 ipcMain.handle('assessment:get-overview',()=>ok(presentClass(store.getClassContext(workbook,'Demo A',people))));
 ipcMain.handle('assessment:get-round',()=>ok(view()));
 ipcMain.handle('assessment:save-teacher',(_event,args)=>{try{store.saveTeacherAssessment(workbook,'Demo A',people,args.roundId,args.studentId,args.answers,args.comment,args.criterionComments);return ok(view());}catch(error){return {ok:false,error:{message:error.message}};}});
}
ipcMain.handle('assessment:config-export',()=>ok(require('../src/assessment/integrationConfig').exportIntegrationFile(bundleFixture,values)));
ipcMain.handle('assessment:config-import',async()=>{require('../src/assessment/integrationConfig').exportIntegrationFile(bundleFixture,values);return ok(await require('../src/assessment/integrationConfig').importIntegrationEnv(bundleFixture,path.join(app.getPath('userData'),'imported.env')));});
ipcMain.on('get-settings',event=>event.reply('settings-data',{theme:'midnight',probabilityDecreaseFactor:1.8,neverSelectedBoostFactor:1.5,visualEffectsEnabled:true,wiggersRuleDurationMinutes:30},{},{}));
ipcMain.on('load-saved-file',event=>event.reply('saved-file-missing',''));
ipcMain.on('get-active-teaching-session',event=>event.reply('active-teaching-session-data',null));
const evaluate=code=>win.webContents.executeJavaScript(code,true);
async function waitFor(code){for(let i=0;i<180;i++){if(await evaluate(code))return;await pause(30);}throw new Error('Timeout: '+code);}
async function shot(name){await pause(100);win.webContents.invalidate();await pause(50);fs.writeFileSync(path.join(__dirname,'..','.ui-polish-smoke',name+'.png'),(await win.webContents.capturePage()).toPNG());}
async function main(){await app.whenReady();win=new BrowserWindow({width:1100,height:800,show:false,frame:false,webPreferences:{nodeIntegration:true,contextIsolation:false,offscreen:true,backgroundThrottling:false}});const errors=[];win.webContents.on('console-message',e=>{if(e.level==='error'||e.level===3)errors.push(e.message);});await win.loadFile(path.join(__dirname,'..','public','index.html'));await waitFor(`!!window.closeRepetiererSettings`);win.webContents.debugger.attach('1.3');const check=process.argv.includes('--stats-drawer')?'./stats-drawer-ui-check.cjs':process.argv.includes('--ai-flight')?'./ai-flight-ui-check.cjs':process.argv.includes('--assessment-polish')?'./assessment-polish-ui-check.cjs':process.argv.includes('--settings-sliders')?'./settings-range-ui-check.cjs':'./liquid-refinement-ui-check.cjs';await require(check)({win,evaluate,waitFor,shot,pause,bundleFixture});assert.deepEqual(errors,[]);}
const timer=setTimeout(()=>{console.error('Focused liquid check timed out');app.exit(1);},45000);
main().then(()=>{clearTimeout(timer);app.exit(0);}).catch(e=>{console.error(e);clearTimeout(timer);app.exit(1);});
