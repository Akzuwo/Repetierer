const fs=require('node:fs'),path=require('node:path'),os=require('node:os'),assert=require('node:assert/strict'),{spawnSync}=require('node:child_process');
const {app}=require('electron');
if(app){
 const profile=process.argv.at(-1);app.setPath('userData',profile);
 app.whenReady().then(()=>{require('../src/localApi/lifecycle').reconcile({app:{isPackaged:false,getAppPath:()=>path.resolve(__dirname,'..'),getPath:name=>app.getPath(name)},profile,enabled:true,startupDirectory:path.join(profile,'Startup')});setTimeout(()=>app.quit(),300);});
}else{
 (async()=>{
  const profile=fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-api-process-'));fs.writeFileSync(path.join(profile,'settings.json'),'{}');
  const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
  const parent=spawnSync(require('electron'),[__filename,profile],{env,windowsHide:true,encoding:'utf8',timeout:20000});assert.equal(parent.status,0,parent.stderr);
  let token;for(let i=0;i<100;i++){try{token=fs.readFileSync(path.join(profile,'local-api-token'),'utf8');break;}catch{}await new Promise(r=>setTimeout(r,100));}
  assert.ok(token,'Service must start after parent exits');const url='http://127.0.0.1:17839/v1/health',headers={Authorization:'Bearer '+token};
  try{
   let healthy=false;for(let i=0;i<100;i++){try{if((await fetch(url,{headers})).ok){healthy=true;break;}}catch{}await new Promise(r=>setTimeout(r,100));}assert.ok(healthy);
   assert.match(fs.readFileSync(path.join(profile,'Startup','Repetierer API.vbs'),'utf8'),/--api-profile/);
   fs.writeFileSync(path.join(profile,'settings.json'),'{"localApiEnabled":false}');
   let stopped=false;for(let i=0;i<100;i++){try{await fetch(url,{headers});}catch{stopped=true;break;}await new Promise(r=>setTimeout(r,100));}assert.ok(stopped);
   console.log('LOCAL_API_PROCESS_PASS: parent exited, headless Electron service still served HTTP, disable stopped it; temporary profile only');
  }finally{fs.writeFileSync(path.join(profile,'settings.json'),'{"localApiEnabled":false}');}
 })().catch(e=>{console.error(e.message);process.exitCode=1;});
}
