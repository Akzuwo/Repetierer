const fs=require('node:fs');
const path=require('node:path');
const crypto=require('node:crypto');
const {loadDossiers,sameFile}=require('./studentDossiers');
const {renderDossierPdf}=require('./studentDossierPdf');
const {getSharedAssessmentStore}=require('./assessment/assessmentStore');
function safeName(value){return String(value).replace(/[<>:"/\\|?*\x00-\x1f]/g,'_').replace(/[. ]+$/,'').slice(0,100)||'Schülerakte';}
function registerStudentDossiers({ipcMain,BrowserWindow,dialog,getProgram,getPaths,getSessions,getPending}){
  const load=()=>loadDossiers({filePath:getProgram().getCurrentFilePath(),paths:getPaths(),sessions:getSessions().readData().sessions,pending:getPending()});
  ipcMain.handle('dossiers:get',async()=>{
    try{return {ok:true,data:await load()};}catch(error){return {ok:false,error:error.message};}
  });
  function requireContext(data,args){
    if(!sameFile(data.filePath,args.filePath) || !sameFile(getProgram().getCurrentFilePath(),args.filePath) || data.fingerprint!==args.fingerprint || crypto.createHash('sha256').update(fs.readFileSync(args.filePath)).digest('hex')!==args.fingerprint)throw new Error('Die Datei wurde gewechselt oder geändert. Bitte die Schülerakten erneut öffnen.');
  }
  let exporting=false;
  ipcMain.handle('dossiers:set-email',async(_event,args={})=>{
    if(exporting)return {ok:false,error:'Eine Aktenaktion läuft bereits.'};
    exporting=true;
    try{
      const data=await load();requireContext(data,args);
      const cls=data.classes.find(c=>c.name===args.className),student=cls?.students.find(s=>s.id===args.studentId&&s.name===args.studentName);
      if(!student)throw new Error('Die Schülerzuordnung hat sich geändert. Bitte die Schülerakten erneut öffnen.');
      const email=getSharedAssessmentStore(getPaths().assessmentPath).setStudentEmail(data.filePath,cls.name,cls.students,student.id,student.name,args.email,args.expectedEmail);
      return {ok:true,email};
    }catch(error){return {ok:false,error:error.message};}
    finally{exporting=false;}
  });
  const mail=require('./studentDossierMail').createDossierMail({load,requireContext,safeName,isBusy:()=>exporting,setBusy:value=>{exporting=value;},
    readConfig:()=>require('./assessment/mainAssessment').requireMailer(getPaths()),
    createMailer:service=>service,renderPdf:record=>renderDossierPdf(BrowserWindow,record)});
  for(const [channel,handler] of [['dossiers:prepare-mail',args=>mail.prepare(args)],['dossiers:send-mail',args=>mail.send(args.token)]])ipcMain.handle(channel,async(_event,args={})=>{
    try{return {ok:true,...await handler(args)};}catch(error){return {ok:false,error:error.message};}
  });
  ipcMain.handle('dossiers:export',async(event,args={})=>{
    if(exporting)return {ok:false,error:'Ein Export läuft bereits.'};
    exporting=true;const written=[];
    try{
      const data=await load(),scope=args.scope;requireContext(data,args);
      if(!['student','class','all'].includes(scope))throw new Error('Ungültige Exportauswahl.');
      const classes=scope==='all'?data.classes:data.classes.filter(c=>c.name===args.className);
      const records=classes.flatMap(c=>c.students).filter(d=>scope!=='student'||d.id===args.studentId);
      if(!records.length)throw new Error('Keine Schüler für den Export gefunden.');
      if(scope==='student' && records[0].name!==args.studentName)throw new Error('Die Schülerzuordnung hat sich geändert. Bitte die Schülerakten erneut öffnen.');
      const parent=BrowserWindow.fromWebContents(event.sender);let folder,filePath;
      if(scope==='student'){
        const choice=await dialog.showSaveDialog(parent,{title:'Schülerakte als PDF exportieren',defaultPath:safeName(records[0].name)+'.pdf',filters:[{name:'PDF',extensions:['pdf']}]});
        if(choice.canceled||!choice.filePath)return {ok:true,cancelled:true};filePath=choice.filePath;
      }else{
        const choice=await dialog.showOpenDialog(parent,{title:'Zielordner für Schülerakten wählen',properties:['openDirectory','createDirectory']});
        if(choice.canceled||!choice.filePaths?.[0])return {ok:true,cancelled:true};folder=choice.filePaths[0];
        // The selected directory is the export target; existing PDFs are preserved.
      }
      const classFolders=new Map();
      for(const d of records){
        requireContext(data,args);
        let target=filePath;
        if(scope!=='student'){
          let dir=folder;
          if(scope==='all'){
            if(!classFolders.has(d.className)){
              dir=path.join(folder,safeName(d.className));
              if([...classFolders.values()].includes(dir))dir=path.join(folder,`${safeName(d.className)} (${classFolders.size+1})`);
              fs.mkdirSync(dir,{recursive:true});classFolders.set(d.className,dir);
            }else dir=classFolders.get(d.className);
          }
          const base=path.join(dir,`${String(d.id+1).padStart(2,'0')} ${safeName(d.name)}`);target=base+'.pdf';
          for(let suffix=2;fs.existsSync(target);suffix++)target=`${base} (${suffix}).pdf`;
        }
          const bytes=await renderDossierPdf(BrowserWindow,d);requireContext(data,args);
          fs.writeFileSync(target,bytes,{flag:scope==='student'?'w':'wx'});written.push(target);
      }
      return {ok:true,count:written.length,folder,filePaths:written};
    }catch(error){return {ok:false,error:error.message,count:written.length,filePaths:written};}
    finally{exporting=false;}
  });
}
module.exports={registerStudentDossiers,safeName};
