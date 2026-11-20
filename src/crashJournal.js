const fs=require('node:fs'),path=require('node:path'),crypto=require('node:crypto');
const LIMIT=200000;
function sanitize(value){
  const text=value instanceof Error?value.stack||value.message:typeof value==='string'?value:JSON.stringify(value,(key,item)=>/^(?:password|pass|token|authorization|proxy-authorization|api[_-]?key|secret|BREVO_SMTP_USER|BREVO_SMTP_PASSWORD)$/i.test(key)?'[Zugangsdaten]':item);
  return String(text??'Unbekannter Fehler')
    .replace(/(?:[A-Z]:[\\/]Users[\\/]|\/home\/)[^\\/\s]+/gi,'[Nutzerverzeichnis]')
    .replace(/(\b(?:authorization|proxy-authorization)\s*["']?\s*[:=]\s*["']?)(?:(?:Bearer|Basic)\s+)?[^\r\n,"'}]+/gi,'$1[Zugangsdaten]')
    .replace(/((?:password|pass|token|api[_-]?key|secret|BREVO_SMTP_USER|BREVO_SMTP_PASSWORD)\s*["']?\s*[:=]\s*)("(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*')/gi,'$1"[Zugangsdaten]"')
    .replace(/\b(?:Bearer|Basic)\s+[a-z0-9+/=._-]+/gi,'[Zugangsdaten]')
    .replace(/\b[\w.+-]+@[\w.-]+\.[a-z]{2,}\b/gi,'[E-Mail]')
    .replace(/\b(?:rt_|xkeysib-|xsmtpsib-|encrypted:)[a-z0-9+/=_:.-]+/gi,'[Zugangsdaten]')
    .replace(/((?:password|pass|token|authorization|api[_-]?key|secret|BREVO_SMTP_USER|BREVO_SMTP_PASSWORD)\s*["']?\s*[:=]\s*["']?)[^\s,"'}]+/gi,'$1[Zugangsdaten]')
    .replace(/(https?:\/\/)[^\s/@]+:[^\s/@]+@/gi,'$1[Zugangsdaten]@')
    .replace(/(https?:\/\/[^\s?#]+)[?#][^\s]+/gi,'$1[Parameter entfernt]').slice(0,16000);
}
function readJson(file,fallback){try{return JSON.parse(fs.readFileSync(file,'utf8'));}catch{return fallback;}}
function atomic(file,value){const temp=file+'.tmp-'+process.pid;fs.writeFileSync(temp,JSON.stringify(value,null,2));fs.renameSync(temp,file);}
function createCrashJournal(dir){
  fs.mkdirSync(dir,{recursive:true});
  const marker=path.join(dir,'run.json');let run;
  function append(runId,kind,message){
    try{
      const file=path.join(dir,`${runId}.log`);
      if(fs.existsSync(file)&&fs.statSync(file).size>=LIMIT)return;
      fs.appendFileSync(file,`${new Date().toISOString()} [${kind}] ${sanitize(message)}\n`);
    }catch{/* Fault reporting must not cause another crash. */}
  }
  function incident(runId,kind,message){
    append(runId,kind,message);
    const file=path.join(dir,`${runId}.incident.json`),old=readJson(file,{});
    atomic(file,{...old,id:runId,createdAt:old.createdAt||new Date().toISOString(),kind,notified:false,sent:old.sent===true});
  }
  function begin(version){
    const previous=readJson(marker,null);
    if(previous?.running)incident(previous.id,'unclean-exit','Der vorherige Prozess wurde nicht ordnungsgemäss beendet (Absturz, erzwungenes Beenden oder Stromunterbruch).');
    run={id:crypto.randomUUID(),pid:process.pid,version,startedAt:new Date().toISOString(),running:true};
    atomic(marker,run);append(run.id,'start',`Repetierer ${version}; ${process.platform} ${process.arch}; Electron ${process.versions.electron||'-'}`);return run;
  }
  function clean(){if(run){const current=readJson(marker,{});if(current.id===run.id)atomic(marker,{...run,running:false});}}
  function pending(){return fs.readdirSync(dir).filter(f=>f.endsWith('.incident.json')).map(f=>readJson(path.join(dir,f),null)).filter(Boolean);}
  async function deliver(send){
    const incidents=pending().filter(item=>!run||item.id!==run.id);
    for(const item of incidents){
      if(!item.sent&&send){
        try{
          const logPath=path.join(dir,`${item.id}.log`),bytes=fs.readFileSync(logPath).subarray(0,LIMIT);
          await send({id:item.id,log:sanitizeLog(bytes.toString('utf8'))});item.sent=true;item.sentAt=new Date().toISOString();delete item.sendError;
        }catch{item.sendError='Versand fehlgeschlagen. Erneuter Versuch beim nächsten Start.';}
      }
      atomic(path.join(dir,`${item.id}.incident.json`),item);
    }
    const notices=incidents.filter(item=>!item.notified);
    for(const item of notices)atomic(path.join(dir,`${item.id}.incident.json`),{...item,notified:true});
    return {crashed:notices.length>0,count:notices.length,sent:incidents.filter(i=>i.sent).length,failed:incidents.some(i=>i.sendError),configured:!!send};
  }
  return {dir,marker,begin,clean,pending,deliver,append,incident,get run(){return run;}};
}
function sanitizeLog(text){return text.split('\n').map(line=>sanitize(line)).join('\n');}
module.exports={createCrashJournal,sanitize,sanitizeLog};
