const path=require('node:path'),fs=require('node:fs');
const {createCrashJournal}=require('./crashJournal');
function installCrashHandler({app,ipcMain,dialog,crashReporter}){
  const journal=createCrashJournal(path.join(app.getPath('userData'),'crashlogs'));
  const run=journal.begin(app.getVersion());
  const record=(kind,error,fatal=false)=>{try{if(fatal)journal.incident(run.id,kind,error);else journal.append(run.id,kind,error);}catch{}};
  const original=console.error;
  console.error=(...args)=>{for(const arg of args)record('main-error',arg);original.apply(console,args);};
  process.on('uncaughtException',error=>{record('uncaught-exception',error,true);app.exit(1);});
  process.on('unhandledRejection',error=>record('unhandled-rejection',error,true));
  try{crashReporter?.start({productName:'Repetierer',uploadToServer:false,compress:true});}catch(error){record('crash-reporter',error);}
  ipcMain.on('crash:renderer-error',(event,payload)=>{
    if(!event.sender.getURL().startsWith('file:'))return;
    record('renderer-error',typeof payload==='string'?payload.slice(0,16000):'Unbekannter Rendererfehler',true);
  });
  const heartbeat=path.join(journal.dir,`${run.id}.heartbeat`);
  const tick=()=>{try{fs.writeFileSync(heartbeat,String(Date.now()));}catch{}};tick();
  const interval=setInterval(tick,3000);interval.unref();let watchdog;
  try{
    watchdog=require('node:child_process').spawn(process.execPath,[path.join(__dirname,'crashWatchdog.js'),journal.dir,run.id,String(process.pid)],{env:{...process.env,ELECTRON_RUN_AS_NODE:'1'},stdio:'ignore',windowsHide:true});
    watchdog.on('error',error=>record('watchdog-error',error));watchdog.unref();
  }catch(error){record('watchdog-error',error);}
  app.on('will-quit',()=>{clearInterval(interval);watchdog?.kill();journal.clean();});
  let startup;
  ipcMain.handle('crash:startup',()=>{
    if(!startup)startup=(async()=>{
      let config;
      try{const paths=require('./storage').getPaths();config=require('./assessment/integrationConfig').readBrevoConfig(paths.integrationEnvPath,paths.brevoEnvPath);}catch{}
      const send=config?async({id,log})=>{
        const mailer=require('./assessment/smtpMailer').createSmtpMailer(config);
        const sender=require('./assessment/senderIdentity');
        await mailer.send({from:{name:'Repetierer Diagnose',address:sender.fromEmail},to:'support@akzuwo.ch',subject:`Repetierer Fehlerbericht ${id}`,text:'Automatischer Diagnosebericht nach einem Absturz oder Hänger. Das bereinigte Protokoll ist angehängt.',attachments:[{filename:`repetierer-${id}.log`,content:Buffer.from([config.password,config.user].filter(Boolean).reduce((text,secret)=>text.split(secret).join('[Zugangsdaten]'),log),'utf8'),contentType:'text/plain'}]});
      }:null;
      return journal.deliver(send);
    })();
    return startup;
  });
  ipcMain.handle('crash:open-logs',()=>require('electron').shell.openPath(journal.dir));
  function attach(window){
    window.webContents.on('did-finish-load',()=>window.webContents.send('crash:ready'));
    let prompting=false,hangTimer;
    async function recover(message){
      if(prompting||window.isDestroyed())return;prompting=true;
      try{
        const result=await dialog.showMessageBox(window,{type:'warning',title:'Repetierer reagiert nicht',message,detail:'Bereits gespeicherte Daten und die fortsetzbare Session bleiben erhalten. Ungespeicherte Eingaben können beim Neuladen verloren gehen.',buttons:['Weiter warten','Ansicht neu laden'],defaultId:0,cancelId:0,noLink:true});
        if(result.response===1&&!window.isDestroyed())window.webContents.reload();
      }catch(error){record('recovery-dialog',error);}finally{prompting=false;}
    }
    window.on('unresponsive',()=>{clearTimeout(hangTimer);hangTimer=setTimeout(()=>{record('renderer-unresponsive','Die Oberfläche antwortet nicht.',true);recover('Die Oberfläche ist blockiert.');},10000);});
    window.on('responsive',()=>clearTimeout(hangTimer));
    window.on('closed',()=>clearTimeout(hangTimer));
    window.webContents.on('render-process-gone',(_event,details)=>{if(details.reason==='clean-exit')return;record('renderer-crashed',`Grund: ${details.reason}; Exit-Code: ${details.exitCode}`,true);recover('Die Oberfläche ist abgestürzt. Du kannst die Ansicht neu laden.');});
    window.webContents.on('did-fail-load',(_event,code,description,_url,isMain)=>{if(isMain&&code!==-3)record('load-failed',`${code}: ${description}`,true);});
    window.webContents.on('console-message',event=>{if(event.level==='error')record('renderer-console',event.message);});
  }
  return {journal,attach};
}
module.exports={installCrashHandler};
