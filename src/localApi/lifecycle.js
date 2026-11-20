const fs=require('node:fs'),path=require('node:path');
const {spawn}=require('node:child_process');
const STARTUP_NAME='Repetierer API.vbs';
function startupScript(exe,args) {
  const quote=v=>'"'+String(v).replace(/"/g,'""')+'"';
  const command=[exe,...args].map(v=>'"'+String(v).replace(/"/g,'')+'"').join(' ');
  return `Set shell = CreateObject("WScript.Shell")\r\nshell.Run ${quote(command)}, 0, False\r\n`;
}
function reconcile({app,profile,enabled,startupDirectory,launch=true}) {
  const args=[...(!app.isPackaged?[app.getAppPath()]:[]),'--background-api','--api-profile',profile];
  const startup=startupDirectory||path.join(app.getPath('appData'),'Microsoft','Windows','Start Menu','Programs','Startup');
  const file=path.join(startup,STARTUP_NAME);
  // Development/test profiles never register a real Windows autostart entry.
  if(app.isPackaged || startupDirectory){
    if(enabled){fs.mkdirSync(startup,{recursive:true});fs.writeFileSync(file,startupScript(process.execPath,args),'utf8');}
    else if(fs.existsSync(file))fs.unlinkSync(file);
  }
  if(enabled && launch){
    const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
    const child=spawn(process.execPath,args,{detached:true,windowsHide:true,stdio:'ignore',env});child.on('error',()=>{});child.unref();
  }
  return {enabled,autostart:enabled&&(app.isPackaged||!!startupDirectory)};
}
async function runBackground(app,argv=process.argv) {
  const index=argv.indexOf('--api-profile');
  const profile=index>=0?path.resolve(argv[index+1]):app.getPath('userData');
  // Do not acquire the desktop app's single-instance lock or import its writers.
  await app.whenReady();
  try {
    let settings={};try{settings=JSON.parse(fs.readFileSync(path.join(profile,'settings.json'),'utf8'));}catch(error){if(error.code!=='ENOENT')throw error;}
    if(argv.includes('--api-register-startup'))reconcile({app,profile,enabled:settings.localApiEnabled!==false,launch:false});
    if(settings.localApiEnabled===false){app.quit();return;}
    await require('./server').startServer({profile,onClose:()=>app.quit()});
  }catch(error){if(error.code!=='EADDRINUSE')console.error('Lokaler API-Dienst konnte nicht starten.');app.exit(error.code==='EADDRINUSE'?0:1);}
}
module.exports={reconcile,runBackground,startupScript,STARTUP_NAME};
