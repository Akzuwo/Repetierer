const fs=require('node:fs/promises');
const path=require('node:path');
const API_DOCUMENT_FILES=['docs/local-api.md','docs/assessment-api.md','docs/mail-service.md'];
async function apiMarkdown(appPath,version){
  const sections=await Promise.all(API_DOCUMENT_FILES.map(async file=>{
    const source=await fs.readFile(path.join(appPath,file),'utf8');
    // The in-app reference contains the usage guide, not private operator/test instructions.
    const content=source.split(/^## (?:Entwicklung und Prüfung|Betrieb und Tests)\s*$/m)[0];
    return content.trim().replace(/^(#{1,5}) /gm,'#$1 ');
  }));
  return `# Repetierer – API-Dokumentation\n\nApp-Version: ${version}\n\nDiese Referenz beschreibt die lokale Lese-API, die Cloud-Umfrage und den Cloudmail-Versand. Beispiele enthalten keine persönlichen Zugangsdaten.\n\n`+sections.join('\n\n---\n\n')+'\n';
}
function registerDocumentation({app,ipcMain,dialog,BrowserWindow}){
  let exporting=false;
  ipcMain.handle('documentation:export',async(event,args={})=>{
    if(args?.id==='manual')return {ok:false,error:'Die Bedienungsanleitung ist noch nicht verfügbar.'};
    if(args?.id!=='api')return {ok:false,error:'Unbekannte Dokumentation.'};
    if(exporting)return {ok:false,error:'Ein Dokumentationsexport läuft bereits.'};
    exporting=true;
    try{
      const content=await apiMarkdown(app.getAppPath(),app.getVersion());
      const result=await dialog.showSaveDialog(BrowserWindow.fromWebContents(event.sender),{
        title:'API-Dokumentation als Markdown speichern',defaultPath:'Repetierer-API-Dokumentation.md',filters:[{name:'Markdown',extensions:['md']}]
      });
      if(result.canceled||!result.filePath)return {ok:true,cancelled:true};
      const filePath=/\.md$/i.test(result.filePath)?result.filePath:result.filePath+'.md';
      await fs.writeFile(filePath,content,'utf8');
      return {ok:true,filePath};
    }catch{return {ok:false,error:'Die API-Dokumentation konnte nicht gespeichert werden. Bitte den Zielordner und die Schreibrechte prüfen.'};}
    finally{exporting=false;}
  });
}
module.exports={API_DOCUMENT_FILES,apiMarkdown,registerDocumentation};
