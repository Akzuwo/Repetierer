const {ipcRenderer}=require('electron');
const report=error=>{try{ipcRenderer.send('crash:renderer-error',String(error?.stack||error?.message||error||'Unbekannter Fehler'));}catch{}};
window.addEventListener('error',event=>report(event.error||`${event.message} (${event.filename}:${event.lineno})`));
window.addEventListener('unhandledrejection',event=>report(event.reason));
const start=async()=>{
  try{
    const result=await ipcRenderer.invoke('crash:startup');
    if(result.crashed){
      const notify=require('./notifications');
      const status=result.failed?'Der automatische Versand ist fehlgeschlagen; das Protokoll bleibt gespeichert.':result.configured?'Das Crashprotokoll wurde automatisch an support@akzuwo.ch gesendet.':'Das Crashprotokoll wurde lokal gespeichert. Für den automatischen Versand ist Brevo erforderlich.';
      notify.warning(`Die App wurde zuletzt unerwartet beendet oder hat einen Fehler/Hänger erfasst. ${status}`,{title:'Vorheriger Appfehler',key:'previous-crash'});
    }
  }catch{/* Minimal test windows may not install the app crash handler. */}
};
ipcRenderer.once('crash:ready',start);
document.getElementById('crash-logs-btn')?.addEventListener('click',async()=>{
  try{const error=await ipcRenderer.invoke('crash:open-logs');if(error)throw new Error(error);}catch(error){require('./notifications').error(error.message,{title:'Crashprotokolle'});}
});
