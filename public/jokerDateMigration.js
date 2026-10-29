const { ipcRenderer } = require('electron');
const motion = require('./uiMotion');
const modal = document.getElementById('joker-date-migration-modal');
const summary = document.getElementById('joker-date-migration-summary');
const preview = document.getElementById('joker-date-migration-preview');
const status = document.getElementById('joker-date-migration-status');
const run = document.getElementById('joker-date-migration-run');
const skip = document.getElementById('joker-date-migration-skip');
let fingerprint, busy = false;
function setBusy(value) {
  busy=value;run.disabled=value || !fingerprint;skip.disabled=value;
  run.textContent=value?'Datei wird erweitert …':'Sicherung erstellen und erweitern';
  modal.setAttribute('aria-busy',String(value));
}
ipcRenderer.on('joker-date-migration-required', (_event, plan, filePath) => {
  fingerprint=plan.error?null:plan.fingerprint;
  summary.textContent=`${filePath || 'Die geladene Excel-Datei'} verwendet noch die bisherige Joker-Datumsspalte.`;
  preview.replaceChildren();
  for(const sheet of (plan.sheets || []).filter(sheet=>sheet.needsMigration)) {
    const item=document.createElement('li');
    item.textContent=`${sheet.name}: ${sheet.columns} Datumsspalten · ${sheet.dates} gespeicherte Joker-Daten${sheet.keepsLegacyColumn?' · eigene Spalten bleiben erhalten; die alte Sammelspalte bleibt als Archiv':''}`;
    preview.append(item);
  }
  status.textContent=plan.error || '';
  status.classList.toggle('update-hidden',!plan.error);
  setBusy(false);motion.openDialog(modal);run.disabled?skip.focus():run.focus();
});
run.addEventListener('click',()=>{if(busy || !fingerprint)return;status.textContent='Sicherungskopie wird erstellt. Bitte die Excel-Datei geschlossen lassen.';status.classList.remove('update-hidden');setBusy(true);ipcRenderer.send('run-joker-date-migration',fingerprint);});
skip.addEventListener('click',()=>{if(busy)return;motion.closeDialog(modal);ipcRenderer.send('skip-joker-date-migration');});
modal.addEventListener('keydown',event=>{if(event.key==='Escape'){event.preventDefault();event.stopImmediatePropagation();if(!busy)skip.click();}});
ipcRenderer.on('joker-date-migration-result',(_event,result)=>{
  setBusy(false);
  if(result.success) {
    motion.closeDialog(modal);
    const notice=document.getElementById('joker-date-migration-notice');
    notice.textContent=`Joker-Datumsspalten ergänzt. ${result.dates || 0} gespeicherte Daten übernommen.${result.backupPath?' Sicherung: '+result.backupPath:''}`;
    notice.classList.remove('update-hidden');
    setTimeout(()=>notice.classList.add('update-hidden'),15000);
    return;
  }
  const reason=result.reason==='excel-locked'?'Die Excel-Datei ist geöffnet oder gesperrt. Bitte schliessen und erneut versuchen.':result.reason==='file-changed'?'Die Datei wurde seit der Vorschau verändert. Bitte die Datei erneut öffnen und die neue Vorschau prüfen.':`Die Erweiterung konnte nicht abgeschlossen werden. ${result.error || 'Bitte Schreibrechte und freien Speicher prüfen.'}`;
  if(result.reason==='file-changed') {fingerprint=null;run.disabled=true;}
  status.textContent=reason+(result.backupPath?' Sicherung vorhanden: '+result.backupPath:'');status.classList.remove('update-hidden');
});
