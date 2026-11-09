const {ipcRenderer} = require('electron');
const motion = require('./uiMotion');
const notify = require('./notifications');
const actions = [
  {id:'notification-info',type:'info',label:'Nachricht',text:'Dies ist eine Testnachricht.'},
  {id:'notification-warning',type:'warning',label:'Warnung',text:'Dies ist eine Testwarnung.'},
  {id:'notification-error',type:'error',label:'Fehler',text:'Dies ist eine Testfehlermeldung.'}
];
let installed=false, launcher, panel;
function closePanel(afterClose) {
  motion.closeDialog(panel,()=>{launcher.focus();afterClose?.();});
}
function openPanel() {
  motion.setDrawerOpen(false);
  launcher.setAttribute('aria-expanded','true');
  motion.openDialog(panel);
  window.repetiererKristall?.refresh();
  panel.querySelector('button').focus();
}
function createPanel() {
  if(panel)return;
  launcher=document.createElement('button');launcher.id='debug-launcher';launcher.className='btn-2 debug-launcher';
  launcher.textContent='Debug';launcher.setAttribute('aria-haspopup','dialog');launcher.setAttribute('aria-expanded','false');
  launcher.addEventListener('click',openPanel);document.body.append(launcher);
  panel=document.createElement('div');panel.id='debug-panel';panel.className='modal';
  panel.setAttribute('role','dialog');panel.setAttribute('aria-modal','true');panel.setAttribute('aria-labelledby','debug-heading');
  panel.innerHTML=`<div class="modal-content debug-panel-content">
    <h2 id="debug-heading">Debug</h2>
    <section aria-labelledby="debug-messages-heading"><h3 id="debug-messages-heading">Meldungen</h3><div class="debug-message-actions"></div></section>
    <section aria-labelledby="debug-updates-heading"><h3 id="debug-updates-heading">Updates</h3><div class="debug-update-actions">
      <button type="button" class="btn-2" data-debug-scenario="update-preview">Fake Update</button>
      <button type="button" class="btn-2" data-debug-scenario="update-restart">Update mit Neustart</button>
    </div></section>
    <form id="debug-mail-form"><h3>Testmail</h3><p>Eine echte E-Mail mit Beispieldaten im normalen Mitmachnoten-Layout. Verwendet die eingerichtete SMTP-Verbindung.</p>
      <label for="debug-mail-recipient">Empfängeradresse</label>
      <input id="debug-mail-recipient" type="email" required autocomplete="email" placeholder="name@beispiel.ch">
      <button type="submit" class="btn-2" id="debug-mail-send">Testmail senden</button>
    </form>
  </div>`;
  for(const mode of ['preview','restart']) panel.querySelector(`[data-debug-scenario="update-${mode}"]`).addEventListener('click',()=>{
    launcher.setAttribute('aria-expanded','false');
    closePanel(()=>ipcRenderer.send('debug-update-show',mode));
  });
  panel.querySelector('#debug-mail-form').addEventListener('submit',async event=>{
    event.preventDefault();
    const button=panel.querySelector('#debug-mail-send');
    if(button.disabled)return;
    button.disabled=true;button.textContent='Wird gesendet …';
    try {
      const result=await ipcRenderer.invoke('debug:send-test-mail',{recipient:panel.querySelector('#debug-mail-recipient').value});
      if(!result.ok)throw new Error(result.error.message);
      notify.success('Die Testmail wurde versendet.',{title:'Testmail'});
    } catch(error) {notify.error(error.message,{title:'Testmail nicht gesendet'});}
    finally {button.disabled=false;button.textContent='Testmail senden';}
  });
  for(const action of actions) {
    const button=document.createElement('button');button.type='button';button.className='btn-2';button.textContent=action.label;
    button.dataset.debugScenario=action.id;
    button.addEventListener('click',()=>{
      launcher.setAttribute('aria-expanded','false');
      closePanel(()=>notify[action.type](action.text,{title:action.label}));
    });
    panel.querySelector('.debug-message-actions').append(button);
  }
  document.body.append(panel);
  window.repetiererKristall?.refresh();
  panel.addEventListener('click',event=>{if(event.target===panel){launcher.setAttribute('aria-expanded','false');closePanel();}},true);
  document.getElementById('debug-scenarios-btn').addEventListener('click',openPanel);
  document.getElementById('debug-actions').classList.remove('update-hidden');
  document.addEventListener('keydown',event=>{
    if(event.ctrlKey && event.shiftKey && event.key.toLowerCase()==='d'){event.preventDefault();openPanel();}
    if(!panel.classList.contains('modal-open'))return;
    if(event.key==='Escape'){event.preventDefault();event.stopImmediatePropagation();launcher.setAttribute('aria-expanded','false');closePanel();}
    if(event.key==='Tab') {
      const buttons=[...panel.querySelectorAll('button:not(:disabled),input:not(:disabled)')];
      if(event.shiftKey && document.activeElement===buttons[0]){event.preventDefault();buttons.at(-1).focus();}
      else if(!event.shiftKey && document.activeElement===buttons.at(-1)){event.preventDefault();buttons[0].focus();}
    }
  },true);
}
function install() {
  if(installed)return;installed=true;
  // Retire saved previews from the former catalogue. Message buttons never
  // reload the document, seed demo data or replace the app's IPC methods.
  sessionStorage.removeItem('repetierer-debug-scenario');
  ipcRenderer.on('debug-mode',(_event,enabled)=>{
    if(!enabled)return;
    if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',createPanel,{once:true});
    else createPanel();
  });
}
module.exports={install};
