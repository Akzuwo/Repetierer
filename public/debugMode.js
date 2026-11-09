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
  panel.innerHTML='<div class="modal-content debug-panel-content"><h2 id="debug-heading">Debug</h2><p>Testmeldung anzeigen</p><div class="debug-message-actions"></div></div>';
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
      const buttons=[...panel.querySelectorAll('button')];
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
