const {ipcRenderer}=require('electron');
const select=document.getElementById('settings-class-select'),hint=document.getElementById('settings-class-hint'),rules=document.getElementById('settings-class-rules');
let context={className:'',filePath:''},dirty=false;
window.getSettingsScopeContext=()=>context;
function tab(name,focus=false) {
 for(const key of ['global','class']) {
  const button=document.getElementById(`settings-${key}-tab`),active=key===name;
  button.classList.toggle('btn-1',active);button.classList.toggle('btn-2',!active);button.setAttribute('aria-selected',String(active));button.tabIndex=active?0:-1;
  document.getElementById(`settings-${key}-fields`).classList.toggle('update-hidden',!active);
  if(active && focus)button.focus();
 }
 window.repetiererKristall?.refresh();
}
for(const name of ['global','class']) {
 const button=document.getElementById(`settings-${name}-tab`);
 button.addEventListener('click',()=>tab(name));
 button.addEventListener('keydown',event=>{if(['ArrowLeft','ArrowRight','Home','End'].includes(event.key)){event.preventDefault();tab(event.key==='Home'?'global':event.key==='End'?'class':name==='global'?'class':'global',true);}});
}
ipcRenderer.on('settings-data',(_event,_settings,_paths,_migration,value)=>{
 context={filePath:value?.filePath||'',className:value?.className||''};dirty=false;
 select.replaceChildren();
 for(const name of value?.classes||[]){const option=document.createElement('option');option.value=option.textContent=name;select.append(option);}
 select.value=context.className;select.disabled=!context.className;rules.disabled=!context.className;
 hint.textContent=context.className?'Diese Regeln gelten nur für diese Klasse und werden lokal in der App gespeichert.':'Bitte zuerst eine Exceldatei mit Klassen laden. Die globalen Einstellungen bleiben verfügbar.';
 select.dispatchEvent(new Event('change',{bubbles:false}));
 window.repetiererKristall?.refresh();
});
select.addEventListener('change',event=>{
 if(!event.isTrusted && select.value===context.className)return;
 if(dirty)return;
 ipcRenderer.send('get-settings',{className:select.value});
});
document.getElementById('settings-view').addEventListener('input',event=>{
 if(event.target===select)return;
 dirty=true;select.disabled=true;
 if(context.className)hint.textContent='Speichere oder verwirf deine Änderungen, bevor du die Klasse wechselst.';
});
