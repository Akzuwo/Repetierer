const {ipcRenderer}=require('electron');
const motion=require('./uiMotion');
const select=document.getElementById('settings-class-select'),hint=document.getElementById('settings-class-hint'),rules=document.getElementById('settings-class-rules');
let context={className:'',filePath:''},ready=false,nextRequest=0;
const fields={
 'visual-effects-setting':'visualEffectsEnabled','notification-duration-setting':'notificationDurationSeconds','local-api-setting':'localApiEnabled','smtp-mail-limit-setting':'localMailDailyLimit',
 'homework-enabled-setting':'homeworkEnabled','homework-joker-setting':'homeworkJokerThreshold','homework-weight-setting':'homeworkWeight','extra-joker-setting':'extraJokerAfterThreeGrades',
 'probability-factor-setting':'probabilityDecreaseFactor','boost-never-selected-setting':'boostNeverSelected','never-selected-boost-factor-setting':'neverSelectedBoostFactor','wiggers-rule-setting':'wiggersRuleEnabled','wiggers-rule-duration-setting':'wiggersRuleDurationMinutes'
};
const pending=new Map(),confirmed=new Map();
const scopeKey=(scope,key)=>JSON.stringify([scope.filePath,scope.className,key]);
window.getSettingsScopeContext=()=>context;
function tab(name,focus=false) {
 const panel=document.getElementById(`settings-${name}-fields`);
 if((panel.classList.contains('update-hidden') || document.getElementById('content').dataset.pageMotion) && motion.beginPageChange(()=>tab(name,focus)))return;
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
 context={filePath:value?.filePath||'',className:value?.className||''};ready=true;
 for(const key of Object.values(fields))confirmed.set(scopeKey(context,key),_settings[key]);
 select.replaceChildren();
 const placeholder=document.createElement('option');placeholder.value='';placeholder.textContent='Klasse auswählen';select.append(placeholder);
 for(const name of value?.classes||[]){const option=document.createElement('option');option.value=option.textContent=name;select.append(option);}
 select.value=context.className;select.disabled=!(value?.classes?.length);rules.disabled=!context.className;
 hint.textContent=context.className?'Diese Regeln gelten nur für diese Klasse und werden lokal in der App gespeichert.':value?.classes?.length?'Bitte oben eine Klasse auswählen.':'Bitte zuerst eine Exceldatei mit Klassen laden. Die globalen Einstellungen bleiben verfügbar.';
 select.dispatchEvent(new Event('change',{bubbles:false}));
 window.repetiererKristall?.refresh();
});
select.addEventListener('change',event=>{
 if(!event.isTrusted && select.value===context.className)return;
 ready=false;
 ipcRenderer.send('get-settings',{className:select.value});
});
document.getElementById('settings-view').addEventListener('input',event=>{
 const input=event.target,key=fields[input.id];
 if(!ready||!key||input.disabled||input.closest('fieldset:disabled'))return;
 if(input.closest('#settings-class-fields')&&!context.className)return;
 const value=input.type==='checkbox'?input.checked:input.valueAsNumber;
 if(typeof value==='number'&&(!Number.isFinite(value)||!input.checkValidity()))return;
 const requestId=++nextRequest,scope={...context};
 pending.set(requestId,{input,key,value,scope});
 ipcRenderer.send('save-settings',{scope:'combined',...scope,live:true,requestId,[key]:value});
});
ipcRenderer.on('settings-saved',(_event,_settings,details)=>{
 const saved=pending.get(details?.requestId);if(!saved)return;
 confirmed.set(scopeKey(saved.scope,saved.key),saved.value);pending.delete(details.requestId);
});
ipcRenderer.on('settings-save-failed',(_event,_message,details)=>{
 const failed=pending.get(details?.requestId);if(!failed)return;pending.delete(details.requestId);
 if(failed.scope.filePath!==context.filePath||failed.scope.className!==context.className)return;
 if([...pending.values()].some(p=>p.input===failed.input))return;
 const value=confirmed.get(scopeKey(failed.scope,failed.key));
 if(failed.input.type==='checkbox')failed.input.checked=!!value;else failed.input.value=value;
 // Refresh dependent labels and disabled states without creating another write.
 failed.input.dispatchEvent(new Event('change',{bubbles:true}));
 window.repetiererKristall?.refresh();
});
