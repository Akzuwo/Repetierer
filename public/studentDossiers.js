const {ipcRenderer}=require('electron');
const motion=require('./uiMotion'),notify=require('./notifications');
const {dossierHtml}=require('../src/studentDossierView');
const page=document.getElementById('dossiers-view'),body=document.getElementById('dossiers-body'),heading=document.getElementById('dossiers-heading'),back=document.getElementById('dossiers-back');
const exportButton=document.getElementById('dossiers-export-btn');
let data=null,className='',studentId=null,request=0,exporting=false,optionMode='export';
const exportModal=document.createElement('div');
exportModal.className='modal';exportModal.id='dossiers-export-modal';
exportModal.setAttribute('role','dialog');exportModal.setAttribute('aria-modal','true');exportModal.setAttribute('aria-labelledby','dossiers-export-heading');exportModal.setAttribute('aria-hidden','true');
exportModal.innerHTML='<div class="modal-content dossier-export-content"><h2 id="dossiers-export-heading">Schülerakten · Optionen</h2><p id="dossiers-options-description">Was möchtest du als PDF exportieren?</p><div id="dossiers-option-modes" class="dossier-option-modes"></div><div id="dossiers-export-options"></div><div class="modal-buttons"><button type="button" class="btn-2" id="dossiers-export-cancel">Abbrechen</button></div></div>';
document.body.append(exportModal);
function closeExport(afterClose){motion.closeDialog(exportModal,()=>{if(!page.classList.contains('update-hidden')&&!exportButton.disabled)exportButton.focus({preventScroll:true});afterClose?.();});}
function cancelDisclosures(){body.querySelectorAll('.dossier-assessment').forEach(details=>motion.cancel(details));}
function dismiss(){page.classList.add('update-hidden');document.getElementById('dossiers-nav-btn').removeAttribute('aria-current');request++;cancelDisclosures();closeExport();}
window.dismissStudentDossiers=dismiss;
window.isStudentDossierOpen=()=>!page.classList.contains('update-hidden');
function button(text,action,className='btn-2'){const b=document.createElement('button');b.type='button';b.className=className;b.textContent=text;b.onclick=action;return b;}
function change(fn){motion.pageChange(()=>{fn();body.scrollTop=0;window.repetiererKristall?.refresh();});}
function open(){
  if(window.requestRepetiererExcelEditorNavigation)window.requestRepetiererExcelEditorNavigation(openPage);
  else openPage();
}
async function openPage(){
  if(motion.beginPageChange(openPage))return;
  window.dismissRepetiererSettingsForNavigation?.();window.dismissRepetiererExcelEditorForNavigation?.();
  window.captureAssessmentDraft?.();
  const content=document.getElementById('content');
  for(const element of content.children)element.classList.add('update-hidden');
  content.classList.remove('session-active','settings-active');
  document.getElementById('session-header').classList.add('update-hidden');
  document.querySelectorAll('#drawer [aria-current]').forEach(e=>e.removeAttribute('aria-current'));
  document.getElementById('dossiers-nav-btn').setAttribute('aria-current','page');
  page.classList.remove('update-hidden');motion.setDrawerOpen(false);
  closeExport();cancelDisclosures();data=null;exportButton.disabled=true;
  className='';studentId=null;heading.textContent='Schülerakten';back.hidden=true;
  body.textContent='Schülerakten werden geladen …';body.setAttribute('aria-busy','true');
  const id=++request;
  try{const result=await ipcRenderer.invoke('dossiers:get');if(id!==request)return;if(!result.ok)throw new Error(result.error);data=result.data;render();}
  catch(error){if(id!==request)return;body.replaceChildren();const p=document.createElement('p');p.textContent=error.message;body.append(p,button('Erneut laden',open));notify.error(error.message,{title:'Schülerakten'});}
  finally{if(id===request)body.setAttribute('aria-busy','false');}
}
function render(){
  cancelDisclosures();body.replaceChildren();back.hidden=!className;
  const cls=data.classes.find(c=>c.name===className),student=cls?.students.find(s=>s.id===studentId);
  heading.textContent=student?student.name:cls?`Schülerakten · ${className}`:'Schülerakten';
  exportButton.disabled=exporting||!data.classes.some(c=>c.students.length);
  if(student){const details=document.createElement('div');details.className='dossier-record';details.innerHTML=dossierHtml(student);body.append(contactEditor(student),details);details.querySelectorAll('.dossier-assessment').forEach(motion.bindDisclosure);}
  else{
    const hint=document.createElement('p');hint.className='dossier-selection-hint';hint.textContent=cls?'Wähle eine Person, um die Akte anzusehen.':'Wähle eine Klasse aus der geladenen Datei.';body.append(hint);
    const list=document.createElement('div');list.className=cls?'dossier-students':'dossier-classes';
    for(const item of cls?cls.students:data.classes){const b=button(item.name,()=>change(()=>{if(cls)studentId=item.id;else className=item.name;render();}),cls?'participation-student':'btn-2');if(cls)b.dataset.dossierStudent=item.id;else b.dataset.dossierClass=item.name;list.append(b);}
    if(!list.children.length){const empty=document.createElement('p');empty.textContent=cls?'Diese Klasse enthält keine Schüler.':'Die Datei enthält keine Repetierer-Klassen.';list.append(empty);}
    body.append(list);
  }
  if(student&&data.legacySessionCount){const note=document.createElement('p');note.className='dossier-footnote';note.textContent=`${data.legacySessionCount} ältere Session(s) ohne Dateizuordnung sind in den Sessionstatistiken nicht enthalten.`;body.append(note);}
}
function context(scope){return {scope,className,studentId,filePath:data.filePath,fingerprint:data.fingerprint,studentName:data.classes.find(c=>c.name===className)?.students.find(s=>s.id===studentId)?.name};}
function contactEditor(student){
  const section=document.createElement('section');section.className='dossier-contact dossier-section';
  const label=document.createElement('h3');label.textContent='E-Mail-Adresse';
  const current=document.createElement('p');current.id='dossier-current-email';current.textContent=student.email||'Noch keine E-Mail-Adresse erfasst.';
  const hint=document.createElement('p');hint.id='dossier-email-hint';hint.textContent='Eine spätere neue Schülerangabe aus der Selbstbeurteilung ersetzt auch eine manuell bearbeitete Adresse.';
  const edit=button('Adresse bearbeiten',()=>{form.hidden=false;edit.hidden=true;input.focus();input.select();});edit.id='dossier-email-edit';
  const form=document.createElement('form');form.hidden=true;form.className='dossier-email-form';
  const fieldLabel=document.createElement('label');fieldLabel.htmlFor='dossier-email-input';fieldLabel.textContent='Aktuelle E-Mail-Adresse';
  const input=document.createElement('input');input.id='dossier-email-input';input.type='email';input.required=true;input.maxLength=254;input.autocomplete='off';input.value=student.email;input.setAttribute('aria-describedby',hint.id);
  const actions=document.createElement('div');actions.className='dossier-contact-actions';
  const save=button('Speichern',null);save.id='dossier-email-save';save.type='submit';
  const cancel=button('Abbrechen',()=>{input.value=student.email;form.hidden=true;edit.hidden=false;edit.focus();});
  actions.append(cancel,save);form.append(fieldLabel,input,actions);
  form.onsubmit=async event=>{
    event.preventDefault();if(exporting||!form.reportValidity())return;
    const args={...context('student'),email:input.value,expectedEmail:student.email},id=request;setBusy(true);
    try{const result=await ipcRenderer.invoke('dossiers:set-email',args);if(!result.ok)throw new Error(result.error);notify.success('E-Mail-Adresse gespeichert.',{title:'Schülerakte'});if(id===request){student.email=result.email;current.textContent=result.email;form.hidden=true;edit.hidden=false;edit.focus();}}
    catch(error){notify.error(error.message,{title:'E-Mail-Adresse'});}
    finally{setBusy(false);if(id===request&&form.hidden)edit.focus({preventScroll:true});}
  };
  section.append(label,current,hint,edit,form);return section;
}
function setBusy(value){
  exporting=value;page.setAttribute('aria-busy',String(value));
  body.querySelectorAll('button,input').forEach(b=>b.disabled=value);back.disabled=value;exportButton.disabled=value||!data;
  if(!value&&!data&&!page.classList.contains('update-hidden'))open();
}
function showExport(){
  if(exporting||!data||exportButton.disabled)return;
  optionMode='export';showOptions();
  motion.openDialog(exportModal);window.repetiererKristall?.refresh();exportModal.querySelector('button:not(:disabled)').focus({preventScroll:true});
}
function showOptions(){
  const options=document.getElementById('dossiers-export-options'),modes=document.getElementById('dossiers-option-modes'),cls=data.classes.find(c=>c.name===className),student=cls?.students.find(s=>s.id===studentId);
  document.getElementById('dossiers-export-heading').textContent='Schülerakten · Optionen';document.getElementById('dossiers-export-cancel').textContent='Abbrechen';
  document.getElementById('dossiers-options-description').textContent=optionMode==='export'?'Was möchtest du als PDF exportieren?':'Jede Person erhält ausschliesslich ihre eigene Akte als PDF über Brevo.';
  options.replaceChildren();modes.replaceChildren();
  for(const [mode,label] of [['export','PDF exportieren'],['mail','Per E-Mail senden']]){const b=button(label,()=>{optionMode=mode;showOptions();modes.querySelector(`[data-dossier-mode="${mode}"]`).focus();});b.dataset.dossierMode=mode;b.setAttribute('aria-pressed',String(mode===optionMode));modes.append(b);}
  for(const [scope,label,detail] of [['all','Alle Klassen',optionMode==='export'?'Je ein Ordner pro Klasse':'Alle Schüler der geladenen Datei'],...(cls?.students.length?[['class','Ganze Klasse',cls.name]]:[]),...(student?[['student','Nur Schüler',student.name]]:[])]){
    const choice=button('',()=>{const args=context(scope);closeExport(()=>optionMode==='export'?exportPdf(scope):prepareMail(args));},'dossier-export-choice');choice.dataset[optionMode==='export'?'dossierExport':'dossierMail']=scope;
    const title=document.createElement('strong'),description=document.createElement('span');title.textContent=label;description.textContent=detail;choice.append(title,description);options.append(choice);
  }
  window.repetiererKristall?.refresh();
}
function showMailPanel(title,description){
  document.getElementById('dossiers-export-heading').textContent=title;document.getElementById('dossiers-options-description').textContent=description;document.getElementById('dossiers-export-cancel').textContent='Abbrechen';
  document.getElementById('dossiers-option-modes').replaceChildren();const options=document.getElementById('dossiers-export-options');options.replaceChildren();return options;
}
function openMailPanel(){motion.openDialog(exportModal);window.repetiererKristall?.refresh();document.getElementById('dossiers-export-cancel').focus({preventScroll:true});}
async function prepareMail(args){
  if(exporting||!data||page.classList.contains('update-hidden'))return;
  const id=request;setBusy(true);
  try{
    const preview=await ipcRenderer.invoke('dossiers:prepare-mail',args);if(!preview.ok)throw new Error(preview.error);if(id!==request)return;
    const options=showMailPanel('Versand prüfen',preview.missing.length?`Versand gesperrt: Bei ${preview.missing.length} Person(en) fehlt eine gültige E-Mail-Adresse. Bitte die Adressen in den Akten ergänzen.`:`${preview.recipients.length} persönliche PDF(s) an die folgenden Empfänger senden?`);
    const list=document.createElement('ul');list.className='dossier-mail-recipients';list.tabIndex=0;
    for(const person of preview.recipients){const li=document.createElement('li'),name=document.createElement('strong'),email=document.createElement('span');name.textContent=`${person.name} · ${person.className}`;email.textContent=person.email||'E-Mail-Adresse fehlt';li.append(name,email);list.append(li);}options.append(list);
    if(preview.token){const send=button('Jetzt senden',()=>{send.disabled=true;closeExport(()=>sendMail(args,preview.token));});send.id='dossiers-mail-confirm';options.append(send);}
    openMailPanel();
  }catch(error){notify.error(error.message,{title:'Schülerakten versenden'});}
  finally{setBusy(false);}
}
async function sendMail(args,token){
  const id=request;setBusy(true);exportButton.lastChild.textContent='Wird gesendet …';notify.info('Schülerakten werden versendet …',{title:'Versand',key:'dossiers-delivery'});
  try{
    const result=await ipcRenderer.invoke('dossiers:send-mail',{token});if(!result.ok)throw new Error(result.error);
    const message=`${result.sent} Schülerakte(n) versendet. ${result.remaining} nicht gesendet.`;
    if(result.remaining||result.error)notify.warning(message+(result.error?' '+result.error:''),{title:'Versand',key:'dossiers-delivery'});else notify.success(message,{title:'Versand',key:'dossiers-delivery'});
    if(id!==request)return;
    const options=showMailPanel('Versandergebnis',message+(result.error?' '+result.error:''));document.getElementById('dossiers-export-cancel').textContent='Schliessen';
    const list=document.createElement('ul');list.className='dossier-mail-recipients';list.tabIndex=0;
    for(const entry of result.results){const li=document.createElement('li');li.textContent=`${entry.name} · ${entry.className}: ${entry.status==='sent'?'Versendet':entry.error}`;list.append(li);}options.append(list);
    if(result.retryToken){const retry=button('Nicht gesendete erneut prüfen',()=>closeExport(()=>prepareMail({...args,retryToken:result.retryToken})));retry.id='dossiers-mail-retry';options.append(retry);}
    openMailPanel();
  }catch(error){notify.error(error.message,{title:'Schülerakten versenden'});}
  finally{exportButton.lastChild.textContent='Optionen';setBusy(false);}
}
exportButton.onclick=showExport;
document.getElementById('dossiers-export-cancel').onclick=()=>closeExport();
exportModal.addEventListener('click',event=>{if(event.target===exportModal)closeExport();});
exportModal.addEventListener('keydown',event=>{
  if(event.key==='Escape'){event.preventDefault();event.stopImmediatePropagation();closeExport();}
  if(event.key==='Tab'){
    const buttons=[...exportModal.querySelectorAll('button:not(:disabled),[tabindex="0"]')],first=buttons[0],last=buttons.at(-1);
    if(event.shiftKey&&document.activeElement===first){event.preventDefault();last.focus();}
    else if(!event.shiftKey&&document.activeElement===last){event.preventDefault();first.focus();}
  }
});
async function exportPdf(scope){
  if(exporting||!data||page.classList.contains('update-hidden'))return;exporting=true;page.setAttribute('aria-busy','true');body.querySelectorAll('button').forEach(b=>b.disabled=true);back.disabled=exportButton.disabled=true;
  try{const result=await ipcRenderer.invoke('dossiers:export',{scope,className,studentId,filePath:data.filePath,fingerprint:data.fingerprint,studentName:data.classes.find(c=>c.name===className)?.students.find(s=>s.id===studentId)?.name});if(!result.ok)throw new Error(`${result.error}${result.count?` ${result.count} Dokument(e) wurden bereits gespeichert.`:''}`);if(!result.cancelled)notify.success(`${result.count} Schülerakte(n) als PDF exportiert.`,{title:'Export'});}
  catch(error){notify.error(error.message,{title:'Export'});}
  finally{exporting=false;page.setAttribute('aria-busy','false');body.querySelectorAll('button').forEach(b=>b.disabled=false);back.disabled=false;exportButton.disabled=!data;if(!data&&!page.classList.contains('update-hidden'))open();else if(!page.classList.contains('update-hidden'))exportButton.focus({preventScroll:true});}
}
back.onclick=()=>change(()=>{if(studentId!==null)studentId=null;else className='';render();});
document.getElementById('dossiers-nav-btn').onclick=open;
ipcRenderer.on('classes',()=>{if(!page.classList.contains('update-hidden')){request++;data=null;exportButton.disabled=true;closeExport();if(!exporting)open();}});

ipcRenderer.on('assessment:auto-synced',async()=>{if(!data||exporting||page.classList.contains('update-hidden')||body.querySelector('.dossier-email-form:not([hidden])'))return;const id=++request;const result=await ipcRenderer.invoke('dossiers:get');if(id===request&&result.ok){data=result.data;render();}});
