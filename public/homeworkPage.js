const {ipcRenderer}=require('electron');
const motion=require('./uiMotion'),notify=require('./notifications');
const page=document.getElementById('homework-view'),body=document.getElementById('homework-body');
let settings={homeworkEnabled:true,homeworkJokerThreshold:2,homeworkWeight:1.2},session,people=[],selected=null,busy=false,request=0;
const modal=document.createElement('div');modal.className='modal';modal.id='homework-migration-modal';
modal.setAttribute('role','dialog');modal.setAttribute('aria-modal','true');modal.setAttribute('aria-labelledby','homework-migration-heading');
modal.innerHTML=`<div class="modal-content"><h2 id="homework-migration-heading">Hausaufgaben in Excel ergänzen</h2><p>Rechts neben den vorhandenen Spalten werden Noten, Datum und die gewichtete Auswertung ergänzt. Vorhandene Daten bleiben erhalten. Vor dem Erweitern wird eine Sicherungskopie erstellt.</p><ul id="homework-migration-preview"></ul><div class="modal-buttons"><button type="button" class="btn-2" id="homework-migration-skip">Später</button><button type="button" class="btn-1" id="homework-migration-run">Sicherung erstellen und erweitern</button></div></div>`;
document.body.append(modal);
const run=modal.querySelector('#homework-migration-run'),skip=modal.querySelector('#homework-migration-skip');
let plan,migrating=false,checking=false;
skip.addEventListener('click',()=>{if(!migrating)motion.closeDialog(modal);});
modal.addEventListener('click',event=>{if(migrating){event.stopImmediatePropagation();event.preventDefault();}},true);
modal.addEventListener('keydown',event=>{
 if(event.key==='Escape'){event.preventDefault();event.stopImmediatePropagation();if(!migrating)skip.click();}
 if(event.key==='Tab'){if(event.shiftKey && document.activeElement===skip){event.preventDefault();run.focus();}else if(!event.shiftKey && document.activeElement===run){event.preventDefault();skip.focus();}}
});
async function ensureColumns() {
 if(checking || !settings.homeworkEnabled || !session)return;
 checking=true;
 try {
  plan=await ipcRenderer.invoke('homework:status');
  if(plan.error)throw new Error(plan.error);
  if(!plan.needsMigration)return true;
  const preview=modal.querySelector('ul');preview.replaceChildren();
  for(const sheet of plan.sheets.filter(s=>s.needsMigration)){const li=document.createElement('li');li.textContent=`${sheet.name}: sechs Plätze für Vorführungen und eine gemeinsame gewichtete Auswertung. Weitere Plätze werden bei Bedarf ergänzt.`;preview.append(li);}
  motion.openDialog(modal);run.focus();return false;
 } catch(error){notify.error(error.message,{title:'Hausaufgaben'});return false;}finally{checking=false;}
}
run.addEventListener('click',async()=>{
 if(migrating || !plan?.fingerprint)return;
 migrating=true;run.disabled=skip.disabled=true;modal.setAttribute('aria-busy','true');
 try {
  const result=await ipcRenderer.invoke('homework:prepare',plan.fingerprint);
  if(!result.success)throw new Error(result.reason==='file-changed'?'Die Datei wurde verändert. Bitte neu öffnen und die Vorschau erneut prüfen.':result.reason==='excel-locked'?'Bitte Excel schliessen und erneut versuchen.':result.error||'Die Spalten konnten nicht angelegt werden.');
  motion.closeDialog(modal);notify.success('Die Hausaufgabenspalten und eine Sicherungskopie wurden angelegt.',{title:'Excel erweitert'});await load();
 } catch(error){notify.error(error.message,{title:'Excel nicht erweitert'});}
 finally{migrating=false;run.disabled=skip.disabled=false;modal.setAttribute('aria-busy','false');}
});
window.applyHomeworkSettings=value=>{settings={...settings,...value};if(selected && body.querySelector('#homework-grade-detail'))updateDetail(selected);if(!settings.homeworkEnabled){selected=null;if(!migrating)motion.closeDialog(modal);}};
window.renderHomework=value=>{if(session?.id!==value?.id){session=value;selected=null;people=[];}load();ensureColumns();};
async function load(animate=false) {
 if(!session || page.classList.contains('update-hidden'))return;
 const id=++request,currentSession=session.id;
 try {
  const result=await ipcRenderer.invoke('homework:persons');
  if(id!==request || window.getHomeworkSession()?.id!==currentSession)return;
  if(!result.success)throw new Error(result.error);
  people=result.persons.filter(p=>session.presentStudents.some(s=>s.id===p.id));
  if(!selected){if(animate)change(renderList);else renderList();}
 }catch(error){body.replaceChildren();const p=document.createElement('p');p.textContent='Die Schüler konnten nicht geladen werden.';const retry=document.createElement('button');retry.className='btn-2';retry.textContent='Erneut versuchen';retry.onclick=load;body.append(p,retry);notify.error(error.message,{title:'Hausaufgaben'});}
}
function change(update){if(motion.beginPageChange(()=>change(update)))return;update();window.repetiererKristall?.refresh();}
function renderList() {
 body.replaceChildren();const hint=document.createElement('p');hint.className='participation-help';hint.textContent=`Wähle eine anwesende Person. Ein Extra-Joker nach jeweils ${settings.homeworkJokerThreshold} Vorführungen. Gewicht neuer Noten: ${settings.homeworkWeight}×.`;
 const list=document.createElement('div');list.id='homework-student-list';
 for(const person of people){const button=document.createElement('button');button.type='button';button.className='participation-student';button.dataset.studentId=person.id;
  const name=document.createElement('strong'),count=document.createElement('span');name.textContent=person.name;count.textContent=`${person.homework.length} ${person.homework.length===1?'Vorführung':'Vorführungen'}`;button.append(name,count);
  button.addEventListener('click',async()=>{if(busy)return;const ready=await ensureColumns();if(!ready)return;change(()=>renderGrade(person));});list.append(button);
 }
 if(!people.length){const empty=document.createElement('p');empty.textContent='Heute sind keine anwesenden Personen verfügbar.';list.append(empty);}
 body.append(hint,list);window.repetiererKristall?.refresh();
}
function updateDetail(person) {body.querySelector('#homework-grade-detail').textContent=`Gewicht ${settings.homeworkWeight}× · Noch ${settings.homeworkJokerThreshold-person.homework.length%settings.homeworkJokerThreshold} ${settings.homeworkJokerThreshold-person.homework.length%settings.homeworkJokerThreshold===1?'Vorführung':'Vorführungen'} bis zum nächsten Extra-Joker.`;}
function renderGrade(person) {
 selected=person;body.innerHTML=`<div class="homework-stage"><button type="button" class="btn-2" id="homework-back-btn">Andere Person wählen</button><h1 id="homework-student-name"></h1><p>Freiwillige Hausaufgaben vorführen</p><form id="homework-grade-form"><label for="homework-grade-input">Note</label><input id="homework-grade-input" inputmode="decimal" type="text" placeholder="z.B. 5.5" required autocomplete="off"><button class="btn-1" type="submit" id="homework-save-btn">Note speichern</button></form><p id="homework-grade-detail"></p></div>`;
 body.querySelector('h1').textContent=person.name;updateDetail(person);
 body.querySelector('#homework-back-btn').onclick=()=>{if(!busy)change(()=>{selected=null;renderList();});};
 body.querySelector('form').addEventListener('submit',async event=>{
  event.preventDefault();if(busy)return;
  const input=body.querySelector('input'),grade=Number(input.value.trim().replace(',','.'));
  if(!input.value.trim() || !Number.isFinite(grade)||grade<1||grade>6){notify.error('Bitte eine Note zwischen 1 und 6 eingeben.',{title:'Hausaufgaben'});input.focus();return;}
  busy=true;body.querySelectorAll('button,input').forEach(e=>e.disabled=true);
  try {
   const result=await ipcRenderer.invoke('homework:save',{sessionId:session.id,personId:person.id,grade});
   if(!result.success)throw new Error(result.error||'Die Note konnte nicht gespeichert werden.');
   if(result.entry.excelWriteSucceeded)notify.success(`Die Note wurde gespeichert.${result.entry.awardExtraJoker?' Ein Extra-Joker wurde gutgeschrieben.':''}`,{title:'Hausaufgaben'});
   selected=null;await load(true);
  }catch(error){notify.error(error.message,{title:'Hausaufgaben nicht gespeichert'});}
  finally{busy=false;if(selected)body.querySelectorAll('button,input').forEach(e=>e.disabled=false);}
 });
 motion.revealStudent(body.querySelector('h1'));body.querySelector('input').focus();
}
for(const [kind,channel] of [['undo','undo-status'],['redo','redo-status']]){
 const button=document.getElementById(`homework-${kind}-btn`);button.onclick=()=>{if(!busy)ipcRenderer.send(`${kind}-last-action`);};
 ipcRenderer.on(channel,(_e,status)=>{button.disabled=!status?.available;});
 ipcRenderer.on(`${kind}-completed`,()=>{if(!page.classList.contains('update-hidden')){selected=null;load(true);}});
}
