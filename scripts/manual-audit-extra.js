// Compiled into manual-audit.js alongside the smoke runner.
var auditEntries = [], auditCovered = new Set();
async function saveImage(name, image, metadata = {}) {
 const file = `${String(auditEntries.length + 1).padStart(3,'0')}__${name}.png`;
 assert.ok(!image.isEmpty(), file + ' is empty');
 fs.writeFileSync(path.join(output,file), image.toPNG());
 auditEntries.push({file,name,...image.getSize(),...metadata});
 console.log('CAPTURE ' + file);
}
async function shot(name) {
 await pause(300); win.webContents.invalidate(); await pause(60);
 const state = await evaluate(`(() => ({modals:[...document.querySelectorAll('.modal')].filter(e=>getComputedStyle(e).display!=='none' && +getComputedStyle(e).opacity>.5).map(e=>e.id),theme:document.documentElement.dataset.theme}))()`);
 state.modals.forEach(id=>auditCovered.add(id));
 await saveImage(name, await win.webContents.capturePage(), state);
 // Capture all vertical content at its actual viewport size, including nested panes.
 const panes = await evaluate(`(() => [document.scrollingElement,...document.querySelectorAll('body *')].filter(e=>{const s=getComputedStyle(e),r=e.getBoundingClientRect();return (e===document.scrollingElement||['auto','scroll'].includes(s.overflowY))&&e.scrollHeight>e.clientHeight+8&&e.clientHeight>70&&r.width>50&&r.bottom>0&&r.top<innerHeight&&e.checkVisibility();}).map((e,i)=>{e.dataset.auditPane=String(i);return {key:i,top:e.scrollTop,max:e.scrollHeight-e.clientHeight,step:Math.max(80,Math.floor(e.clientHeight*.8))}}))()`);
 for (const pane of panes) {
  for(let y=0,page=1;y<pane.max+pane.step;y+=pane.step,page++) {
   const position=Math.min(y,pane.max); if(position===pane.top) continue;
   await evaluate(`document.querySelector('[data-audit-pane="${pane.key}"]').scrollTop=${position}`); await pause(90);
   await saveImage(`${name}__bereich-${pane.key+1}__seite-${page}`,await win.webContents.capturePage(),{...state,scrollTop:position});
   if(position===pane.max) break;
  }
  await evaluate(`document.querySelector('[data-audit-pane="${pane.key}"]').scrollTop=${pane.top}`);
 }
}
async function cleanAudit() {
 await evaluate(`(() => {document.querySelectorAll('.modal').forEach(e=>{e.style.display='none';e.classList.remove('modal-open');});document.getElementById('update-panel').classList.add('update-hidden');document.getElementById('drawer').classList.remove('drawer-open');document.getElementById('drawer-scrim').classList.remove('scrim-visible');})()`);
 await pause(100);
}
async function eventShot(name,channel,...args) {await cleanAudit();win.webContents.send(channel,...args);await shot(name);}
async function extraAudit() {
 win.setSize(1100,800); await cleanAudit();
 await evaluate(`document.documentElement.dataset.theme='midnight'`);
 await click('repetition-nav-btn'); await pause(300);
 for(const [name,button] of [['navigation','drawer-toggle'],['datenverwaltung','drawer-data-view-btn'],['protokoll','drawer-protocol-view-btn']]) {
  if(name==='protokoll') await click('drawer-back-btn');
  await click(button);await shot(name);
  await click('drawer-help-btn');await shot(name+'-hilfe');await click('close-teacher-help-btn');await pause(300);
 }
 await cleanAudit();
 for(const button of ['session-start-help-btn','repetition-help-btn']) {await click(button);await shot(button);await cleanAudit();}
 await eventShot('freiwillig-auswahl','persons-list',students);
 await evaluate(`document.querySelector('#person-list input').click()`);await shot('freiwillig-markiert');
 await eventShot('abwesenheiten-leer','absence-persons-data',[]);
 await eventShot('abwesenheiten-auswahl','absence-persons-data',students.map((s,i)=>({...s,absent:i===1})));
 await eventShot('chancen-leer','probability-data',[]);
 await eventShot('chancen-mit-boost-und-sperrzeit','probability-data',students.map((s,i)=>({...s,probability:i?.35:.65,weight:i?1:1.8,joker:i,grades:i?3:0,boostedNeverSelected:!i,wiggersRuleActive:!!i,wiggersRuleRemainingMinutes:12})));
 await evaluate(`document.getElementById('class-name').innerText='L23a'`);
 await eventShot('excel-editor-leer','editor-persons-data',[]);
 await eventShot('excel-editor-gefuellt','editor-persons-data',students.map(s=>({...s,grades:2,joker:1})));
 await evaluate(`document.getElementById('excel-editor-name').value='Nora Test'`);await click('excel-editor-add-btn');await shot('excel-editor-person-hinzugefuegt');
 await evaluate(`document.querySelector('#excel-editor-list button').click()`);await shot('excel-editor-loeschung-vorgemerkt');
 await eventShot('backup-leer','backup-data',{entries:[]},{backupPath:'Beispieldaten/backup.json'});
 await eventShot('backup-gefuellt','backup-data',{entries:[{type:'grade',grade:5.5,className:'L23a',personName:'Anna Müller',createdAt:'2026-10-20T08:30:00Z',excelWriteSucceeded:true,filePath:'Unterricht.xlsx'},{type:'joker',className:'L23a',personName:'Max Beispiel',createdAt:'2026-10-20T08:40:00Z',excelWriteSucceeded:false,filePath:'Unterricht.xlsx'}]},{backupPath:'Beispieldaten/backup.json'});
 await eventShot('logs-leer','log-data','',{logPath:'Beispieldaten/app.log'});
 await eventShot('logs-gefuellt','log-data','[INFO] Session gestartet: L23a\n[INFO] Note lokal gespeichert\n[WARN] Excel-Datei geöffnet; Änderung wartet.',{logPath:'Beispieldaten/app.log'});
 await eventShot('klassenimport-vorschau','class-list-import-preview',{detectedClassName:'L23a',count:2,names:students.map(s=>s.name)});
 await eventShot('klassenimport-lange-liste','class-list-import-preview',{detectedClassName:'M24b',count:32,truncated:true,names:Array.from({length:25},(_,i)=>`Beispielperson ${i+1}`)});
 await eventShot('excel-neuladen-konflikt','reload-excel-conflict',3);
 await eventShot('joker-migration','joker-migration-required',{emptyCells:18,usedJokerCells:4},'Unterricht.xlsx');
 await click('joker-migration-help-btn');await shot('joker-migration-hilfe');
 await cleanAudit();await click('settings-btn');await shot('einstellungen-alle-abschnitte');
 await click('settings-help-btn');await shot('einstellungen-hilfe');
 await cleanAudit();await click('settings-btn');win.webContents.send('ai-setup-progress',{state:'downloading',percent:42,message:'Sprachmodell wird heruntergeladen … 42 %'});await shot('einstellungen-ki-download');
 win.webContents.send('ai-setup-progress',{state:'error',message:'Modell konnte nicht geladen werden. Bitte erneut versuchen.'});await shot('einstellungen-ki-fehler');
 await eventShot('update-neuigkeiten','update-news-data',{version:'7.0.0',notes:fs.readFileSync(path.join(__dirname,'..','release-notes.txt'),'utf8')});
 await cleanAudit();await click('problem-report-btn');await shot('problembericht-laedt');
 win.webContents.send('report-preview-data',{appId:'Repetierer',type:'bug',version:'7.0.0',environment:'audit',consoleOutput:'[INFO] Beispiel-Session gestartet',metadata:{clientTimestamp:'2026-10-20T10:30:00+02:00',os:'Windows',platform:'win32'}});await shot('problembericht-vorschau');
 await click('toggle-problem-report-details-btn');await shot('problembericht-details');
 await click('send-problem-report-btn');await shot('problembericht-sendet');
 win.webContents.send('problem-report-result',{ok:false});await shot('problembericht-fehler');
 for(const status of ['available','downloading','progress','downloaded','installing','installed','error']) await eventShot('update-'+status,'update-status',{status,version:'7.0.1',percent:42,message:'Verbindung fehlgeschlagen.',error:'Beispiel: Server nicht erreichbar.'});
 for(const [channel,args] of [['saved-file-missing',['Unterricht.xlsx']],['excel-write-pending',[{personName:'Anna Müller'}]],['pending-excel-flushed',[{written:2,remaining:0}]],['excel-reload-failed',[{reason:'excel-locked'}]],['undo-failed',[{}]],['redo-failed',[{}]],['session-protocol-empty',[]],['session-protocol-exported',[{pdfPath:'Protokoll.pdf',csvPath:'Protokoll.csv'}]],['session-protocol-export-failed',[{}]],['class-list-import-failed',[{reason:'invalid-file'}]],['joker-migration-failed',[{reason:'excel-locked'}]],['teaching-session-start-failed',[{}]],['teaching-session-end-failed',[{}]],['editor-persons-save-failed',[{reason:'excel-locked'}]]]) await eventShot('meldung-'+channel,channel,...args);
 await assessmentAudit();
 const all = await evaluate(`[...document.querySelectorAll('.modal')].map(e=>e.id)`);
 const missing = all.filter(id=>!auditCovered.has(id));
 assert.deepEqual(missing,[], 'Every modal has a screenshot');
}
async function assessmentAudit() {
 await cleanAudit();
 // Same renderer and store; control backend responses only for rare round states.
 let mode='open';
 ipcMain.removeHandler('assessment:get-overview');ipcMain.handle('assessment:get-overview',()=>ok(presentClass(assessmentStore.getClassContext(workbook,'L23a',students))));
 ipcMain.removeHandler('assessment:brevo-status');ipcMain.handle('assessment:brevo-status',()=>ok({configured:true}));
 ipcMain.removeHandler('assessment:get-round');ipcMain.handle('assessment:get-round',()=>{
  const v=commentView();v.assessmentConfigured=true;v.round.cloudStatus=mode==='closed'?'closed':'open';
  v.deliveryValidation={valid:true,incomplete:[]};v.invitation={subject:'Selbstbeurteilung HS 2026',text:'Bitte fülle die Selbstbeurteilung bis zum 30. Oktober aus.\nhttps://beispiel.invalid/?round=audit',complete:'Selbstbeurteilung HS 2026\nBitte fülle die Selbstbeurteilung aus.'};
  v.summary.failedCount=1;v.summary.sentCount=mode==='sent'?1:0;
  if(mode==='conflicts') {v.round.unmatchedResponses=[{name:'Anna M.',email:'anna@example.invalid',externalResponseId:'unmatched'}];v.round.duplicates=[{email:'max@example.invalid',count:2,selectedResponseId:'latest',responses:[{externalResponseId:'older',timestamp:'2026-10-19T10:00:00Z'},{externalResponseId:'latest',timestamp:'2026-10-20T10:00:00Z'}]}];v.summary.duplicateCount=1;v.summary.lastSyncStatus='failed';v.summary.lastSyncResult={invalid:1,ignoredRound:1};}
  return ok(v);
 });
 await click('assessment-nav-btn');await pause(300);await shot('beurteilungen-klassenwahl');
 await evaluate(`(()=>{const s=document.getElementById('assessment-class-select');s.value='L23a';s.dispatchEvent(new Event('change'));})()`);await waitFor(`!!document.querySelector('[data-round-id]')`);await shot('beurteilungen-rundenliste');
 const open=async()=>{await evaluate(`document.querySelector('[data-round-id]').click()`);await waitFor(`!!document.getElementById('assessment-teacher-btn')`);};
 await open();await shot('beurteilungen-runde-offen');
 await evaluate(`document.querySelector('.assessment-danger-zone').open=true`);await shot('beurteilungen-weitere-aktionen');
 await click('assessment-delete-round-btn');await shot('beurteilungen-loeschen-bestaetigung');await click('assessment-delete-cancel-btn');await pause(200);
 await click('assessment-send-btn');await shot('beurteilungen-versand-bestaetigung');await click('assessment-confirm-cancel-btn');await pause(200);
 await click('assessment-retry-btn');await shot('beurteilungen-versand-wiederholen');await click('assessment-confirm-cancel-btn');await pause(200);
 await evaluate(`document.getElementById('assessment-invitation-area').classList.add('invitation-expanded')`);await shot('beurteilungen-einladung-ausgeklappt');
 for(const m of ['closed','conflicts','sent']) {mode=m;await click('assessment-back-btn');await waitFor(`!!document.querySelector('[data-round-id]')`);await open();await shot('beurteilungen-runde-'+m);if(m==='sent'){await click('assessment-delete-round-btn');await shot('beurteilungen-loeschen-versandwarnung');await click('assessment-delete-cancel-btn');await pause(200);}}
 await click('assessment-teacher-btn');await shot('beurteilungen-lehrer-kommentare');
 await click('teacher-next-btn');await pause(200);await shot('beurteilungen-lehrer-naechster');
}
async function surveyAudit() {
 const http=require('node:http');
 const {DEFAULT_CRITERIA}=require('../src/assessment/assessmentStore');
 let surveyMode='open',postMode='success',pending;
 const roundId='00000000-0000-4000-8000-000000000001';
 const server=http.createServer((req,res)=>{
  if(req.url.includes('/api/')) {
   res.setHeader('Content-Type','application/json');
   if(req.method==='POST') {if(postMode==='pending'){pending=res;return;}res.statusCode=postMode==='success'?200:409;res.end(JSON.stringify(postMode==='success'?{ok:true}:{error:postMode}));return;}
   if(surveyMode==='loading'){pending=res;return;}
   if(surveyMode==='error'){res.statusCode=503;res.end('{}');return;}
   res.end(JSON.stringify({round:{title:'HS 2026 · L23a',closed:surveyMode==='closed',deadline:'2026-10-30T17:00:00Z',expiresAt:'2027-01-01T00:00:00Z',criteria:DEFAULT_CRITERIA}}));return;
  }
  const file=req.url.split('?')[0];
  const base=path.join(__dirname,'..','cloudflare');
  const target=file.includes('Geist-Variable')?path.join(base,'assets','fonts','Geist-Variable.woff2'):path.join(base,'web',file==='/'?'index.html':path.basename(file));
  if(!fs.existsSync(target)){res.statusCode=404;res.end();return;}
  res.setHeader('Content-Type',target.endsWith('.css')?'text/css':target.endsWith('.js')?'text/javascript':target.endsWith('.woff2')?'font/woff2':'text/html');res.end(fs.readFileSync(target));
 });
 await new Promise(r=>server.listen(0,'127.0.0.1',r));
 const mainWin=win;
 win=new BrowserWindow({width:1100,height:800,show:false,webPreferences:{offscreen:true,backgroundThrottling:false}});
 const url=`http://127.0.0.1:${server.address().port}/?round=${roundId}`;
 const load=async mode=>{surveyMode=mode;await win.loadURL(url);await evaluate(`localStorage.clear()`);await win.reload();await pause(350);};
 await win.loadURL(url.split('?')[0]);await shot('umfrage-link-unvollstaendig');
 await load('loading');await shot('umfrage-laedt');pending.end(JSON.stringify({round:{title:'HS 2026',closed:true}}));
 await load('error');await shot('umfrage-ladefehler');
 await load('closed');await shot('umfrage-geschlossen');
 await load('open');await evaluate('document.fonts.ready');await shot('umfrage-formular-leer');
 await evaluate(`(()=>{document.querySelector('[name=name]').value='Anna Müller';document.querySelector('[name=email]').value='anna@example.invalid';document.querySelectorAll('fieldset').forEach(e=>e.querySelector('input[value="4"]').click());document.querySelector('[name=additionalQuality]').value='Ich bringe eigene Beispiele ein.';document.querySelector('[name=comment]').value='Ich möchte mich häufiger melden.';})()`);await shot('umfrage-formular-ausgefuellt');
 const submit=()=>evaluate(`document.getElementById('survey').requestSubmit()`);
 for(const code of ['duplicate_response','round_closed','round_closed_or_full','round_not_found','invalid_email','invalid_answers']) {postMode=code;await submit();await pause(200);await evaluate(`document.getElementById('submit-status').scrollIntoView({block:'end'})`);await shot('umfrage-abgabefehler-'+code);}
 postMode='pending';await submit();await pause(100);await shot('umfrage-wird-abgegeben');pending.end('{}');await pause(200);await shot('umfrage-erfolgreich-abgegeben');
 await win.reload();await pause(200);await shot('umfrage-bereits-abgegeben');
 win.setSize(390,844);await load('open');await shot('umfrage-mobil-formular');
 win.destroy();win=mainWin;await new Promise(r=>server.close(r));
}
async function finishAudit() {
 const escape=s=>String(s).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('"','&quot;');
 const expected=await evaluate(`[...document.querySelectorAll('.modal')].map(e=>e.id)`);
 const manifest={createdAt:new Date().toISOString(),description:'Echter Electron-Renderer; isolierte synthetische IPC-Daten. Umfrage über lokalen HTTP-Testserver. Keine Live-Mails, Updates oder Berichte.',knownIssues:fs.existsSync(path.join(output,'renderer-issues.json'))?JSON.parse(fs.readFileSync(path.join(output,'renderer-issues.json'))):[],modalCoverage:expected,entries:auditEntries};
 fs.writeFileSync(path.join(output,'manifest.json'),JSON.stringify(manifest,null,2));
 const template=fs.readFileSync(path.join(__dirname,'manual-audit-template.html'),'utf8');
 const cards=auditEntries.map((e,i)=>`<article id="bild-${i+1}" data-file="${escape(e.file)}"><h2>${i+1}. ${escape(e.name)}</h2><p><a href="${escape(e.file)}" target="_blank">${escape(e.file)}</a> · ${e.width} × ${e.height}</p><img src="${escape(e.file)}" loading="lazy" alt="${escape(e.name)}"><label for="comment-${i+1}">Dein Kommentar</label><textarea id="comment-${i+1}" rows="4" placeholder="Was soll sich auf diesem Screen ändern?"></textarea><small class="save-state"></small></article>`).join('\n');
 fs.writeFileSync(path.join(output,'index.html'),template.replace('<!-- AUDIT_CARDS -->',cards).replaceAll('AUDIT_COUNT',String(auditEntries.length)));
 fs.writeFileSync(path.join(output,'README.md'),`# Manueller UI-Audit\n\nindex.html im Browser öffnen. ${auditEntries.length} eindeutig benannte Screenshots. Unter jedem Bild kommentieren; oben als Markdown herunterladen. Kommentare werden automatisch im aktuellen Browser gespeichert. Markdown ist die portable Sicherung; für Browserwechsel kann sie wieder importiert werden.\n\nAlle ${expected.length} HTML-Modals erfasst. Lange Ansichten sind zusätzlich in überlappenden Scrollabschnitten aufgenommen. Hauptansicht 1100 × 800, kompakte Ansichten und mobile Umfrage zusätzlich.\n\nEchte aktuelle Rendererdateien mit synthetischen Beispieldaten und kontrollierten IPC-Antworten. Seltene Fehler-, Update- und Versandzustände werden gezielt ausgelöst. Das ist eine visuelle Bestandsaufnahme, kein Nachweis echter Server- oder E-Mail-Funktion. Betriebssystem-Dateiauswahl und Installer sind keine eigenen App-Screens und nicht enthalten.\n\nReproduzieren: im Projekt Repetierer: node scripts/manual-audit.js. Alte Builds werden nicht verändert.\n`);
 console.log(`AUDIT_PASS: ${auditEntries.length} screenshots; ${expected.length} modals; ${output}`);
}
