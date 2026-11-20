const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),os=require('node:os'),path=require('node:path');
const Excel=require('exceljs'),{AssessmentStore,DEFAULT_CRITERIA}=require('../src/assessment/assessmentStore');
const {loadDossiers,participationForClass}=require('../src/studentDossiers'),{dossierHtml,pdfHtml}=require('../src/studentDossierView');
const {suggestedGrade}=require('../src/assessment/grade');
async function fixture(){
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-dossier-test-')),filePath=path.join(dir,'Klassen.xlsx'),paths={assessmentPath:path.join(dir,'assessments.json')},book=new Excel.Workbook();
  for(const name of ['Klasse A','Klasse B']){const s=book.addWorksheet(name);s.getCell('A1').value='repetierer';s.getCell('A6').value='Anna <Beispiel>';s.getCell('A7').value='Max';s.getCell('B6').value=4;s.getCell('K6').value='10.11.2026';require('../src/jokerDates').migrateWorksheet(s);require('../src/homework').migrateWorksheet(s);}
  const sheet=book.getWorksheet('Klasse A');require('../src/jokerDates').append(sheet,6,'09.11.2026');
  require('../src/homework').append(sheet,6,{id:'h1',grade:6,weight:1.2,date:'11.11.2026',voluntary:true});
  await book.xlsx.writeFile(filePath);
  const store=new AssessmentStore(paths.assessmentPath),persons=[{id:0,name:'Anna <Beispiel>'},{id:1,name:'Max'}];
  const round=store.createRound(filePath,'Klasse A',persons,{title:'Mündliche Runde',deadline:'2026-11-20'}),student=store.getClassContext(filePath,'Klasse A',persons).students[0];
  const answers=Object.fromEntries(round.criteria.map(c=>[c.id,4]));store.saveTeacherAssessment(filePath,'Klasse A',persons,round.id,student.id,answers,'Lehrerkommentar',{'participation_frequency':'Detailkommentar'},5.5);
  const target=store.data.workbooks[require('../src/assessment/assessmentStore').hashPath(filePath)];const a=Object.values(target.classes)[0].rounds[0].assessments[student.id];a.studentAnswers=answers;a.studentComment='Schülerkommentar';a.studentAdditionalQuality='Zusatz';store.save();
  const sessions=[{id:'s1',filePath,className:'Klasse A',startedAt:'2026-11-11',presentStudents:persons,absentStudents:[],participationEntries:[{studentId:0,studentName:persons[0].name,points:5,timestamp:'2026-11-11'}],homeworkEntries:[]},{id:'s2',filePath,className:'Klasse A',presentStudents:[persons[1]],absentStudents:[persons[0]],participationEntries:[]},{id:'other',filePath:filePath+'.other',className:'Klasse A',presentStudents:persons,participationEntries:[{studentName:persons[0].name,points:1}]},{id:'legacy',className:'Klasse A',presentStudents:persons,participationEntries:[]}];
  return {dir,filePath,paths,sessions,store,round,student,persons};
}
test('dossiers join Excel, oral assessments and scoped sessions without writing or double counting',async()=>{
  const f=await fixture(),before=fs.readFileSync(f.filePath),oralBefore=fs.readFileSync(f.paths.assessmentPath);
  const data=await loadDossiers({...f,pending:[{type:'homework',personName:f.persons[0].name,filePath:f.filePath,className:'Klasse A',homeworkId:'h1',grade:6,weight:1.2},{type:'grade',personName:f.persons[0].name,filePath:f.filePath,className:'Klasse A',grade:5,date:'12.11.2026'}]});
  const d=data.classes[0].students[0];assert.equal(data.classes.length,2);assert.equal(d.repetition.length,2);assert.equal(d.homework.length,1);assert.equal(d.jokers.length,1);assert.equal(d.summary.voluntaryCount,1);assert.equal(d.summary.combinedAverage,5);assert.equal(d.participation.visitedSessions,1);assert.equal(d.participation.absentSessions,1);assert.equal(d.participation.ratingCount,1);assert.equal(d.participation.averagePoints,5);assert.equal(d.summary.oralAverage,5.5);assert.equal(data.legacySessionCount,1);
  assert.equal(data.classes[1].students[0].assessments.length,0);assert.equal(data.classes[1].students[0].homework.length,0);
  assert.deepEqual(fs.readFileSync(f.filePath),before);assert.deepEqual(fs.readFileSync(f.paths.assessmentPath),oralBefore);
  const html=dossierHtml(d);assert.match(html,/Schülerkommentar/);assert.match(html,/Detailkommentar/);assert.match(html,/<details/);assert.doesNotMatch(html,/<input|<textarea|contenteditable/);assert.match(pdfHtml(d),/Anna &lt;Beispiel&gt;/);assert.doesNotMatch(pdfHtml(d),/<details/);
});
test('six completed repetition slots remain visible; missing history stays unknown',async()=>{
  const f=await fixture(),book=new Excel.Workbook();await book.xlsx.readFile(f.filePath);const s=book.getWorksheet('Klasse A');for(let c=2;c<=7;c++)s.getCell(7,c).value=4;await book.xlsx.writeFile(f.filePath);
  const d=(await loadDossiers(f)).classes[0].students[1];assert.equal(d.repetition.length,6);assert.equal(d.repetition[0].date,null);assert.equal(d.summary.oralAverage,null);assert.match(dossierHtml(d),/Datum nicht erfasst/);
  const result=participationForClass(f.filePath,'Klasse A',[{id:4,name:f.persons[0].name}],f.sessions);assert.equal(result[0].ratingCount,1,'Roster reorder follows name, not reused row ID');
});
test('grade proposal reverses negative criteria, requires all ratings, stores manual grades and rejects invalid grades',async()=>{
  const perfect=Object.fromEntries(DEFAULT_CRITERIA.map(c=>[c.id,['insufficient_contributions','original_unsuitable_contributions'].includes(c.id)?1:5]));assert.equal(suggestedGrade(DEFAULT_CRITERIA,perfect),6);assert.equal(suggestedGrade(DEFAULT_CRITERIA,{}),null);assert.equal(suggestedGrade(DEFAULT_CRITERIA,Object.fromEntries(DEFAULT_CRITERIA.map(c=>[c.id,3]))),3.5);
  const f=await fixture();assert.throws(()=>f.store.saveTeacherAssessment(f.filePath,'Klasse A',f.persons,f.round.id,f.student.id,perfect,'',{},7),/zwischen 1 und 6/);
  f.store.saveTeacherAssessment(f.filePath,'Klasse A',f.persons,f.round.id,f.student.id,perfect,'',{},4.8);
  const reloaded=new AssessmentStore(f.paths.assessmentPath).getRound(f.filePath,'Klasse A',f.persons,f.round.id).round.assessments[f.student.id];assert.equal(reloaded.teacherGrade,4.8);assert.equal(reloaded.suggestedGrade,6);assert.equal(reloaded.teacherGradeManual,true);
});
test('corrupt oral history fails visibly instead of silently dropping student information',async()=>{const f=await fixture();fs.writeFileSync(f.paths.assessmentPath,'broken');await assert.rejects(loadDossiers(f),/nicht gelesen/);});
test('equal names use matching source IDs and never copy one pupil contribution to both',async()=>{
  const people=[{id:0,name:'Max Muster'},{id:1,name:'Max Muster'}],filePath=path.join(os.tmpdir(),'same-names.xlsx');
  const sessions=[{filePath,className:'A',presentStudents:people,absentStudents:[],participationEntries:[{studentId:0,studentName:'Max Muster',points:5}]}];
  const result=participationForClass(filePath,'A',people,sessions);assert.equal(result[0].ratingCount,1);assert.equal(result[1].ratingCount,0);
  assert.equal(participationForClass(filePath,'A',[people[0]],sessions)[0].ratingCount,0,'A removed homonym makes the historical identity ambiguous');
});
test('PDF requests reject another workbook or a changed fingerprint before opening a dialog',async()=>{
  const f=await fixture(),handlers=new Map(),data=await loadDossiers(f);let dialogs=0,currentFile=f.filePath;
  require('../src/studentDossierIpc').registerStudentDossiers({ipcMain:{handle:(name,fn)=>handlers.set(name,fn)},BrowserWindow:{fromWebContents:()=>null},dialog:{showSaveDialog:async()=>{dialogs++;return {canceled:true};}},getProgram:()=>({getCurrentFilePath:()=>currentFile}),getPaths:()=>f.paths,getSessions:()=>({readData:()=>({sessions:f.sessions})}),getPending:()=>[]});
  const invoke=args=>handlers.get('dossiers:export')({sender:{}},{scope:'student',className:'Klasse A',studentId:0,studentName:f.persons[0].name,filePath:f.filePath,fingerprint:data.fingerprint,...args});
  assert.equal((await invoke({filePath:f.filePath+'.other'})).ok,false);assert.equal((await invoke({fingerprint:'changed'})).ok,false);assert.equal((await invoke({studentName:'Another student'})).ok,false);assert.equal(dialogs,0);
  assert.equal((await invoke({})).cancelled,true);assert.equal(dialogs,1);
});

test('current contact persists manual changes, new submissions replace them and old round sync cannot regress it',async()=>{
 const f=await fixture(),s=f.store,submit=(round,email,id,timestamp)=>s.applyResponses(f.filePath,'Klasse A',f.persons,round.id,[{externalResponseId:id,email,name:f.student.name,timestamp,answers:{participation_frequency:4}}]);
 submit(f.round,'old@example.ch','old','2026-11-01T10:00:00Z');
 s.setStudentEmail(f.filePath,'Klasse A',f.persons,0,f.student.name,'manual@example.ch','old@example.ch');
 submit(f.round,'old@example.ch','old','2026-11-01T10:00:00Z');
 assert.equal(s.getClassContext(f.filePath,'Klasse A',f.persons).students[0].email,'manual@example.ch','Replayed known response preserves manual address');
 const newer=s.createRound(f.filePath,'Klasse A',f.persons,{title:'Neu',deadline:'2026-11-20'});submit(newer,'new@example.ch','new','2026-11-17T10:00:00Z');
 submit(f.round,'older@example.ch','late-old','2026-11-02T10:00:00Z');
 assert.equal(s.getClassContext(f.filePath,'Klasse A',f.persons).students[0].email,'new@example.ch');
 const raw=new AssessmentStore(f.paths.assessmentPath).getClassContext(f.filePath,'Klasse A',f.persons).students[0];assert.equal(raw.email,'new@example.ch');for(const email of ['old@example.ch','manual@example.ch','older@example.ch'])assert.ok(raw.emailHistory.some(h=>h.email===email));
 const data=await loadDossiers(f),serialized=JSON.stringify(data);assert.equal(data.classes[0].students[0].email,'new@example.ch');assert.doesNotMatch(serialized,/old@example|manual@example|older@example|emailHistory|studentEmail/);
 const {publicRound,publicStudent}=require('../src/assessment/studentContact');assert.doesNotMatch(JSON.stringify(publicRound(s.getRound(f.filePath,'Klasse A',f.persons,f.round.id).round,[raw])),/old@example|older@example/);assert.equal(publicStudent(raw).emailHistory,undefined);
 assert.throws(()=>s.setStudentEmail(f.filePath,'Klasse A',f.persons,0,f.student.name,'wrong','new@example.ch'),/gültige/);assert.equal(s.getClassContext(f.filePath,'Klasse A',f.persons).students[0].email,'new@example.ch');
 assert.throws(()=>s.setStudentEmail(f.filePath,'Klasse A',f.persons,0,f.student.name,'ok@example.ch','old@example.ch'),/inzwischen/);
});

test('legacy contact history uses submission dates across rounds without modifying stored source on read',async()=>{
 const f=await fixture(),cls=Object.values(Object.values(f.store.data.workbooks)[0].classes)[0],student=cls.students[0];delete student.emailHistory;student.email='old@example.ch';
 const a=cls.rounds[0].assessments[student.id];a.studentEmail='new@example.ch';a.submittedAt='2026-11-17T10:00:00Z';a.externalResponseId='new';
 const old=structuredClone(cls.rounds[0]);old.id='old';old.assessments[student.id].studentEmail='old@example.ch';old.assessments[student.id].submittedAt='2026-11-01T10:00:00Z';old.assessments[student.id].externalResponseId='old';cls.rounds.unshift(old);f.store.save();
 const bytes=fs.readFileSync(f.paths.assessmentPath),d=(await loadDossiers(f)).classes[0].students[0];assert.equal(d.email,'new@example.ch');assert.deepEqual(fs.readFileSync(f.paths.assessmentPath),bytes);
});

test('homonymous pupil contacts remain distinct and failed saves restore memory',async()=>{
 const f=await fixture(),persons=[{id:0,name:'Max'},{id:1,name:'Max'}],s=f.store;
 s.setStudentEmail(f.filePath,'Klasse B',persons,0,'Max','one@example.ch','');s.setStudentEmail(f.filePath,'Klasse B',persons,1,'Max','two@example.ch','');
 const cls=s.getClassContext(f.filePath,'Klasse B',persons);assert.equal(cls.students.find(p=>p.sourceId===0).email,'one@example.ch');assert.equal(cls.students.find(p=>p.sourceId===1).email,'two@example.ch');
 s.save=()=>{throw new Error('fixture disk failure');};assert.throws(()=>s.setStudentEmail(f.filePath,'Klasse B',persons,0,'Max','changed@example.ch','one@example.ch'),/disk failure/);assert.equal(s.getClassContext(f.filePath,'Klasse B',persons).students.find(p=>p.sourceId===0).email,'one@example.ch');
});

test('an inserted homonym never takes the contact of an existing later source row',async()=>{
 const f=await fixture(),s=f.store,existing=[{id:1,name:'Max'}];
 s.setStudentEmail(f.filePath,'Klasse B',existing,1,'Max','existing@example.ch','');
 const students=s.getClassContext(f.filePath,'Klasse B',[{id:0,name:'Max'},...existing]).students.filter(p=>p.active);
 assert.equal(students.find(p=>p.sourceId===0).email,'');
 assert.equal(students.find(p=>p.sourceId===1).email,'existing@example.ch');
 const cls=s.getClassContext(f.filePath,'Klasse B',existing);
 assert.equal(cls.students.filter(p=>p.active).length,1);
 const round={assessments:{[students.find(p=>p.sourceId===1).id]:{studentEmail:'old@example.ch'}},externalResponses:[{email:'old@example.ch'}],preferredResponseIds:{'old@example.ch':'old'}};
 const visible=require('../src/assessment/studentContact').publicRound(round,cls.students);
 assert.doesNotMatch(JSON.stringify(visible),/old@example|preferredResponseIds|externalResponses|emailHistory/);
});
