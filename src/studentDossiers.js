const path = require('node:path');
const fs = require('node:fs'),crypto=require('node:crypto');
const Excel = require('exceljs');
const homework = require('./homework');
const jokerDates = require('./jokerDates');
const {AssessmentStore, hashPath} = require('./assessment/assessmentStore');
const normalize = value => String(value || '').normalize('NFKC').trim().replace(/\s+/g, ' ').toLocaleLowerCase('de');
const sameFile = (a,b) => !!a && !!b && path.resolve(a).toLowerCase() === path.resolve(b).toLowerCase();
const mean = values => values.length ? values.reduce((sum,v)=>sum+v,0)/values.length : null;
const entryName=entry=>normalize(entry.studentName ?? entry.personName ?? entry.name);
function identityMatcher(persons, sourceRoster=persons){
  const count=(roster,name)=>roster.filter(p=>entryName(p)===name).length;
  return (entry,person)=>{
    const name=normalize(person.name);if(entryName(entry)!==name)return false;
    const currentCount=count(persons,name),sourceCount=count(sourceRoster,name);
    if(currentCount>1 || sourceCount>1){
      if(currentCount!==sourceCount)return false;
      const id=entry.studentId??entry.personId??entry.id;
      return id!==undefined && String(id)===String(person.id);
    }
    return true;
  };
}

function participationForClass(filePath, className, persons, sessions) {
  const relevant=sessions.filter(s=>sameFile(s.filePath,filePath) && s.className===className);
  const matches=(session,entry,person)=>identityMatcher(persons,[...(session.presentStudents||[]),...(session.absentStudents||[])])(entry,person);
  return persons.map(person=>{
    const present=relevant.filter(s=>(s.presentStudents||[]).some(p=>matches(s,p,person)));
    const absent=relevant.filter(s=>(s.absentStudents||[]).some(p=>matches(s,p,person)));
    const entries=present.flatMap(s=>(s.participationEntries||[]).filter(e=>matches(s,e,person)).map(e=>({...e,sessionDate:s.startedAt})));
    const participating=present.filter(s=>(s.participationEntries||[]).some(e=>matches(s,e,person))).length;
    return {studentId:person.id,studentName:person.name,visitedSessions:present.length,absentSessions:absent.length,sessionsWithParticipation:participating,ratingCount:entries.length,averagePoints:mean(entries.map(e=>e.points)),participationRate:present.length?participating/present.length:null,entries};
  });
}

// Read a separate workbook: opening an archive never changes the active session/class.
async function loadDossiers({filePath, paths, sessions, pending=[]}) {
  if(!filePath)throw new Error('Bitte zuerst eine Repetierer-Datei laden.');
  const bytes=fs.readFileSync(filePath),fingerprint=crypto.createHash('sha256').update(bytes).digest('hex');
  const book=new Excel.Workbook();await book.xlsx.load(bytes);
  const assessments=new AssessmentStore(paths.assessmentPath).data.workbooks[hashPath(filePath)];
  const classes=[];
  for(const sheet of book.worksheets.filter(s=>s.getCell('A1').value==='repetierer')) {
    const persons=[];
    for(let row=6;row<=30;row++){const name=sheet.getCell(row,1).value;if(!name)break;persons.push({id:row-6,name:String(name)});}
    const belongs=identityMatcher(persons);
    const participation=participationForClass(filePath,sheet.name,persons,sessions);
    const assessmentClass=Object.values(assessments?.classes||{}).find(c=>c.name===sheet.name);
    const records=persons.map(person=>{
      const row=person.id+6;
      const repetition=[];
      for(let col=2;col<=7;col++){
        const raw=sheet.getCell(row,col).value;
        if(raw===null || raw===undefined || raw==='')continue;
        const grade=typeof raw==='string'?Number(raw.replace(',','.')):raw;
        repetition.push({grade:Number.isFinite(grade)?grade:null,date:sheet.getCell(row,col+9).value||null,position:col-1});
      }
      const home=homework.history(sheet,row);
      const current=jokerDates.layout(sheet);
      const jokers=(current.columns.length?current.columns.flatMap(c=>jokerDates.dates(sheet.getCell(row,c.column).value)):jokerDates.dates(sheet.getCell(row,17).value)).map(date=>({date}));
      const waiting=pending.filter(e=>sameFile(e.filePath,filePath)&&e.className===sheet.name&&belongs(e,person));
      for(const e of waiting){
        if(e.type==='grade')repetition.push({grade:e.grade,date:e.date,pending:true});
        if(e.type==='joker')jokers.push({date:e.date,pending:true});
        if(e.type==='homework'&&!home.some(h=>h.id===e.homeworkId))home.push({...e,id:e.homeworkId,pending:true});
      }
      const sessionHome=sessions.filter(s=>sameFile(s.filePath,filePath)&&s.className===sheet.name).flatMap(s=>s.homeworkEntries||[]).filter(e=>!e.undone&&belongs(e,person));
      const history=home.map(h=>{const source=sessionHome.find(e=>e.homeworkId===h.id);return {...h,voluntary:h.voluntary??source?.voluntary??null};});
      const candidates=(assessmentClass?.students||[]).filter(s=>normalize(s.name)===normalize(person.name));
      const exact=candidates.filter(s=>s.sourceId===person.id);
      const student=exact.length===1?exact[0]:candidates.length===1&&persons.filter(p=>normalize(p.name)===normalize(person.name)).length===1?candidates[0]:null;
      const rounds=(assessmentClass?.rounds||[]).flatMap(round=>{
        const a=round.assessments?.[student?.id];
        const {studentEmail,...visible}=a||{};
        return a && (a.studentAnswers||a.teacherAnswers)?[{id:round.id,title:round.title,deadline:round.deadline,createdAt:round.createdAt,criteria:round.criteria,...visible}]:[];
      });
      const entries=[...repetition.filter(e=>Number.isFinite(e.grade)).map(e=>({...e,weight:1})),...history];
      return {...person,email:student?.email||'',className:sheet.name,repetition,homework:history,jokers,assessments:rounds,participation:participation.find(p=>p.studentId===person.id),summary:{repetitionCount:repetition.length,homeworkCount:history.length,repetitionAverage:mean(repetition.filter(e=>Number.isFinite(e.grade)).map(e=>e.grade)),homeworkAverage:mean(history.map(e=>e.grade)),combinedAverage:mean(entries.map(e=>e.grade)),weightedAverage:entries.length?entries.reduce((s,e)=>s+e.grade*e.weight,0)/entries.reduce((s,e)=>s+e.weight,0):null,oralAverage:mean(rounds.filter(a=>Number.isFinite(a.teacherGrade)).map(a=>a.teacherGrade)),voluntaryCount:history.filter(h=>h.voluntary===true).length,voluntaryUnknown:history.filter(h=>h.voluntary===null).length}};
    });
    classes.push({name:sheet.name,students:records});
  }
  return {filePath,fingerprint,classes,legacySessionCount:sessions.filter(s=>!s.filePath).length};
}
module.exports={loadDossiers,participationForClass,sameFile};
