const {normalizeEmail,isValidEmail}=require('./matching');
const timestamp=value=>Number.isFinite(Date.parse(value))?Date.parse(value):0;
const responseKey=(roundId,response)=>`${roundId}:${response.externalResponseId||response.responseId||''}:${response.timestamp||response.submittedAt||''}:${normalizeEmail(response.email||response.studentEmail)}`;

// Private, append-only history. Only student.email crosses the renderer boundary.
function initializeContact(student,classRecord){
  if(Array.isArray(student.emailHistory))return;
  student.emailHistory=[];
  const legacy=normalizeEmail(student.email);
  if(legacy)student.emailHistory.push({email:legacy,source:'legacy'});
  const submissions=classRecord.rounds.flatMap(round=>{
    const a=round.assessments?.[student.id];
    return a?.studentEmail?[{roundId:round.id,email:a.studentEmail,submittedAt:a.submittedAt,externalResponseId:a.externalResponseId||a.responseId}]:[];
  }).sort((a,b)=>timestamp(a.submittedAt)-timestamp(b.submittedAt));
  for(const submission of submissions)recordSubmission(student,submission.roundId,submission);
}
function recordSubmission(student,roundId,response){
  const email=normalizeEmail(response.email||response.studentEmail);
  if(!isValidEmail(email))return;
  const key=responseKey(roundId,response);
  if(student.emailHistory.some(entry=>entry.key===key))return;
  const submittedAt=response.timestamp||response.submittedAt||null;
  const latest=student.emailHistory.filter(entry=>entry.source==='student').reduce((max,entry)=>Math.max(max,timestamp(entry.submittedAt)),0);
  student.emailHistory.push({email,source:'student',key,roundId,submittedAt,changedAt:new Date().toISOString()});
  if(timestamp(submittedAt)>=latest)student.email=email;
}
function setManualEmail(student,email){
  if(typeof email!=='string'||email.length>254||!isValidEmail(email)||/[<>,;:"\\\r\n]/.test(email))throw new Error('Bitte eine gültige einzelne E-Mail-Adresse eingeben.');
  const normalized=normalizeEmail(email);
  if(normalized!==student.email){student.emailHistory.push({email:normalized,source:'manual',changedAt:new Date().toISOString()});student.email=normalized;}
}
function publicStudent(student){const {emailHistory,...visible}=student;return visible;}
function publicRound(round,students){
  const visible=JSON.parse(JSON.stringify(round));
  delete visible.externalResponses;
  delete visible.preferredResponseIds;
  const replacements=new Map();
  for(const student of students){
    if(visible.assessments?.[student.id])visible.assessments[student.id].studentEmail=student.email||'';
    for(const entry of student.emailHistory||[])if(entry.email!==student.email)replacements.set(entry.email,student.email||'');
  }
  for(const response of visible.unmatchedResponses||[])if(replacements.has(response.email))response.email=replacements.get(response.email);
  for(const group of visible.duplicates||[]){
    if(replacements.has(group.email))group.email=replacements.get(group.email);
    for(const response of group.responses||[])if(replacements.has(response.email))response.email=replacements.get(response.email);
  }
  return visible;
}
module.exports={initializeContact,recordSubmission,setManualEmail,publicStudent,publicRound};
