const {DEFAULT_CRITERIA} = require('./assessmentStore');
const {buildResultMessage} = require('./mailTemplates');
const {isValidEmail} = require('./matching');

// The production template with fictional data. No workbook or delivery status changes.
function buildDebugMessage(recipient) {
  recipient=String(recipient || '').trim();
  if(!isValidEmail(recipient))throw new Error('Bitte eine gültige Empfängeradresse eingeben.');
  const studentAnswers={},teacherAnswers={};
  DEFAULT_CRITERIA.forEach((criterion,index)=>{studentAnswers[criterion.id]=index%2+3;teacherAnswers[criterion.id]=4;});
  return buildResultMessage({
    student:{name:'Alex Beispiel'},
    round:{title:'Beispielrunde',criteria:DEFAULT_CRITERIA},
    assessment:{studentEmail:recipient,studentAnswers,teacherAnswers,studentAdditionalQuality:'Ich habe die Gruppenarbeit unterstützt.',studentComment:'Ich möchte mich häufiger melden.',teacherComment:'Du arbeitest aufmerksam mit. Bringe deine Ideen weiterhin aktiv ein.'}
  });
}
module.exports={buildDebugMessage};
