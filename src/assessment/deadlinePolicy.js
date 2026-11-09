const DAY = 86400000;
const MAX_DEADLINE_DAYS = 90;
function zurichDate(time = Date.now()) {
  return new Intl.DateTimeFormat('en-CA', {timeZone:'Europe/Zurich',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date(time));
}
function deadlineEndZurich(date) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date)) || !Number.isFinite(Date.parse(`${date}T12:00:00Z`)) || new Date(`${date}T12:00:00Z`).toISOString().slice(0,10) !== date) throw invalid('Die Abgabefrist ist ungültig.');
  const midnight = Date.parse(`${date}T23:59:59Z`);
  const offset = new Intl.DateTimeFormat('en-US', {timeZone:'Europe/Zurich',timeZoneName:'longOffset'}).formatToParts(new Date(midnight)).find(p=>p.type==='timeZoneName').value.match(/GMT\+(\d+)/);
  return new Date(midnight-Number(offset[1])*3600000).toISOString();
}
function deadlineBounds(now = Date.now()) {
  const limit = now + MAX_DEADLINE_DAYS*DAY;
  let max = zurichDate(limit);
  if (Date.parse(deadlineEndZurich(max)) > limit) max = zurichDate(limit-DAY);
  return {min:zurichDate(now),max};
}
function validateDeadline(value, now = Date.now(), required = false) {
  if (!value && !required) return null;
  const date = String(value || '').trim();
  const time = Date.parse(deadlineEndZurich(date));
  if (time <= now) throw invalid('Die Abgabefrist liegt in der Vergangenheit. Bitte ein neues Datum wählen.');
  if (time > now+MAX_DEADLINE_DAYS*DAY) throw invalid('Die Abgabefrist darf höchstens 90 Tage in der Zukunft liegen. Bitte ein früheres Datum wählen.');
  return date;
}
function invalid(message) {return Object.assign(new Error(message),{code:'VALIDATION_ERROR'});}
module.exports = {DAY,MAX_DEADLINE_DAYS,zurichDate,deadlineEndZurich,deadlineBounds,validateDeadline};
