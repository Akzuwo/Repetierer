function formatBuildDate(date = new Date()) {
  const parts = new Intl.DateTimeFormat('de-CH',{timeZone:'Europe/Zurich',day:'2-digit',month:'2-digit',year:'numeric'}).formatToParts(date);
  return ['day','month','year'].map(type=>parts.find(part=>part.type===type).value).join('');
}
function validBuildDate(value) {
  if(typeof value!=='string' || !/^\d{8}$/.test(value))return false;
  const date=new Date(`${value.slice(4)}-${value.slice(2,4)}-${value.slice(0,2)}T12:00:00Z`);
  return Number.isFinite(date.getTime()) && formatBuildDate(date)===value;
}
function getBuildInfo(metadata = {}) {
  const number=Number(metadata.buildNumber ?? metadata.localBuildNumber);
  const buildNumber=Number.isSafeInteger(number) && number>0 ? number : null;
  const buildDate=validBuildDate(metadata.buildDate) ? metadata.buildDate : null;
  return {version:String(metadata.version || ''),buildNumber,buildDate,build:buildNumber ? `${buildNumber}${buildDate ? '.'+buildDate : ''}` : 'unbekannt'};
}
function createBuildMetadata(metadata, env = process.env, date = new Date()) {
  const buildNumber=Number(env.LOCAL_BUILD_NUMBER || metadata.buildNumber);
  const buildDate=env.LOCAL_BUILD_DATE || formatBuildDate(date);
  if(!Number.isSafeInteger(buildNumber) || buildNumber<1 || !validBuildDate(buildDate))throw new Error('Ungültige Buildnummer oder ungültiges Builddatum.');
  return getBuildInfo({...metadata,buildNumber,buildDate});
}
module.exports={formatBuildDate,validBuildDate,getBuildInfo,createBuildMetadata};
