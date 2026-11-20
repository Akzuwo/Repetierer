const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const DEFAULT_PORT = 17839;
const filenames = {settings:'settings.json', backup:'backup.json', sessions:'unterrichtssessions.json', assessments:'mitmachnoten.json', migrations:'migrations.json'};
function read(profile, name, fallback = {}) {
  try { return JSON.parse(fs.readFileSync(path.join(profile, filenames[name]), 'utf8')); }
  catch (error) { if (error.code === 'ENOENT') return fallback; throw Object.assign(new Error('Daten werden gerade gespeichert oder sind ungültig. Erneut versuchen.'), {status:503}); }
}
function credentials(profile) {
  fs.mkdirSync(profile,{recursive:true});
  const file=path.join(profile,'local-api-token');
  try {fs.writeFileSync(file,crypto.randomBytes(32).toString('base64url'),{flag:'wx',mode:0o600});} catch(error){if(error.code!=='EEXIST')throw error;}
  const token=fs.readFileSync(file,'utf8').trim();
  if(!/^[A-Za-z0-9_-]{43}$/.test(token))throw new Error('Ungültiger lokaler API-Schlüssel.');
  return token;
}
// Credential material is never part of any API result, including future nested settings.
function redact(value) {
  if(Array.isArray(value))return value.map(redact);
  if(value && typeof value==='object')return Object.fromEntries(Object.entries(value).filter(([key])=>!/token|password|secret|credential|api.?key|smtp|brevo|authorization/i.test(key)).map(([key,v])=>[key,redact(v)]));
  return value;
}
async function workbook(profile) {
  const filePath=read(profile,'settings').excelFilePath;
  if(!filePath)return {filePath:null,classes:[]};
  const {loadDossiers}=require('../studentDossiers');
  const sessions=read(profile,'sessions',{sessions:[]}).sessions||[];
  const pending=read(profile,'backup').pendingExcelEntries||[];
  return loadDossiers({filePath,paths:{assessmentPath:path.join(profile,filenames.assessments)},sessions,pending});
}
async function startServer({profile,port=DEFAULT_PORT,checkEnabled=true,onClose=()=>{}}) {
  const token=credentials(profile);
  const server=http.createServer(async(req,res)=>{
    const reply=(status,data)=>{res.writeHead(status,{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store','X-Content-Type-Options':'nosniff'});res.end(JSON.stringify(data));};
    try {
      if(req.headers.origin || !['127.0.0.1','localhost'].includes((req.headers.host||'').split(':')[0]))return reply(403,{error:'origin_denied'});
      const supplied=Buffer.from(req.headers.authorization?.replace(/^Bearer /,'')||''),expected=Buffer.from(token);
      if(supplied.length!==expected.length || !crypto.timingSafeEqual(supplied,expected))return reply(401,{error:'unauthorized'});
      if(req.method!=='GET'){res.setHeader('Allow','GET');return reply(405,{error:'read_only'});}
      const url=new URL(req.url,'http://127.0.0.1'),route=url.pathname;
      let data;
      if(route==='/v1/health')data={service:'repetierer-local-api',version:1,readOnly:true};
      else if(route==='/v1')data={endpoints:['health','settings','backup','sessions','assessments','migrations','workbook','classes','students','snapshot'].map(p=>'/v1/'+p)};
      else if(Object.hasOwn(filenames,route.slice(4)) && route.startsWith('/v1/'))data=read(profile,route.slice(4));
      else if(['/v1/workbook','/v1/classes','/v1/students'].includes(route)) {
        const book=await workbook(profile);
        data=route.endsWith('classes')?book.classes.map(c=>({name:c.name,studentCount:c.students.length})):route.endsWith('students')?book.classes.filter(c=>!url.searchParams.has('class')||c.name===url.searchParams.get('class')).flatMap(c=>c.students):book;
      } else if(route==='/v1/snapshot') {
        data=Object.fromEntries(Object.keys(filenames).map(k=>[k,read(profile,k)]));data.workbook=await workbook(profile);
      } else return reply(404,{error:'not_found'});
      return reply(200,{success:true,readAt:new Date().toISOString(),data:redact(data)});
    }catch(error){reply(error.status||503,{error:'data_unavailable',message:'Die Daten konnten nicht gelesen werden. Erneut versuchen.'});}
  });
  server.requestTimeout=10000;server.headersTimeout=10000;server.maxConnections=16;
  await new Promise((resolve,reject)=>{server.once('error',reject);server.listen(port,'127.0.0.1',resolve);});
  const info=path.join(profile,'local-api.json');
  fs.writeFileSync(info,JSON.stringify({pid:process.pid,url:`http://127.0.0.1:${server.address().port}/v1`,tokenFile:path.join(profile,'local-api-token')}),{mode:0o600});
  let watcher,timer,closed=false;
  const close=()=>{if(closed)return;closed=true;watcher?.close();clearInterval(timer);server.closeAllConnections();server.close(onClose);};
  const check=()=>{try{if(read(profile,'settings').localApiEnabled===false)close();}catch{ /* Never change app data while a writer is active. */ }};
  if(checkEnabled){watcher=fs.watch(profile,(_,file)=>{if(file==='settings.json')check();});timer=setInterval(check,30000);timer.unref();check();}
  return {server,close,url:`http://127.0.0.1:${server.address()?.port}/v1`};
}
module.exports={startServer,DEFAULT_PORT,redact,credentials};
