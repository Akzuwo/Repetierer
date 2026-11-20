const crypto=require('node:crypto');
const {isValidEmail}=require('./assessment/matching');
const {sanitizeMailError}=require('./assessment/assessmentDelivery');
const digest=value=>crypto.createHash('sha256').update(JSON.stringify(value)).digest('hex');
const identity=record=>JSON.stringify([record.className,record.id,record.name]);
function selectRecords(data,args){
  if(!['student','class','all'].includes(args.scope))throw new Error('Ungültige Aktenauswahl.');
  const records=data.classes.filter(c=>args.scope==='all'||c.name===args.className).flatMap(c=>c.students).filter(d=>args.scope!=='student'||d.id===args.studentId&&d.name===args.studentName);
  if(!records.length)throw new Error('Keine Schüler für diese Auswahl gefunden.');
  return records;
}
function createDossierMail({load,requireContext,readConfig,createMailer,renderPdf,safeName,isBusy,setBusy}){
  const plans=new Map(),retries=new Map();
  function remember(map,value){
    for(const [key,item] of map)if(item.expires<Date.now())map.delete(key);
    while(map.size>=20)map.delete(map.keys().next().value);
    const key=crypto.randomUUID();map.set(key,{...value,expires:Date.now()+10*60*1000});return key;
  }
  async function prepare(args){
    if(isBusy())throw new Error('Eine Aktenaktion läuft bereits.');
    const data=await load();requireContext(data,args);await readConfig();
    let records=selectRecords(data,args);
    if(args.retryToken){
      const retry=retries.get(args.retryToken);
      if(!retry||retry.expires<Date.now()||retry.filePath!==data.filePath)throw new Error('Die Wiederholung ist abgelaufen. Bitte die Auswahl erneut prüfen.');
      records=records.filter(d=>retry.identities.includes(identity(d)));
      if(records.length!==retry.identities.length)throw new Error('Die Schülerzuordnung hat sich geändert. Bitte die Schülerakten erneut öffnen.');
    }
    const missing=records.filter(d=>!isValidEmail(d.email)||d.email.length>254||/[<>,;:"\\\r\n]/.test(d.email));
    const token=missing.length?null:remember(plans,{args:{...args},hash:digest(records),identities:records.map(identity)});
    return {token,recipients:records.map(d=>({name:d.name,className:d.className,email:d.email})),missing:missing.map(d=>({name:d.name,className:d.className}))};
  }
  async function send(token){
    if(isBusy())throw new Error('Eine Aktenaktion läuft bereits.');
    const plan=plans.get(token);
    if(!plan||plan.expires<Date.now())throw new Error('Die Versandvorschau ist abgelaufen. Bitte die Auswahl erneut prüfen.');
    plans.delete(token);setBusy(true);
    let records=[],results=[];
    try{
      const data=await load();requireContext(data,plan.args);
      records=selectRecords(data,plan.args).filter(d=>plan.identities.includes(identity(d)));
      if(digest(records)!==plan.hash)throw new Error('Die Akten oder Adressen wurden geändert. Bitte die Versandvorschau erneut öffnen.');
      const config=await readConfig(),mailer=createMailer(config);
      for(const record of records){
        const current=await load();requireContext(current,plan.args);
        const fresh=selectRecords(current,plan.args).filter(d=>plan.identities.includes(identity(d)));
        if(digest(fresh)!==plan.hash)throw new Error('Die Akten oder Adressen wurden während des Versands geändert.');
        try{
          const content=await renderPdf(record);
          // Recheck after asynchronous PDF rendering, before handing anything to SMTP.
          const latest=await load();requireContext(latest,plan.args);
          if(digest(selectRecords(latest,plan.args).filter(d=>plan.identities.includes(identity(d))))!==plan.hash)throw Object.assign(new Error('Die Akten oder Adressen wurden geändert.'),{code:'DOSSIER_CHANGED'});
          const result=await mailer.send({from:{name:config.fromName,address:config.fromEmail},to:record.email,subject:`Deine Schülerakte · ${record.className}`,text:`Guten Tag ${record.name}\n\nIm Anhang findest du deine persönliche Schülerakte als PDF.\n\nFreundliche Grüsse\n${config.fromName}`,attachments:[{filename:safeName(record.name)+'.pdf',content,contentType:'application/pdf'}]});
          if(result?.rejected?.length||result?.accepted&&result.accepted.length===0)throw new Error('SMTP hat die Adresse abgelehnt.');
          results.push({identity:identity(record),name:record.name,className:record.className,status:'sent'});
        }catch(error){
          if(error.code==='DOSSIER_CHANGED')throw error;
          results.push({identity:identity(record),name:record.name,className:record.className,status:'failed',error:sanitizeMailError(error)});
        }
      }
      return finish();
    }catch(error){return finish(error.message);}
    finally{setBusy(false);}
    function finish(error){
      const sent=new Set(results.filter(r=>r.status==='sent').map(r=>r.identity));
      const remaining=records.filter(d=>!sent.has(identity(d)));
      const retryToken=remaining.length?remember(retries,{filePath:plan.args.filePath,identities:remaining.map(identity)}):null;
      return {sent:sent.size,failed:results.filter(r=>r.status==='failed').length,remaining:remaining.length,results:results.map(({identity,...r})=>r),retryToken,error};
    }
  }
  return {prepare,send};
}
module.exports={createDossierMail,selectRecords};
