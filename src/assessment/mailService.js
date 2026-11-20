const fs=require('node:fs'),path=require('node:path'),crypto=require('node:crypto');
const sender=require('./senderIdentity');
function mailError(message,code='MAIL_SERVICE_ERROR'){return Object.assign(new Error(message),{code});}
const smtpQueues=new Map();
async function resolveMailService({client,smtpConfig,getSmtpConfig,createSmtpMailer,profile,localDailyLimit=200}) {
  if(client){
    const capabilities=await client.mailCapabilities();
    if(capabilities.configured===true){
      return {mode:'cloud',configured:true,...sender,...capabilities,verify:async()=>true,send:message=>client.sendMail(message,message.idempotencyKey||crypto.randomUUID())};
    }
    if(capabilities.configured!==false && !capabilities.legacy)throw mailError('Der Server hat keinen eindeutigen Mailstatus geliefert.');
  }
  smtpConfig=smtpConfig||getSmtpConfig?.();
  if(!smtpConfig)throw mailError('Lokaler SMTP-Versand ist noch nicht eingerichtet.','BREVO_NOT_CONFIGURED');
  const mailer=createSmtpMailer(smtpConfig);
  const quotaFile=path.join(profile,'smtp-mail-usage.json');
  return {mode:'smtp',configured:true,...sender,dailyLimit:localDailyLimit,verify:()=>mailer.verify(),send:message=>{
    const action=(smtpQueues.get(profile)||Promise.resolve()).then(async()=>{
      const day=new Date().toISOString().slice(0,10);let quota={day,used:0};
      try {const saved=JSON.parse(fs.readFileSync(quotaFile,'utf8'));if(saved.day===day)quota=saved;}catch(error){if(error.code!=='ENOENT')throw mailError('Das lokale Mailkontingent konnte nicht gelesen werden.');}
      if(quota.used>=localDailyLimit)throw mailError('Das lokale Tageslimit ist erreicht.','MAIL_DAILY_LIMIT');
      // Reserve before I/O. Uncertain/failed sends remain counted to avoid oversending.
      fs.writeFileSync(quotaFile,JSON.stringify({day,used:quota.used+1}),{mode:0o600});
      return mailer.send(message);
    });smtpQueues.set(profile,action.catch(()=>{}));return action;
  }};
}
module.exports={resolveMailService,mailError};
