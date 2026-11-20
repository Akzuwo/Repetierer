const HEALTH_INTERVAL_MS=5*60*1000;
function createHealthMonitor({check,onUpdate=()=>{},setIntervalFn=setInterval,clearIntervalFn=clearInterval}){
 let state=null,running=null,timer=null,revision=0,stopped=false;
 async function refresh(){
  if(stopped)return state;
  if(running)return running;
  const current=revision;
  running=(async()=>{
   let result;
   try{result=await check();}
   catch{result={connection:{reachable:false,error:'Die Verbindung konnte nicht geprüft werden.'},brevo:{configured:false,invalid:true,error:'Die Verbindung konnte nicht geprüft werden.'}};}
   if(!stopped&&current===revision){state={...result,checkedAt:Date.now()};onUpdate(state);}
  })().finally(()=>{running=null;}).then(()=>!stopped&&current!==revision?refresh():state);
  return running;
 }
 return {
  start(){if(stopped||timer!==null)return;timer=setIntervalFn(()=>refresh(),HEALTH_INTERVAL_MS);timer?.unref?.();return refresh();},
  get(){return state?Promise.resolve(state):refresh();},
  invalidate(){revision++;state=null;return refresh();},
  stop(){stopped=true;if(timer!==null)clearIntervalFn(timer);timer=null;}
 };
}
module.exports={HEALTH_INTERVAL_MS,createHealthMonitor};
