// Separate Node process: a blocked Electron main loop cannot stop this monitor.
const fs=require('node:fs'),path=require('node:path');
const {createCrashJournal}=require('./crashJournal');
function isStale(now,lastTick,threshold=45000){return now-lastTick>threshold;}
function monitor(dir,runId,pid,{pollMs=5000,stallMs=45000}={}){
  const journal=createCrashJournal(dir),heartbeat=path.join(dir,`${runId}.heartbeat`);let reported=false,lastCheck=Date.now();
  const timer=setInterval(()=>{
    try{process.kill(pid,0);}catch{clearInterval(timer);return;}
    try{
      const now=Date.now(),sleepGap=now-lastCheck;lastCheck=now;
      const age=now-fs.statSync(heartbeat).mtimeMs;
      // A sleeping computer also pauses this monitor: wait a fresh interval on wake.
      if(sleepGap>Math.max(15000,pollMs*3)){fs.utimesSync(heartbeat,new Date(),new Date());return;}
      if(isStale(Date.now(),fs.statSync(heartbeat).mtimeMs,stallMs)&&!reported){journal.incident(runId,'main-unresponsive','Der Hauptprozess antwortet seit mindestens 45 Sekunden nicht mehr.');reported=true;}
      if(age<Math.min(10000,stallMs/2))reported=false;
    }catch{/* Parent may be shutting down. */}
  },pollMs);
  return timer;
}
if(require.main===module)monitor(process.argv[2],process.argv[3],Number(process.argv[4]));
module.exports={isStale,monitor};
