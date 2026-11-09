const SYNC_INTERVAL_MS = 30000;
// The same queue protects manual actions and timer-driven reads from crossing.
function createAutoSync({getTargets,sync,onResult=()=>{},onError=()=>{},intervalMs=SYNC_INTERVAL_MS,setIntervalFn=setInterval,clearIntervalFn=clearInterval}) {
  const queues = new Map();
  let running = false, stopped = false;
  const key = target => JSON.stringify([target.filePath,target.className,target.roundId]);
  function run(target,action) {
    const id=key(target), previous=queues.get(id) || Promise.resolve();
    const task=previous.catch(()=>{}).then(action);
    queues.set(id,task);
    task.finally(()=>{if(queues.get(id)===task)queues.delete(id);}).catch(()=>{});
    return task;
  }
  async function tick() {
    if(running || stopped)return;
    running=true;
    try {
      for(const target of await getTargets()) {
        if(stopped)break;
        // A slow request is retried at the next interval, never overlapped.
        if(queues.has(key(target)))continue;
        try {const result=await run(target,()=>sync(target));if(!stopped)await onResult(target,result);}
        catch(error){if(!stopped)onError(target,error);}
      }
    } finally {running=false;}
  }
  const timer=setIntervalFn(()=>tick().catch(error=>onError(null,error)),intervalMs);timer?.unref?.();
  return {run,tick,stop(){stopped=true;clearIntervalFn(timer);}};
}
module.exports={createAutoSync,SYNC_INTERVAL_MS};
