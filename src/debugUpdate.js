function createDebugUpdate({enabled,version,send,restart,delay=ms=>new Promise(resolve=>setTimeout(resolve,ms))}) {
  let mode=null,busy=false;
  const emit=(status,data={})=>send(status,{version,simulated:true,...data});
  return {
    show(requestedMode) {
      if(!enabled || busy || !['preview','restart'].includes(requestedMode))return false;
      mode=requestedMode;
      emit('available',{simulationMode:mode});return true;
    },
    cancel(){if(!busy)mode=null;},
    async approve() {
      if(!enabled || !mode)return false;
      if(busy)return true;
      busy=true;
      try {
        if(mode==='preview'){emit('simulated-preview');return true;}
        emit('downloading',{percent:0});
        for(const percent of [25,60,100]){await delay(500);emit('progress',{percent});}
        emit('downloaded');await delay(500);emit('installing');await delay(700);
        await restart();return true;
      } finally {mode=null;busy=false;}
    }
  };
}
module.exports={createDebugUpdate};
