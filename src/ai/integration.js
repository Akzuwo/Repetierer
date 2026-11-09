function createAiIntegration({service,getEnabled,setEnabled}) {
  let settingUp = false;
  function requireEnabled() {
    if(!getEnabled())throw new Error('Die KI-Integration ist entfernt. Du kannst sie in den Einstellungen wieder einrichten.');
  }
  return {
    requireEnabled,
    async status(){return getEnabled() ? {...await service.status(),integrationEnabled:true} : {state:'disabled',integrationEnabled:false,message:'KI-Integration entfernt'};},
    async setup(onProgress){
      if(settingUp)throw new Error('Die Einrichtung läuft bereits.');
      settingUp=true;
      try {const result=await service.setup(onProgress);if(result.state==='ready')setEnabled(true);return result;}
      finally {settingUp=false;}
    },
    remove(){if(settingUp)throw new Error('Bitte die laufende KI-Einrichtung zuerst abschliessen.');setEnabled(false);return {state:'disabled',integrationEnabled:false,message:'Die KI-Integration wurde entfernt. Ollama und vorhandene Modelle bleiben auf dem Gerät.'};},
    async chat(...args){requireEnabled();const result=await service.chat(...args);requireEnabled();return result;}
  };
}
module.exports={createAiIntegration};
