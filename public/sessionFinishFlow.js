// Coordinates decisions only. Persistence and stale-state checks remain in main.
function createSessionFinishFlow({ask,getState,finishSession,editGrade,confirmQuit,cancelQuit,reportError}){
 let busy=false,waiting=null,busyQuit=null,deferredQuit=null;
 async function run({quitRequest=null,resume=false,expectedSessionId=null}={}){
  if(busy){if(quitRequest&&quitRequest!==busyQuit)deferredQuit=quitRequest;return;}
  if(!resume&&waiting){if(waiting.quitRequest===quitRequest)waiting=null;else cancelWaiting();}
  busy=true;busyQuit=quitRequest;
  let keepRequest=false;
  try{
   let state=await getState();
   if(expectedSessionId&&state.sessionId!==expectedSessionId)return;
   if(state.saving){reportError('Die Bewertung wird noch gespeichert. Bitte danach erneut versuchen.');return;}
   if(!state.sessionId){if(quitRequest)confirmQuit(quitRequest);return;}
   if(quitRequest&&!resume){
    const answer=await ask({title:'Session beenden und schliessen?',content:`Die Session der Klasse ${state.className} läuft noch.`,buttons:['Zurück zur App','Beenden und schliessen'],highlight:2,focus:1});
    if(answer.action!=='button2')return;
   }
   let discard=false,allowEmpty=false;
   if(state.selectionToken){
    const answer=await ask({title:'Noch eine Note speichern?',content:`Die Note für ${state.studentName} wurde noch nicht gespeichert.`,buttons:['Note erfassen','Ohne Note beenden'],highlight:1,focus:1,destructive:2});
    if(answer.action==='cancel')return;
    if(answer.action==='button1'){
     waiting={quitRequest,sessionId:state.sessionId,selectionToken:state.selectionToken};keepRequest=true;editGrade();return;
    }
    discard=true;
   }
   if(!state.hasGrade){
    const answer=await ask({title:'Session beenden?',content:'In dieser Session wurde noch keine Repetition abgeschlossen.',buttons:['Session fortsetzen','Trotzdem beenden'],highlight:1,focus:1,destructive:2});
    if(answer.action!=='button2')return;allowEmpty=true;
   }else if(!quitRequest&&!resume&&!discard){
    const answer=await ask({title:'Session beenden?',content:'Die gespeicherten Einträge bleiben erhalten.',buttons:['Session fortsetzen','Session beenden'],highlight:1,focus:1,destructive:2});
    if(answer.action!=='button2')return;
   }
   const result=await finishSession({sessionId:state.sessionId,selectionToken:state.selectionToken,discardRepetition:discard,allowEmpty});
   if(!result?.success){reportError('Die Session konnte nicht beendet werden. Bitte den aktuellen Stand prüfen und erneut versuchen.');return;}
   if(quitRequest){keepRequest=true;confirmQuit(quitRequest);}
  }catch{reportError('Der Beendigungsvorgang konnte nicht abgeschlossen werden. Deine Eingaben bleiben erhalten.');}
  finally{busy=false;busyQuit=null;if(quitRequest&&!keepRequest)cancelQuit(quitRequest);if(deferredQuit){const next=deferredQuit;deferredQuit=null;run({quitRequest:next});}}
 }
 function cancelWaiting(){if(waiting?.quitRequest)cancelQuit(waiting.quitRequest);waiting=null;}
 return {
  run,
  async saved(result){
   const pending=waiting;if(!pending)return;if(result?.type!=='grade'){cancelWaiting();return;}
   waiting=null;
   if(result.sessionId!==pending.sessionId||result.selectionToken!==pending.selectionToken){if(pending.quitRequest)cancelQuit(pending.quitRequest);return;}
   return run({quitRequest:pending.quitRequest,resume:true,expectedSessionId:pending.sessionId});
  },
  cancelWaiting,
  failed:cancelWaiting,
  get pending(){return !!waiting;}
 };
}
module.exports={createSessionFinishFlow};
