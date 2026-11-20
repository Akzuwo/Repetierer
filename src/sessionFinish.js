// Shared main-process validation for ending a session from the UI or before quit.
function createSessionFinish({getSession,getSelection,isSaving=()=>false,endSession,onEnded=()=>{}}) {
 function snapshot() {
  const session=getSession(),selection=getSelection();
  return {sessionId:session?.id||null,className:session?.className||'',selectionToken:selection?.selectionToken||null,
   studentName:selection?.selectionToken?selection.name:'',saving:isSaving(),
   hasGrade:!!session?.repetitionEntries?.some(entry=>entry.type==='grade'&&!entry.undone&&Number.isFinite(entry.grade))};
 }
 function finish(request={}) {
  const state=snapshot();
  if(!state.sessionId||request.sessionId!==state.sessionId)return {success:false,reason:'session-changed'};
  if(state.saving)return {success:false,reason:'save-in-progress'};
  if((request.selectionToken||null)!==state.selectionToken)return {success:false,reason:'selection-changed'};
  if(state.selectionToken&&request.discardRepetition!==true)return {success:false,reason:'repetition-open'};
  if(!state.hasGrade&&request.allowEmpty!==true)return {success:false,reason:'no-repetition'};
  try {
   const result=endSession(state.sessionId);
   if(!result?.success)return {success:false,reason:result?.reason||'session-end-failed'};
   onEnded(result.session);
   return {success:true,sessionId:state.sessionId};
  } catch {return {success:false,reason:'session-end-failed'};}
 }
 return {snapshot,finish};
}
module.exports={createSessionFinish};
