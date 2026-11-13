const {randomUUID}=require('node:crypto');

function createSessionCloseGuard({app,ipcMain,getActiveSession}) {
 const windows=new Set(),pending=new Map();let approved=false;
 function ask(window,event) {
  if(approved)return;const session=getActiveSession();if(!session)return;
  event?.preventDefault();
  if(window.isDestroyed())return;
  if(!pending.has(window)) {
   const request={requestId:randomUUID(),sessionId:session.id};pending.set(window,request);
   window.webContents.send('session-close-confirmation',{requestId:request.requestId,className:session.className});
  }
  window.show();window.focus();
 }
 ipcMain.on('session-close-cancel',(event,requestId)=>{
  for(const [window,request]of pending)if(window.webContents===event.sender && request.requestId===requestId)pending.delete(window);
 });
 ipcMain.on('session-close-confirm',(event,requestId)=>{
  for(const [window,request]of pending) {
   if(window.webContents!==event.sender || request.requestId!==requestId)continue;
   const active=getActiveSession();pending.delete(window);
   if(active && active.id!==request.sessionId){ask(window);return;}
   approved=true;app.quit();return;
  }
 });
 app.on('before-quit',event=>{
  if(approved || !getActiveSession())return;
  const window=[...windows].find(window=>!window.isDestroyed());
  if(window)ask(window,event);
 });
 return {attach(window){
  windows.add(window);window.on('close',event=>ask(window,event));
  window.once('closed',()=>{windows.delete(window);pending.delete(window);});
 }};
}
module.exports={createSessionCloseGuard};
