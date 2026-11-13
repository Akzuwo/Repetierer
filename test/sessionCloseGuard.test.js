const {test}=require('node:test'),assert=require('node:assert/strict'),{EventEmitter}=require('node:events');
const {createSessionCloseGuard}=require('../src/sessionCloseGuard');
function fixture(){
 const app=new EventEmitter(),ipcMain=new EventEmitter();let active={id:'session-a',className:'Klasse A'},quits=0;
 app.quit=()=>quits++;
 const win=new EventEmitter(),messages=[];
 Object.assign(win,{isDestroyed:()=>false,show(){},focus(){},webContents:{send:(...args)=>messages.push(args)}});
 const guard=createSessionCloseGuard({app,ipcMain,getActiveSession:()=>active});guard.attach(win);
 const event=()=>({prevented:false,preventDefault(){this.prevented=true;}});
 return {app,ipcMain,win,messages,event,setActive:value=>active=value,quits:()=>quits};
}
test('native close and app quit require one session confirmation',()=>{
 const f=fixture(),native=f.event();f.win.emit('close',native);assert.equal(native.prevented,true);
 const quit=f.event();f.app.emit('before-quit',quit);assert.equal(quit.prevented,true);assert.equal(f.messages.length,1);
 const request=f.messages[0][1];assert.equal(request.className,'Klasse A');
 f.ipcMain.emit('session-close-confirm',{sender:{}},request.requestId);assert.equal(f.quits(),0);
 f.ipcMain.emit('session-close-confirm',{sender:f.win.webContents},'wrong');assert.equal(f.quits(),0);
 f.ipcMain.emit('session-close-confirm',{sender:f.win.webContents},request.requestId);assert.equal(f.quits(),1);
 const approved=f.event();f.app.emit('before-quit',approved);assert.equal(approved.prevented,false);
});
test('cancel leaves the session active and requires a fresh confirmation',()=>{
 const f=fixture();f.win.emit('close',f.event());const old=f.messages[0][1].requestId;
 f.ipcMain.emit('session-close-cancel',{sender:f.win.webContents},old);
 f.ipcMain.emit('session-close-confirm',{sender:f.win.webContents},old);assert.equal(f.quits(),0);
 f.win.emit('close',f.event());assert.equal(f.messages.length,2);assert.notEqual(f.messages[1][1].requestId,old);
});
test('another session invalidates an old confirmation; no session closes directly',()=>{
 const f=fixture();f.win.emit('close',f.event());const old=f.messages[0][1].requestId;
 f.setActive({id:'session-b',className:'Klasse B'});
 f.ipcMain.emit('session-close-confirm',{sender:f.win.webContents},old);assert.equal(f.quits(),0);assert.equal(f.messages[1][1].className,'Klasse B');
 f.setActive(null);const event=f.event();f.win.emit('close',event);assert.equal(event.prevented,false);
});
