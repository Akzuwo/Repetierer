const assert=require('node:assert/strict'),{ipcMain}=require('electron');
module.exports=async function({win,evaluate,waitFor,pause}){
 let answer,requests=0;ipcMain.handle('ai:ask',()=>{requests++;return new Promise(resolve=>{answer=resolve;});});
 win.showInactive();await waitFor(`document.visibilityState==='visible'`);
 await evaluate(`document.getElementById('ai-open-btn').click()`);
 await waitFor(`!document.getElementById('ai-input').disabled`);await pause(450);
 await evaluate(`window.earlyLoader=false;window.loaderObserver=new MutationObserver(()=>{if(document.querySelector('.ai-send-flight')&&!document.getElementById('ai-activity').classList.contains('update-hidden'))window.earlyLoader=true;});loaderObserver.observe(document.getElementById('ai-activity'),{attributes:true,attributeFilter:['class']})`);
 async function send(text){answer=null;const state=await evaluate(`(()=>{const i=document.getElementById('ai-input');i.value=${JSON.stringify(text)};i.dispatchEvent(new Event('input'));document.getElementById('ai-form').requestSubmit();return {flying:!!document.querySelector('.ai-send-flight'),hidden:document.getElementById('ai-activity').classList.contains('update-hidden')};})()`);for(let i=0;!answer&&i<100;i++)await pause(5);assert.ok(answer,'Request starts without waiting for flight');return state;}
 const fast=await send('Schnelle Antwort testen');assert.equal(fast.flying,true);assert.equal(fast.hidden,true);
 answer({type:'answer',text:'Schnelle Antwort.'});await pause(50);
 assert.equal(await evaluate(`document.querySelectorAll('.ai-assistant').length`),0,'A fast answer waits for landing');
 await waitFor(`!!document.querySelector('.ai-assistant') && document.getElementById('ai-activity').dataset.busy==='false'`);
 const slow=await send('Langsame Antwort testen');assert.equal(slow.hidden,true);
 await waitFor(`!document.querySelector('.ai-send-flight')&&!document.getElementById('ai-activity').classList.contains('update-hidden')`);
 answer({type:'answer',text:'Langsame Antwort.'});await waitFor(`document.getElementById('ai-activity').dataset.busy==='false'`);
 await send('Bewegungsreduktion während des Flugs');
 await win.webContents.debugger.sendCommand('Emulation.setEmulatedMedia',{features:[{name:'prefers-reduced-motion',value:'reduce'}]});
 await waitFor(`!document.querySelector('.ai-send-flight')&&!document.getElementById('ai-activity').classList.contains('update-hidden')`);
 answer({type:'answer',text:'Ohne Flug.'});await waitFor(`document.getElementById('ai-activity').dataset.busy==='false'`);
 const reduced=await send('Direkt senden');assert.equal(reduced.flying,false);assert.equal(reduced.hidden,false);answer({type:'answer',text:'Direkte Antwort.'});await waitFor(`document.getElementById('ai-activity').dataset.busy==='false'`);
 await win.webContents.debugger.sendCommand('Emulation.setEmulatedMedia',{features:[{name:'prefers-reduced-motion',value:'no-preference'}]});
 await send('Chat während des Flugs schliessen');await evaluate(`document.getElementById('ai-close-btn').click()`);
 await waitFor(`!document.querySelector('.ai-send-flight')&&getComputedStyle(document.getElementById('ai-modal')).display==='none'`);
 answer({type:'answer',text:'Geschlossen.'});await waitFor(`document.getElementById('ai-activity').dataset.busy==='false'`);
 assert.equal(await evaluate(`window.earlyLoader`),false);assert.equal(requests,5);
 await evaluate(`window.loaderObserver.disconnect()`);ipcMain.removeHandler('ai:ask');
 console.log('AI_FLIGHT_UI_PASS: dots only after landing, immediate backend request, fast/slow answers, live reduced motion, no-flight path and interrupted close');
};
