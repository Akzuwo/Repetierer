const test = require('node:test'), assert = require('node:assert/strict');
const {normalizeDuration,normalizeNotification,inferType} = require('../src/notificationPolicy');
test('notification duration defaults to six seconds and is bounded to two through eight',()=>{
 for(const value of [undefined,null,'',NaN,Infinity,'invalid']) assert.equal(normalizeDuration(value),6);
 assert.equal(normalizeDuration(-10),2);assert.equal(normalizeDuration(99),8);
 assert.equal(normalizeDuration('4'),4);assert.equal(normalizeDuration(4.7),5);
 assert.equal(normalizeNotification('success','Saved',{},8).durationSeconds,8);
 assert.equal(normalizeNotification('info','Text',{durationSeconds:2},8).durationSeconds,2);
});
test('notification payloads preserve plain text and normalize type, errors and empty input',()=>{
 assert.equal(normalizeNotification('error','   '),null);
 assert.equal(normalizeNotification('error',new Error('Failure')).text,'Failure');
 assert.equal(normalizeNotification('warn','Text').type,'warning');
 assert.equal(normalizeNotification('unknown','Text').type,'info');
 assert.equal(normalizeNotification('info','<script>unsafe</script>').text,'<script>unsafe</script>');
 assert.equal(normalizeNotification('info','x'.repeat(2000)).text.length,1600);
 assert.equal(inferType('Nicht gespeichert','Erneut versuchen'),'error');
 assert.equal(inferType('Speichern ausstehend','Excel ist gesperrt'),'warning');
 assert.equal(inferType('Excel aktualisiert','Fertig'),'success');
 assert.equal(inferType('Excel-Datei ist geöffnet','Der Eintrag wurde lokal gespeichert.'),'warning');
});
test('settings persist duration independently and existing settings receive a six-second default',()=>{
 const fs=require('node:fs'),os=require('node:os'),path=require('node:path'),Module=require('node:module');
 const dir=fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-notify-')),load=Module._load;
 Module._load=function(request,parent,main){return request==='electron'?{app:{getPath:()=>dir}}:load.call(this,request,parent,main);};
 try {
  const storage=require('../src/storage');
  assert.equal(storage.getAppSettings().notificationDurationSeconds,6);
  storage.saveAppSettings({...storage.getAppSettings(),notificationDurationSeconds:8});
  assert.equal(storage.getAppSettings().notificationDurationSeconds,8);
  storage.saveAppSettings({...storage.getAppSettings(),notificationDurationSeconds:-1});
  assert.equal(storage.getAppSettings().notificationDurationSeconds,2);
  assert.equal(JSON.parse(fs.readFileSync(storage.getPaths().settingsPath)).notificationDurationSeconds,2);
 }finally{Module._load=load;fs.rmSync(dir,{recursive:true,force:true});delete require.cache[require.resolve('../src/storage')];}
});
test('main notifications queue before startup, bound bursts and replace keyed progress',()=>{
 const Module=require('node:module'),{EventEmitter}=require('node:events'),load=Module._load;
 Module._load=function(request,parent,main){return request==='./storage' && parent.filename.endsWith('notifications.js')?{getAppSettings:()=>({notificationDurationSeconds:7})}:load.call(this,request,parent,main);};
 try {
  const notify=require('../src/notifications');
  for(let i=0;i<8;i++)notify.info('Message '+i);
  notify.info('10 %',{key:'download'});notify.info('20 %',{key:'download'});
  const win=new EventEmitter(),contents=new EventEmitter(),sent=[];
  win.webContents=contents;win.isDestroyed=()=>false;contents.isDestroyed=()=>false;contents.isLoadingMainFrame=()=>false;contents.send=(...args)=>sent.push(args);
  notify.registerWindow(win);contents.emit('did-finish-load');
  assert.equal(sent.length,3);assert.equal(sent.at(-1)[1].text,'20 %');
  assert.ok(sent.every(([channel,payload])=>channel==='ui:notification' && payload.durationSeconds===7));
  notify.Error('Failure');assert.equal(sent.at(-1)[1].type,'error');
  win.emit('closed');
 }finally{Module._load=load;delete require.cache[require.resolve('../src/notifications')];}
});
