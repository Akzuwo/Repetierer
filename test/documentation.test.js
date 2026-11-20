const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs/promises'),path=require('node:path'),os=require('node:os');
const {apiMarkdown,registerDocumentation,API_DOCUMENT_FILES}=require('../src/documentation');
const root=path.resolve(__dirname,'..');
test('offline Markdown combines all public API references, without operator instructions or credentials',async()=>{
 const content=await apiMarkdown(root,'7.1.0');
 assert.match(content,/^# Repetierer – API-Dokumentation/);assert.match(content,/App-Version: 7\.1\.0/);
 for(const endpoint of ['/v1/snapshot','/teacher/rounds/:id/reopen','/teacher/mail/send'])assert.ok(content.includes(endpoint));
 assert.match(content,/```powershell/);assert.match(content,/```python/);assert.match(content,/```json/);
 assert.doesNotMatch(content,/\.cloudflare-private|Entwicklung und Prüfung|Betrieb und Tests|\brt_[A-Za-z0-9_-]{43}\b|\b(?:xkeysib|xsmtpsib)-[A-Za-z0-9_-]{20,}/);
});
function fixture(dialog){let handler;registerDocumentation({app:{getAppPath:()=>root,getVersion:()=> 'test'},ipcMain:{handle:(channel,fn)=>{assert.equal(channel,'documentation:export');handler=fn;}},dialog:{showSaveDialog:(_parent,options)=>dialog(options)},BrowserWindow:{fromWebContents:()=>null}});return args=>handler({sender:{}},args);}
test('export saves actual Markdown, cancellation creates no file and unavailable IDs never open the picker',async()=>{
 const dir=await fs.mkdtemp(path.join(os.tmpdir(),'repetierer-documentation-'));let calls=0,mode='save';
 const invoke=fixture(async options=>{calls++;assert.equal(options.filters[0].extensions[0],'md');return mode==='cancel'?{canceled:true}:{filePath:path.join(dir,'reference')};});
 assert.equal((await invoke({id:'manual'})).ok,false);assert.equal((await invoke({id:'../../settings.json'})).ok,false);assert.equal(calls,0);
 const result=await invoke({id:'api'});assert.equal(result.ok,true);assert.equal(result.filePath,path.join(dir,'reference.md'));assert.match(await fs.readFile(result.filePath,'utf8'),/\/v1\/health/);
 const before=await fs.readdir(dir);mode='cancel';assert.equal((await invoke({id:'api'})).cancelled,true);assert.deepEqual(await fs.readdir(dir),before);
});
test('failed writes report failure; concurrent clicks open a single save dialog',async()=>{
 const dir=await fs.mkdtemp(path.join(os.tmpdir(),'repetierer-doc-fail-'));
 const failing=fixture(async()=>({filePath:path.join(dir,'missing','api.md')}));assert.equal((await failing({id:'api'})).ok,false);
 let release,opened;const ready=new Promise(resolve=>opened=resolve);
 const invoke=fixture(()=>{opened();return new Promise(resolve=>release=resolve);});
 const first=invoke({id:'api'});await ready;assert.equal((await invoke({id:'api'})).ok,false);release({canceled:true});assert.equal((await first).cancelled,true);
});
test('all documentation sources are included by the actual Electron packaging filter',async()=>{
 const {FileMatcher}=require('app-builder-lib/out/fileMatcher');
 const filter=new FileMatcher(root,path.join(root,'unused-test-destination'),value=>value,require('../package.json').build.files).createFilter();
 for(const file of [...API_DOCUMENT_FILES,'src/documentation.js'])assert.ok(filter(path.join(root,file),await fs.stat(path.join(root,file))),file+' must ship in the application');
 assert.equal(filter(path.join(root,'docs/cloudflare-deployment.md'),await fs.stat(path.join(root,API_DOCUMENT_FILES[0]))),false,'Private deployment instructions do not ship, even in checkouts without private files');
});
