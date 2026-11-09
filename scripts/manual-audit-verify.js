const {app,BrowserWindow}=require('electron');
if(!app){const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;const r=require('child_process').spawnSync(require('electron'),[__filename],{env,stdio:'inherit',windowsHide:true});process.exit(r.status??1);}
const fs=require('fs'),path=require('path'),os=require('os'),assert=require('assert/strict');
const output=path.resolve(__dirname,'..','..','manual-audit');
app.setPath('userData',fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-gallery-test-')));
async function main(){
 await app.whenReady();
 const manifest=JSON.parse(fs.readFileSync(path.join(output,'manifest.json')));
 const files=manifest.entries.map(e=>e.file);
 assert.equal(new Set(files).size,files.length);
 for(const f of files){const png=fs.readFileSync(path.join(output,f));assert.equal(png.readUInt32BE(0),0x89504e47);assert.ok(png.length>3000);assert.ok(!require('electron').nativeImage.createFromBuffer(png).isEmpty());}
 const w=new BrowserWindow({width:1200,height:900,show:false,webPreferences:{offscreen:true}}),errors=[];
 w.webContents.on('console-message',e=>{if(e.level==='error'||e.level===3)errors.push(e.message);});
 const evaluate=s=>w.webContents.executeJavaScript(s,true);
 await w.loadFile(path.join(output,'index.html'));
 assert.equal(await evaluate('document.querySelectorAll("article").length'),files.length);
 await evaluate(`document.images[0].decode()`);
 assert.equal(await evaluate('document.images[0].naturalWidth>0'),true);
 const comment='Bitte Abstände ändern.\n\n- Button klarer beschriften\n## Eigene Überschrift\nUmlaute: äöü · <script>Text</script>';
 await evaluate(`(()=>{const t=document.querySelector('textarea');t.value=${JSON.stringify(comment)};t.dispatchEvent(new Event('input'));})()`);
 await w.reload();await new Promise(r=>setTimeout(r,250));
 assert.equal(await evaluate('document.querySelector("textarea").value'),comment);
 await evaluate(`document.getElementById('filter').click()`);
 assert.equal(await evaluate('[...document.querySelectorAll("article")].filter(e=>!e.hidden).length'),1);
 await evaluate(`document.getElementById('filter').click();document.getElementById('search').value='nothing-matches-this-query';document.getElementById('search').dispatchEvent(new Event('input'))`);
 assert.equal(await evaluate('[...document.querySelectorAll("article")].filter(e=>!e.hidden).length'),0);
 const download=path.join(app.getPath('userData'),'comments-export.md');
 const downloaded=new Promise((resolve,reject)=>w.webContents.session.once('will-download',(_,item)=>{item.setSavePath(download);item.once('done',(_,state)=>state==='completed'?resolve():reject(new Error(state)));}));
 await evaluate(`document.getElementById('export').click()`);await downloaded;
 const md=fs.readFileSync(download,'utf8');assert.ok(md.includes(comment));assert.ok(md.includes(files[0]));
 await evaluate(`(()=>{const t=document.querySelector('textarea');t.value='';t.dispatchEvent(new Event('input'));})()`);
 await evaluate(`(()=>{const f=new File([${JSON.stringify(md)}],'comments.md',{type:'text/markdown'}),d=new DataTransfer();d.items.add(f);const input=document.getElementById('import');input.files=d.files;input.dispatchEvent(new Event('change'));})()`);
 await new Promise(r=>setTimeout(r,250));assert.equal(await evaluate('document.querySelector("textarea").value'),comment);
 await evaluate(`localStorage.clear()`);await w.reload();await new Promise(r=>setTimeout(r,500));await evaluate(`document.images[0].decode()`);
 fs.writeFileSync(path.join(output,'gallery-verification.png'),(await w.webContents.capturePage()).toPNG());
 assert.deepEqual(errors,[]);
 const result={passed:true,screenshots:files.length,modals:manifest.modalCoverage.length,checks:['PNG files and unique names','all gallery images decode','comment persistence after reload','comment filter','search','actual Markdown download','Markdown import roundtrip including multiline and HTML-like text','no gallery console errors']};
 fs.writeFileSync(path.join(output,'verification.json'),JSON.stringify(result,null,2));console.log(JSON.stringify(result,null,2));w.destroy();
}
main().then(()=>app.exit(0)).catch(e=>{console.error(e);app.exit(1)});
