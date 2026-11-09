// Convert the approved transparent PNG to Windows ICO sizes. No generation or
// background editing takes place here; alpha is preserved by nativeImage.
const {app,nativeImage}=require('electron');
if(!app){const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;const r=require('node:child_process').spawnSync(require('electron'),[__filename],{env,stdio:'inherit',windowsHide:true});process.exit(r.status??1);}
const fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict');
app.whenReady().then(()=>{
  const dir=path.join(__dirname,'../public/icons'),png=nativeImage.createFromPath(path.join(dir,'icon.png'));
  assert.equal(png.isEmpty(),false);assert.deepEqual(png.getSize(),{width:1254,height:1254});
  const sizes=[16,20,24,32,40,48,64,96,128,256];
  const frames=sizes.map(size=>{
    const image=png.resize({width:size,height:size,quality:'best'}),bitmap=image.toBitmap();
    // Normalize sub-2% resampling noise to a transparent Windows mask.
    for(let p=0;p<bitmap.length;p+=4)if(bitmap[p+3]<=3)bitmap.fill(0,p,p+4);
    return nativeImage.createFromBitmap(bitmap,{width:size,height:size}).toPNG();
  });
  const header=Buffer.alloc(6+16*sizes.length);header.writeUInt16LE(1,2);header.writeUInt16LE(sizes.length,4);
  let offset=header.length;
  for(let i=0;i<sizes.length;i++){
    const entry=6+i*16;header[entry]=header[entry+1]=sizes[i]===256 ? 0 : sizes[i];
    header.writeUInt16LE(1,entry+4);header.writeUInt16LE(32,entry+6);header.writeUInt32LE(frames[i].length,entry+8);header.writeUInt32LE(offset,entry+12);offset+=frames[i].length;
    assert.equal(nativeImage.createFromBuffer(frames[i]).toBitmap()[3],0,'Transparent corner at '+sizes[i]);
  }
  fs.writeFileSync(path.join(dir,'icon.ico'),Buffer.concat([header,...frames]));
  console.log('APP_ICON_EXPORT_PASS: ten PNG-backed 32-bit ICO frames with transparent corners');
  app.exit(0);
}).catch(e=>{console.error(e);app.exit(1);});
