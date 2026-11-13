const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const {FileMatcher}=require('app-builder-lib/out/fileMatcher');
const root=path.resolve(__dirname,'..');
const files=require('../package.json').build.files;
const filter=new FileMatcher(root,path.join(root,'unused-test-destination'),value=>value,files).createFilter();
const included=file=>{const absolute=path.join(root,file);return filter(absolute,fs.statSync(absolute));};
test('EXE includes every local renderer module required by index.html',()=>{
 const pending=['public/index.html','app.js'],seen=new Set();
 while(pending.length){
  const file=pending.pop();if(seen.has(file))continue;seen.add(file);
  assert.ok(included(file),`${file} is missing from build.files`);
  const source=fs.readFileSync(path.join(root,file),'utf8');
  for(const entry of source.matchAll(/require\(['"](\.[^'"]+)['"]\)/g)){
   const resolved=path.posix.normalize(path.posix.join(path.posix.dirname(file),entry[1]));
   const dependency=path.posix.extname(resolved)?resolved:resolved+'.js';
   pending.push(dependency);
  }
 }
 assert.ok(seen.has('public/settingsScope.js'));
 assert.ok(seen.has('public/homeworkPage.js'));
 assert.ok(seen.has('src/sessionCloseGuard.js'));
 assert.equal(included('scripts/homework-ui-check.cjs'),false,'Test scripts stay outside the EXE');
 assert.equal(included('Excel Template.xlsx'),false,'Workbook data stays outside the EXE');
});
