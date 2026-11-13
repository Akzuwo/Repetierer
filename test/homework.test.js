const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),os=require('node:os'),path=require('node:path'),Module=require('node:module');
const Excel=require('exceljs'),homework=require('../src/homework');
const profile=fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-homework-test-'));
const load=Module._load;Module._load=function(request,parent,isMain){return request==='electron'?{app:{getPath:()=>profile}}:load.call(this,request,parent,isMain);};
const program=require('../src/program'),reader=require('../src/excelReader'),storage=require('../src/storage');
Module._load=load;
const call=(fn,...args)=>new Promise(resolve=>fn(...args,(_people,result)=>resolve(result)));
async function fixture(){
 const dir=fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-homework-book-')),file=path.join(dir,'class.xlsx'),book=new Excel.Workbook();
 for(const name of ['3a','3b']){const sheet=book.addWorksheet(name);sheet.getCell('A1').value='repetierer';sheet.getCell('A6').value='Anna';sheet.getCell('A7').value='Max';sheet.getCell('B6').value=4;sheet.getCell('H6').value=1;sheet.getCell('H7').value=1;sheet.getCell('I6').value={formula:'AVERAGE(B6:G6)',result:4};sheet.getCell('AO5').value='Eigenes Feld';sheet.getCell('AO6').value='behalten';}
 book.addWorksheet('Notizen').getCell('A6').value='Original';await book.xlsx.writeFile(file);
 await new Promise(r=>program.setFile(file,r));await new Promise(r=>program.setClass('3a',r));return {dir,file,original:fs.readFileSync(file)};
}
const read=async file=>{const book=new Excel.Workbook();await book.xlsx.readFile(file);return book;};
const migrate=()=>new Promise(r=>program.migrateHomework(program.getHomeworkStatus().fingerprint,r));
const save=grade=>new Promise(r=>program.saveHomework(0,grade,r));
test('class settings stay only in the app and are isolated by workbook and class',async()=>{
 const f=await fixture();assert.equal(program.getCurrentSettings().homeworkEnabled,true);assert.equal(program.getCurrentSettings().homeworkJokerThreshold,2);
 storage.saveScopedSettings({homeworkEnabled:false,homeworkWeight:2,probabilityDecreaseFactor:7,notificationDurationSeconds:4},f.file,'3a');
 assert.equal(program.getCurrentSettings().homeworkEnabled,false);assert.equal(program.getCurrentSettings().probabilityDecreaseFactor,7);
 assert.equal(storage.getClassSettings(f.file,'3b').homeworkEnabled,true);assert.equal(storage.getClassSettings(f.file+'.other','3a').probabilityDecreaseFactor,3);
 assert.equal(storage.getAppSettings().notificationDurationSeconds,4);assert.equal(storage.getAppSettings().homeworkWeight,1.2);
 assert.equal(await save(5),null);assert.deepEqual(fs.readFileSync(f.file),f.original,'Rules never modify Excel');
 storage.saveScopedSettings({homeworkEnabled:true,homeworkWeight:1.2,homeworkJokerThreshold:2},f.file,'3a');
});
test('first activation extends all classes once, exact backup, stale preview and locks are safe',async()=>{
 const f=await fixture(),plan=program.getHomeworkStatus();assert.equal(plan.needsMigration,true);assert.deepEqual(plan.sheets.map(s=>s.name),['3a','3b']);assert.deepEqual(fs.readFileSync(f.file),f.original);
 assert.equal((await new Promise(r=>program.migrateHomework('stale',r))).reason,'file-changed');
 const lock=path.join(f.dir,'~$class.xlsx');fs.writeFileSync(lock,'lock');assert.equal((await migrate()).reason,'excel-locked');assert.deepEqual(fs.readFileSync(f.file),f.original);fs.unlinkSync(lock);
 const result=await migrate();assert.equal(result.success,true);assert.deepEqual(fs.readFileSync(result.backupPath),f.original);
 const book=await read(f.file),sheet=book.getWorksheet('3a');assert.equal(homework.layout(sheet).slots.length,6);assert.equal(homework.layout(sheet).average.column,42);assert.equal(sheet.getCell('AO6').value,'behalten');assert.equal(sheet.getCell('I6').formula,'AVERAGE(B6:G6)');assert.equal(homework.layout(book.getWorksheet('3b')).slots.length,6);
 assert.equal(result.migratedSheets,2);assert.equal(book.getWorksheet('Notizen').getCell('A6').value,'Original');
 await new Promise(r=>program.setClass('3b',r));assert.equal(program.getHomeworkStatus().needsMigration,false,'Next class needs no second dialog');
 const before=fs.readFileSync(f.file);assert.equal((await migrate()).alreadyMigrated,true);assert.deepEqual(fs.readFileSync(f.file),before);
});
test('each second presentation earns one joker, immutable weights, stable undo/redo IDs and expanding history',async()=>{
 const f=await fixture();await migrate();const first=await save(6);assert.equal(first.excelWriteSucceeded,true);assert.equal(first.awardExtraJoker,false);
 assert.ok(Math.abs(program.getStudentGradeHistory('3a',0).weightedAverage-5.090909)<.00001);
 storage.saveScopedSettings({homeworkWeight:2,homeworkJokerThreshold:2},f.file,'3a');const second=await save(5);assert.equal(second.awardExtraJoker,true);assert.equal(second.weight,2);
 let rows=program.getHomeworkPersons();assert.equal(rows[0].joker,2);assert.deepEqual(rows[0].homework.map(e=>e.weight),[1.2,2]);
 const backup=storage.addBackupEntry(second);assert.notEqual(backup.id,second.id);assert.equal((await new Promise(r=>program.undoLastExcelEntry(backup,r))).success,true);assert.equal(program.getHomeworkPersons()[0].joker,1);
 assert.equal((await new Promise(r=>program.redoExcelEntry(backup,r))).success,true);assert.equal(program.getHomeworkPersons()[0].joker,2);
 assert.equal((await new Promise(r=>program.redoExcelEntry({...backup,id:'another-backup-id'},r))).success,true);assert.equal(program.getHomeworkPersons()[0].homework.length,2);assert.equal(program.getHomeworkPersons()[0].joker,2);
 for(let i=0;i<5;i++)await save(4);rows=program.getHomeworkPersons();assert.equal(rows[0].homework.length,7);assert.equal(rows[0].joker,4);assert.equal(program.getClassPersons('3a')[0].grades,1,'Homework does not consume repetition slots');
 const book=await read(f.file);assert.equal(homework.layout(book.getWorksheet('3a')).slots.length,7);assert.equal(book.getWorksheet('3a').getCell('AO6').value,'behalten');
});
test('locked Excel queues distinct presentations with the correct joker threshold, retry is idempotent',async()=>{
 const f=await fixture();await migrate();const before=fs.readFileSync(f.file),lock=path.join(f.dir,'~$class.xlsx');fs.writeFileSync(lock,'lock');
 const first=await save(5);assert.equal(first.excelWriteSucceeded,false);storage.addPendingExcelEntry(storage.addBackupEntry(first));
 const second=await save(6);assert.equal(second.awardExtraJoker,true);storage.addPendingExcelEntry(storage.addBackupEntry(second));assert.deepEqual(fs.readFileSync(f.file),before);fs.unlinkSync(lock);
 const pending=storage.getPendingExcelEntries().filter(e=>e.filePath===f.file);assert.equal((await new Promise(r=>program.applyPendingExcelEntries(pending,r))).applied.length,2);
 assert.equal(program.getHomeworkPersons()[0].homework.length,2);assert.equal(program.getHomeworkPersons()[0].joker,2);
 await new Promise(r=>program.applyPendingExcelEntries(pending,r));assert.equal(program.getHomeworkPersons()[0].homework.length,2);assert.equal(program.getHomeworkPersons()[0].joker,2);
 storage.removePendingExcelEntries(pending.map(e=>e.id));
});
test('roster reorder moves grades, dates, weights and IDs together; invalid grades cannot write',async()=>{
 const f=await fixture();await migrate();await save(5);const before=fs.readFileSync(f.file);assert.equal(await save(9),null);assert.deepEqual(fs.readFileSync(f.file),before);
 assert.equal((await new Promise(r=>program.saveEditorPersons([{id:1,name:'Max',joker:1},{id:0,name:'Anna',joker:1}],r,'3a'))).success,true);
 const people=program.getHomeworkPersons();assert.equal(people[1].name,'Anna');assert.equal(people[1].homework[0].grade,5);assert.equal(people[0].homework.length,0);
 const book=await read(f.file),sheet=book.getWorksheet('3a'),column=homework.layout(sheet).average.column;assert.match(sheet.getCell(7,column).formula,/B7:G7/);
});

test('shipped template has homework data slots on every class; formulas compute the weighted result',async()=>{
 const book=await read(path.join(__dirname,'../Excel Template.xlsx'));
 for(const sheet of book.worksheets){
  assert.equal(homework.layout(sheet).slots.length,6);const slot=homework.layout(sheet).slots[0];
  assert.equal(sheet.getColumn(slot.weight).hidden,true);assert.equal(sheet.getColumn(slot.meta).hidden,true);
  assert.equal(sheet.getCell(6,slot.weight).value,null,'Blank slots contain no class settings');
  sheet.getCell('B6').value=4;homework.append(sheet,6,{id:'template-test',grade:6,weight:1.2,date:'09.11.2026'});
  assert.ok(Math.abs(homework.weightedAverage(sheet,6)-5.090909)<.00001);
  const cell=sheet.getCell(6,homework.layout(sheet).average.column);assert.ok(Math.abs(cell.value.result-5.090909)<.00001);assert.match(cell.formula,/ISNUMBER/);
 }
});

test('pending homework and undo never cross into another workbook with the same class name',async()=>{
 const a=await fixture();await migrate();const entry=await save(5);const b=await fixture();const before=fs.readFileSync(b.file);
 const result=await new Promise(r=>program.applyPendingExcelEntries([entry],r));assert.deepEqual(result.applied,[]);
 assert.equal((await new Promise(r=>program.undoLastExcelEntry(entry,r))).reason,'other-workbook');
 assert.equal((await new Promise(r=>program.redoExcelEntry(entry,r))).reason,'other-workbook');
 assert.deepEqual(fs.readFileSync(b.file),before);assert.notEqual(a.file,b.file);
});
test('legacy class rules become defaults; global saves and another class preserve overrides',async()=>{
 const f=await fixture();storage.saveAppSettings({extraJokerAfterThreeGrades:true,probabilityDecreaseFactor:4,boostNeverSelected:true,neverSelectedBoostFactor:2,wiggersRuleEnabled:true,wiggersRuleDurationMinutes:45,homeworkEnabled:true,homeworkJokerThreshold:2,homeworkWeight:1.2});
 assert.equal(storage.getClassSettings(f.file,'3b').probabilityDecreaseFactor,4);
 storage.saveScopedSettings({probabilityDecreaseFactor:8,wiggersRuleDurationMinutes:20,homeworkJokerThreshold:3,homeworkWeight:1.5},f.file,'3a');
 assert.equal(program.getCurrentSettings().probabilityDecreaseFactor,8);assert.equal(program.getCurrentSettings().homeworkJokerThreshold,3);
 assert.equal(storage.getClassSettings(f.file,'3b').probabilityDecreaseFactor,4);assert.equal(storage.getClassSettings(f.file,'3b').homeworkJokerThreshold,2);
 const bytes=fs.readFileSync(f.file);storage.saveScopedSettings({notificationDurationSeconds:7},null,null);
 assert.equal(program.getCurrentSettings().probabilityDecreaseFactor,8);assert.equal(program.getCurrentSettings().notificationDurationSeconds,7);assert.deepEqual(fs.readFileSync(f.file),bytes);
});


test('workbook migration completes remaining classes and preserves existing homework entries',async()=>{
 const f=await fixture(),book=await read(f.file),first=book.getWorksheet('3a');
 homework.migrateWorksheet(first);homework.append(first,6,{id:'existing-homework',grade:5,weight:1.2,date:'13.11.2026'});
 await book.xlsx.writeFile(f.file);await new Promise(r=>program.setFile(f.file,r));await new Promise(r=>program.setClass('3b',r));
 const before=fs.readFileSync(f.file),result=await migrate();assert.equal(result.success,true);assert.equal(result.migratedSheets,1);
 assert.deepEqual(fs.readFileSync(result.backupPath),before);
 const updated=await read(f.file);assert.equal(homework.history(updated.getWorksheet('3a'),6)[0].id,'existing-homework');
 assert.equal(homework.history(updated.getWorksheet('3a'),6)[0].grade,5);
 assert.equal(homework.layout(updated.getWorksheet('3b')).slots.length,6);
 assert.equal(updated.getWorksheet('3b').getCell('AO6').value,'behalten');
 const stable=fs.readFileSync(f.file);assert.equal((await migrate()).alreadyMigrated,true);assert.deepEqual(fs.readFileSync(f.file),stable);
});
