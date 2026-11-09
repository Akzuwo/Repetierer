const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const os=require('node:os');
const Excel=require('exceljs');
const reader=require('../src/excelReader');
const dates=require('../src/jokerDates');
const load=file=>new Promise(resolve=>reader.init(file,resolve));
const migrate=token=>new Promise(resolve=>reader.migrate_joker_dates(token,resolve));
const call=(fn,...args)=>new Promise(resolve=>fn(...args,(persons,result)=>resolve({persons,...result})));
async function fixture(custom=false) {
 const dir=fs.mkdtempSync(path.join(os.tmpdir(),'joker-dates-')),file=path.join(dir,'Klassen.xlsx');
 const book=new Excel.Workbook();
 const sheet=book.addWorksheet('3a');sheet.getCell('A1').value='repetierer';sheet.getCell('Q3').value='Joker Date';sheet.mergeCells('Q3:Q4');
 sheet.getCell('A6').value='Anna';sheet.getCell('H6').value=3;sheet.getCell('B6').value=5.5;sheet.getCell('I6').value={formula:'AVERAGE(B6:G6)'};
 sheet.getCell('Q6').value='01.09.2026; 02.09.2026; 02.09.2026';sheet.getCell('A7').value='Max';sheet.getCell('H7').value=1;
 if(custom){sheet.getCell('R3').value='Eigene Notiz';sheet.getCell('R6').value='behalten';sheet.mergeCells('R3:R4');}
 const second=book.addWorksheet('4b');second.getCell('A1').value='repetierer';second.getCell('Q5').value='Joker Daten';second.getCell('A6').value='Lena';second.getCell('H6').value=1;
 const other=book.addWorksheet('Notizen');other.getCell('Q6').value='nicht verändern';
 await book.xlsx.writeFile(file);await load(file);return {file,dir,original:fs.readFileSync(file)};
}
async function read(file){const book=new Excel.Workbook();await book.xlsx.readFile(file);return book;}
test('Migration preserves grades, formulas, counters, custom columns and exact original backup; repeat is a no-op',async()=>{
 const f=await fixture(true);const status=reader.getJokerDateMigrationStatus();
 assert.equal(status.needsMigration,true);assert.equal(status.sheets.length,2);assert.equal(status.sheets[0].dates,3);assert.equal(status.sheets[0].keepsLegacyColumn,true);
 assert.deepEqual(fs.readFileSync(f.file),f.original,'Preview never writes the file');
 const result=await migrate(status.fingerprint);assert.equal(result.success,true);assert.equal(result.migratedSheets,2);assert.deepEqual(fs.readFileSync(result.backupPath),f.original);
 const book=await read(f.file),sheet=book.getWorksheet('3a'),layout=dates.layout(sheet);
 assert.equal(layout.columns.length,6);assert.equal(layout.columns[0].column,19);
 assert.deepEqual(layout.columns.slice(0,3).map(item=>sheet.getCell(6,item.column).value),['01.09.2026','02.09.2026','02.09.2026']);
 assert.equal(sheet.getCell('H6').value,3);assert.equal(sheet.getCell('B6').value,5.5);assert.equal(sheet.getCell('I6').value.formula,'AVERAGE(B6:G6)');assert.equal(sheet.getCell('R6').value,'behalten');assert.equal(book.getWorksheet('Notizen').getCell('Q6').value,'nicht verändern');
 assert.equal(reader.getJokerDateMigrationStatus().needsMigration,false);
 const migrated=fs.readFileSync(f.file);assert.equal((await migrate(reader.getJokerDateMigrationStatus().fingerprint)).alreadyMigrated,true);assert.deepEqual(fs.readFileSync(f.file),migrated);
});
test('Stale preview, locked files and invalid formula histories never change the source',async()=>{
 const f=await fixture();const token=reader.getJokerDateMigrationStatus().fingerprint;
 const book=await read(f.file);book.getWorksheet('3a').getCell('B6').value=6;await book.xlsx.writeFile(f.file);const changed=fs.readFileSync(f.file);
 assert.equal((await migrate(token)).reason,'file-changed');assert.deepEqual(fs.readFileSync(f.file),changed);
 await load(f.file);fs.writeFileSync(path.join(f.dir,'~$Klassen.xlsx'),'lock');
 assert.equal((await migrate(reader.getJokerDateMigrationStatus().fingerprint)).reason,'excel-locked');assert.deepEqual(fs.readFileSync(f.file),changed);
 const formulaSheet=new Excel.Workbook().addWorksheet('x');formulaSheet.getCell('A6').value='Anna';formulaSheet.getCell('Q6').value={formula:'TODAY()'};
 assert.throws(()=>dates.inspect(formulaSheet),/nicht unterstützten/);
});
test('Direct Joker writes, pending entries, undo/redo and roster reorder retain separate dates',async()=>{
 const f=await fixture();assert.equal((await migrate(reader.getJokerDateMigrationStatus().fingerprint)).success,true);
 const people=reader.read('3a');assert.equal((await call(reader.write_joker,'3a',people[0])).success,true);
 let book=await read(f.file);const today=book.getWorksheet('3a').getCell('T6').value;assert.match(today,/^\d{2}\.\d{2}\.\d{4}$/);assert.equal(book.getWorksheet('3a').getCell('H6').value,2);
 const entry={id:'test',type:'joker',className:'3a',personId:0,date:today};
 assert.equal((await call(reader.undo_entry,'3a',entry)).success,true);book=await read(f.file);assert.equal(book.getWorksheet('3a').getCell('T6').value,null);assert.equal(book.getWorksheet('3a').getCell('S6').value,'02.09.2026');
 assert.equal((await call(reader.apply_entries,'3a',[entry])).success,true);
 const edit=await call(reader.edit_persons,'3a',[{id:1,name:'Max',joker:1},{id:0,name:'Anna',joker:2}]);assert.equal(edit.success,true);
 book=await read(f.file);assert.equal(book.getWorksheet('3a').getCell('T7').value,today);assert.equal(book.getWorksheet('3a').getCell('Q6').value,null);assert.equal(book.getWorksheet('3a').getCell('I7').value.formula,'AVERAGE(B6:G6)');
});
test('History expands without overwriting custom fields; duplicate dates undo one occurrence only',()=>{
 const sheet=new Excel.Workbook().addWorksheet('3a');sheet.getCell('A6').value='Anna';dates.migrateWorksheet(sheet);
 sheet.getCell('W5').value='Eigenes Feld';sheet.getCell('W6').value='bleibt';
 for(let i=0;i<7;i++) dates.append(sheet,6,'27.10.2026');
 assert.equal(sheet.getCell('W6').value,'bleibt');assert.equal(dates.layout(sheet).columns.length,7);assert.equal(sheet.getCell('X6').value,'27.10.2026');
 assert.equal(dates.remove(sheet,6,'27.10.2026'),true);assert.equal(sheet.getCell('X6').value,null);assert.equal(sheet.getCell('V6').value,'27.10.2026');
});
test('New template has date headers on every class sheet and all original formulas intact',async()=>{
 const book=await read(path.join(__dirname,'../Excel Template.xlsx'));
 assert.equal(book.worksheets.length,3);
 for(const sheet of book.worksheets){assert.equal(sheet.getCell('A1').value,'repetierer');assert.equal(dates.layout(sheet).columns.length,6);assert.equal(sheet.getCell('V3').value,'Joker Datum 6');assert.equal(sheet.getCell('H6').value,1);assert.match(sheet.getCell('I6').value.formula,/COUNT/);assert.equal(dates.inspect(sheet).needsMigration,false);}
});
