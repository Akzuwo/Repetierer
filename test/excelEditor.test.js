const test = require('node:test'), assert = require('node:assert/strict');
const fs = require('node:fs'), os = require('node:os'), path = require('node:path'), Module = require('node:module');
const ExcelJS = require('exceljs');

test('edits an explicit class without a session or selected repetition class and preserves all grade cells', async () => {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-editor-')), originalLoad = Module._load;
	Module._load = function(request,parent,isMain) {return request === 'electron' ? {app:{getPath:()=>dir}} : originalLoad.call(this,request,parent,isMain);};
	try {
		const file = path.join(dir,'classes.xlsx'), book = new ExcelJS.Workbook();
		for (const cls of ['A','B']) {const sheet = book.addWorksheet(cls); sheet.getCell('A1').value='repetierer';sheet.getCell('A6').value=`Person ${cls}`;sheet.getCell('H6').value=2;}
		for (let col=2; col<=7; col++) book.getWorksheet('B').getCell(6,col).value=col;
		await book.xlsx.writeFile(file);
		const program = require('../src/program');
		await new Promise(resolve=>program.setFile(file,resolve));
		assert.equal(program.getCurrentClass(),undefined);
		const roster = program.getClassPersons('B');
		assert.equal(roster.length,1,'Includes students with six existing grades');
		const save = (people,cls) => new Promise(resolve=>program.saveEditorPersons(people,resolve,cls));
		assert.equal((await save([{id:0,name:'B bearbeitet',joker:4},{id:null,name:'Neu',joker:1}],'B')).success,true);
		assert.equal(program.getCurrentClass(),undefined,'Editing never selects a repetition class');
		await new Promise(resolve=>program.setClass('A',resolve));
		assert.equal((await save([{id:0,name:'B nochmals',joker:3}],'B')).success,true);
		assert.equal(program.getCurrentClass(),'A');
		assert.equal(program.getPersons()[0].name,'Person A');
		const before = fs.readFileSync(file);
		assert.equal((await save([{id:0,name:'invalid',joker:0}],'Missing')).success,false);
		assert.deepEqual(fs.readFileSync(file),before,'Unknown class cannot change the workbook');
		const saved = new ExcelJS.Workbook(); await saved.xlsx.readFile(file);
		assert.equal(saved.getWorksheet('A').getCell('A6').value,'Person A');
		assert.equal(saved.getWorksheet('B').getCell('A6').value,'B nochmals');
		assert.deepEqual([2,3,4,5,6,7].map(col=>saved.getWorksheet('B').getCell(6,col).value),[2,3,4,5,6,7]);
		assert.equal(saved.getWorksheet('B').getCell('H6').value,3);
	} finally {Module._load = originalLoad; fs.rmSync(dir,{recursive:true,force:true});}
});
