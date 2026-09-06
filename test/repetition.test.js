const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const Module = require('module');
const ExcelJS = require('exceljs');

test('bestehende Repetition gewichtet weiterhin ausschließlich nach Repetitionsnoten', async () => {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'repetierer-regression-'));
	const originalLoad = Module._load;
	Module._load = function(request, parent, isMain) {
		if (request === 'electron') return { app: { getPath: () => dir } };
		return originalLoad.call(this, request, parent, isMain);
	};
	try {
		const workbookPath = path.join(dir, 'class.xlsx');
		const workbook = new ExcelJS.Workbook();
		const sheet = workbook.addWorksheet('3a');
		sheet.getCell('A1').value = 'repetierer';
		sheet.getCell('A6').value = 'Anna';
		sheet.getCell('A7').value = 'Max';
		sheet.getCell('B7').value = 5;
		await workbook.xlsx.writeFile(workbookPath);
		const program = require('../src/program.js');
		await new Promise(resolve => program.setFile(workbookPath, resolve));
		await new Promise(resolve => program.setClass('3a', resolve));
		const probabilities = program.getProbabilities();
		assert.equal(probabilities[0].weight, 729);
		assert.equal(probabilities[1].weight, 243);
		const originalRandom = Math.random;
		Math.random = () => 0;
		try { assert.equal(program.selectPerson()[0], 'Anna'); } finally { Math.random = originalRandom; }
	} finally {
		Module._load = originalLoad;
		fs.rmSync(dir, { recursive: true, force: true });
	}
});
