const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const {createRepetitionTimer} = require('../src/repetitionTimer');
const {withRepetitionDurations} = require('../src/classStats');

test('duration starts when the selected name is ready and freezes before Excel I/O', () => {
	let now = 100;
	const timer = createRepetitionTimer(() => now);
	assert.equal(timer.ready(null), false);
	const token = timer.select();
	now = 3500;
	assert.equal(timer.capture(), undefined, 'Roulette time is excluded');
	assert.equal(timer.ready('stale'), false);
	assert.equal(timer.ready(token), true);
	now = 5500;
	assert.equal(timer.ready(token), true, 'Duplicate readiness does not restart the clock');
	now = 6543;
	assert.equal(timer.capture(), 3043);
	now = 21000;
	assert.equal(timer.capture(), 3043, 'Writing/retrying cannot inflate the duration');
});

test('cancel, reselect and invalid clocks never reuse an earlier duration', () => {
	let now = 100;
	const timer = createRepetitionTimer(() => now);
	const first = timer.select(); timer.ready(first);
	timer.cancel();
	assert.equal(timer.token, null);
	assert.equal(timer.capture(), undefined);
	assert.equal(timer.ready(first), false);
	const next = timer.select();
	assert.notEqual(next, first);
	assert.equal(timer.ready(first), false);
	timer.ready(next); now = 99;
	assert.equal(timer.capture(), undefined);
	timer.select(); now = NaN; timer.ready(timer.token);
	assert.equal(timer.capture(), undefined);
});

test('averages only accepted timed grades for the matching workbook and class', () => {
	const filePath = path.resolve('duration-fixture.xlsx');
	const stats = [{className:'3a'}, {className:'3b'}, {className:'Old data'}];
	const grade = (durationMs, extra = {}) => ({type:'grade', filePath, className:'3a', excelWriteSucceeded:true, durationMs, ...extra});
	const entries = [grade(60000), grade(180000), grade(90000,{className:'3b'}),
		grade(undefined), grade(null), grade('999'), grade(-1), grade(Infinity),
		grade(999999,{type:'joker'}), grade(999999,{filePath:path.resolve('other.xlsx')}),
		grade(999999,{excelWriteSucceeded:false,excelWriteReason:'no-slot'}),
		grade(120000,{excelWriteSucceeded:false,excelWriteReason:'excel-locked'}),
		grade(120000,{excelWriteSucceeded:false,excelWriteReason:'excel-write-failed'})];
	const result = withRepetitionDurations(stats, entries, filePath);
	assert.equal(result[0].averageRepetitionDurationMs, 120000);
	assert.equal(result[0].timedRepetitionsCount, 4);
	assert.equal(result[1].averageRepetitionDurationMs, 90000);
	assert.equal(result[2].averageRepetitionDurationMs, null);
	assert.equal(result[2].timedRepetitionsCount, 0);
	assert.deepEqual(stats, [{className:'3a'}, {className:'3b'}, {className:'Old data'}]);
	assert.equal(withRepetitionDurations(stats, entries, null)[0].timedRepetitionsCount, 0);
});

test('timed grades persist with the real Excel and backup paths, including undo and redo', async () => {
	const fs = require('node:fs'), os = require('node:os'), Module = require('node:module');
	const ExcelJS = require('exceljs');
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'repetierer-duration-'));
	const originalLoad = Module._load;
	Module._load = function(request, parent, isMain) {
		if (request === 'electron') return {app:{getPath:()=>dir}};
		return originalLoad.call(this, request, parent, isMain);
	};
	try {
		const file = path.join(dir, 'class.xlsx');
		const workbook = new ExcelJS.Workbook(), sheet = workbook.addWorksheet('3a');
		sheet.getCell('A1').value = 'repetierer'; sheet.getCell('A6').value = 'Anna Beispiel';
		await workbook.xlsx.writeFile(file);
		const program = require('../src/program'), storage = require('../src/storage');
		await new Promise(resolve => program.setFile(file, resolve));
		await new Promise(resolve => program.setClass('3a', resolve));
		assert.equal(program.getClassStatistics()[0].averageRepetitionDurationMs, null);
		program.selectSpecificPerson(0);
		const token = program.getSelectedPerson().selectionToken;
		assert.equal(program.startRepetitionTiming(token), true);
		const entry = await new Promise(resolve => program.saveGrade(5, (_error, saved) => resolve(saved)));
		assert.equal(entry.excelWriteSucceeded, true);
		assert.equal(typeof entry.durationMs, 'number');
		const saved = storage.addBackupEntry(entry);
		assert.equal(storage.getBackup().entries[0].durationMs, entry.durationMs);
		await new Promise(resolve => program.setFile(file, resolve));
		assert.equal(program.getClassStatistics()[0].averageRepetitionDurationMs, entry.durationMs, 'Survives workbook reload');
		assert.equal(program.startRepetitionTiming(token), false);
		assert.equal((await new Promise(resolve => program.undoLastExcelEntry(saved, resolve))).success, true);
		storage.removeBackupEntries([saved.id]);
		assert.equal(program.getClassStatistics()[0].averageRepetitionDurationMs, null);
		assert.equal((await new Promise(resolve => program.redoExcelEntry(saved, resolve))).success, true);
		storage.addBackupEntry(saved);
		assert.equal(program.getClassStatistics()[0].averageRepetitionDurationMs, entry.durationMs);
		assert.equal(program.getClassStatistics()[0].timedRepetitionsCount, 1);
		program.selectSpecificPerson(0); program.cancelRepetitionTiming();
		const cancelled = await new Promise(resolve => program.saveGrade(4, (_error, saved) => resolve(saved)));
		assert.equal(cancelled.durationMs, undefined);
	} finally {Module._load = originalLoad; fs.rmSync(dir, {recursive:true, force:true});}
});
