const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const ExcelJS = require('exceljs');
const { createAiToolRegistry } = require('../src/ai/toolRegistry.js');
const { createAiAssistantService } = require('../src/ai/assistantService.js');
const { createOllamaService, MODEL } = require('../src/ai/ollamaService.js');
const excelReader = require('../src/excelReader.js');

function fixture() {
	const people = {
		'6a': [{ id: 0, name: 'Anna Müller', grades: 1, joker: 1 }, { id: 1, name: 'Ben Müller', grades: 0, joker: 1 }, { id: 2, name: 'Lia Meier', grades: 0, joker: 0 }],
		'6b': [{ id: 0, name: 'Clara Keller', grades: 2, joker: 0 }]
	};
	const active = { id: 'session-1', className: '6a', startedAt: '2026-09-16', presentStudents: people['6a'], absentStudents: [], participationEntries: [], repetitionEntries: [] };
	let gradeCommits = 0;
	let navigation = null;
	const program = {
		getClasses: () => Object.keys(people), getClassPersons: classId => people[classId] || [], getCurrentClass: () => '6a',
		getStudentGradeHistory: (classId, studentId) => ({ classId, studentId, grades: studentId === 0 ? [{ grade: 5.5, date: '15.09.2026', position: 1 }] : [] })
	};
	const sessions = {
		getActiveSession: () => active, readData: () => ({ sessions: [active] }),
		addParticipation: (id, studentId, points) => {
			active.participationEntries.push({ studentId, points });
			return { success: true, session: active };
		}
	};
	const makeRegistry = () => createAiToolRegistry({ program, sessions, commitGrade: async (classId, studentId, grade) => { gradeCommits++; return { saved: true, classId, studentId, grade }; }, notifyNavigation: command => { navigation = command; }, getUiContext: () => ({ page: 'repetition', classId: '6a', studentId: null, sessionId: active.id }) });
	return { makeRegistry, get gradeCommits() { return gradeCommits; }, get navigation() { return navigation; }, active };
}

test('Tool Registry exposes only approved tools and validates arguments', async () => {
	const registry = fixture().makeRegistry();
	assert.equal(registry.definitions.some(item => item.function.name === 'runSql'), false);
	await assert.rejects(registry.execute('runSql', {}), /Unbekanntes Tool/);
	await assert.rejects(registry.execute('getStudent', { classId: '6a', studentId: '0' }), /studentId/);
	await assert.rejects(registry.execute('getStudent', { classId: '6a', studentId: 0, sql: 'DROP' }), /Unbekanntes Argument/);
	await assert.rejects(registry.execute('getStudent', { classId: 'fake', studentId: 0 }), /Unbekannte Klasse/);
});

test('Student search resolves unique names and rejects ambiguous writes', async () => {
	const registry = fixture().makeRegistry();
	await assert.rejects(registry.execute('getStudentGrades', { classId: '6a', studentId: 0 }), /searchStudents/);
	const ambiguous = await registry.execute('searchStudents', { query: 'Müller', classId: '6a' });
	assert.equal(ambiguous.ambiguous, true);
	assert.equal(ambiguous.matches.length, 2);
	await assert.rejects(registry.execute('createRepetitionGrade', { classId: '6a', studentId: 0, grade: 5.5 }), /eindeutig/);
	const unique = await registry.execute('searchStudents', { query: 'Anna Müller', classId: '6a' });
	assert.equal(unique.matches.length, 1);
	const pending = await registry.execute('createRepetitionGrade', { classId: '6a', studentId: 0, grade: 5.5 });
	assert.equal(pending.confirmation.lines[2], 'Note 5.5');
	assert.deepEqual((await registry.execute('searchStudents', { query: 'Niemand' })).matches, []);
});

test('Read tools return scoped structured data; navigation uses approved command', async () => {
	const setup = fixture(); const registry = setup.makeRegistry();
	await registry.execute('searchStudents', { query: 'Anna Müller' });
	assert.equal((await registry.execute('getStudentGrades', { classId: '6a', studentId: 0 })).grades[0].grade, 5.5);
	assert.equal((await registry.execute('getClassStudents', { classId: '6a' })).students[1].gradeCount, 0);
	assert.deepEqual((await registry.execute('getStudentParticipation', { classId: '6a', studentId: 0 })).entries, []);
	await registry.execute('openStudent', { classId: '6a', studentId: 0 });
	assert.deepEqual(setup.navigation, { page: 'student', classId: '6a', studentId: 0, name: 'Anna Müller', gradeCount: 1, remainingJokers: 1 });
});

test('Invalid grades, IDs and points cannot reach a write method', async () => {
	const setup = fixture(); const registry = setup.makeRegistry();
	await registry.execute('searchStudents', { query: 'Anna Müller' });
	for (const grade of [0, 6.1, '5.5', Infinity]) await assert.rejects(registry.execute('createRepetitionGrade', { classId: '6a', studentId: 0, grade }), /Note|grade/);
	await assert.rejects(registry.execute('addParticipationPoints', { classId: '6a', studentId: 0, points: 7 }), /Mitarbeitspunkte/);
	assert.equal(setup.gradeCommits, 0);
});

test('Participation write uses the existing session store only after a valid confirmation', async () => {
	const setup = fixture(); const registry = setup.makeRegistry();
	await registry.execute('searchStudents', { query: 'Anna Müller' });
	const prepared = await registry.execute('addParticipationPoints', { classId: '6a', studentId: 0, points: 2 });
	assert.equal(setup.active.participationEntries.length, 0);
	const committed = await registry.commit(prepared.confirmation);
	assert.equal(committed.saved, true);
	assert.deepEqual(setup.active.participationEntries, [{ studentId: 0, points: 2 }]);
});

test('A write is committed only after confirmation and cancellation leaves data intact', async () => {
	const setup = fixture();
	let turn = 0;
	const ollama = { chat: async () => ++turn === 1 ? { content: '', tool_calls: [{ function: { name: 'searchStudents', arguments: { query: 'Anna Müller' } } }] } : turn === 2 ? { content: '', tool_calls: [{ function: { name: 'createRepetitionGrade', arguments: { classId: '6a', studentId: 0, grade: 5.5 } } }] } : { content: 'Die Note wurde eingetragen.' } };
	const assistant = createAiAssistantService({ ollama, makeRegistry: setup.makeRegistry });
	const pending = await assistant.ask('Gib Anna Müller eine 5.5.');
	assert.equal(pending.type, 'confirmation'); assert.equal(setup.gradeCommits, 0);
	await assert.rejects(assistant.decide('wrong', true), /ungültig/);
	const result = await assistant.decide(pending.token, true);
	assert.equal(result.text, 'Die Note wurde eingetragen.'); assert.equal(setup.gradeCommits, 1);
	turn = 0;
	const another = await assistant.ask('Nochmals');
	assert.equal(another.type, 'confirmation');
	assert.match((await assistant.decide(another.token, false)).text, /Abgebrochen/);
	assert.equal(setup.gradeCommits, 1);
});

test('A failed confirmed write is reported as failed without another model claim', async () => {
	const setup = fixture(); let calls = 0;
	const assistant = createAiAssistantService({
		ollama: { chat: async () => { calls++; return calls === 1 ? { tool_calls: [{ function: { name: 'searchStudents', arguments: { query: 'Anna Müller' } } }] } : { tool_calls: [{ function: { name: 'createRepetitionGrade', arguments: { classId: '6a', studentId: 0, grade: 5.5 } } }] }; } },
		makeRegistry: () => { const registry = setup.makeRegistry(); registry.commit = async () => { throw new Error('Excel gesperrt'); }; return registry; }
	});
	const pending = await assistant.ask('Gib Anna eine 5.5');
	const response = await assistant.decide(pending.token, true);
	assert.match(response.text, /Änderung fehlgeschlagen: Excel gesperrt/);
	assert.equal(calls, 2);
});

test('Agent loop sends successive tool results back to the local model', async () => {
	const setup = fixture(); const seen = [];
	const calls = [
		{ content: '', tool_calls: [{ function: { name: 'searchStudents', arguments: { query: 'Anna Müller' } } }] },
		{ content: '', tool_calls: [{ function: { name: 'getStudentGrades', arguments: { classId: '6a', studentId: 0 } } }] },
		{ content: 'Anna hatte zuletzt 5.5.' }
	];
	const assistant = createAiAssistantService({ ollama: { chat: async messages => { seen.push(structuredClone(messages)); return calls.shift(); } }, makeRegistry: setup.makeRegistry });
	const result = await assistant.ask('Welche Note hatte Anna zuletzt?');
	assert.equal(result.text, 'Anna hatte zuletzt 5.5.');
	assert.equal(seen[2].filter(message => message.role === 'tool').length, 2);
	assert.match(seen[0][1].content, /App-Kontext/);
});

test('Ollama reports missing service and missing model without a cloud fallback', async () => {
	const offline = createOllamaService({ platform: 'win32', localAppData: fs.mkdtempSync(path.join(os.tmpdir(), 'repetierer-no-ollama-')), fetchImpl: async () => { throw new Error('offline'); } });
	assert.equal((await offline.status()).state, 'not-installed');
	await assert.rejects(offline.chat([{ role: 'user', content: 'Hallo' }]), /nicht erreichbar/);
	const missing = createOllamaService({ fetchImpl: async () => ({ ok: true, json: async () => ({ models: [] }) }) });
	assert.equal((await missing.status()).state, 'model-missing');
	await assert.rejects(missing.chat([{ role: 'user', content: 'Hallo' }]), new RegExp(MODEL));
});

test('Ordinary question needs no tool', async () => {
	const assistant = createAiAssistantService({ ollama: { chat: async () => ({ content: 'Hallo!' }) }, makeRegistry: fixture().makeRegistry });
	assert.deepEqual(await assistant.ask('Hallo'), { type: 'answer', text: 'Hallo!' });
});

test('Chat UI and setup controls are included in the packaged app', () => {
	const root = path.join(__dirname, '..');
	const html = fs.readFileSync(path.join(root, 'public', 'index.html'), 'utf8');
	const css = fs.readFileSync(path.join(root, 'public', 'style.css'), 'utf8');
	const renderer = fs.readFileSync(path.join(root, 'public', 'aiRenderer.js'), 'utf8');
	const packageData = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
	for (const id of ['ai-open-btn', 'ai-modal', 'ai-input', 'ai-send-btn', 'ai-setup-btn', 'ai-setup-progress', 'ai-messages']) assert.match(html, new RegExp(`id="${id}"`));
	assert.match(html, /require\('\.\/aiRenderer\.js'\)/);
	assert.ok(packageData.build.files.includes('public/aiRenderer.js'));
	assert.match(css, /\.ai-title-button\s*\{[^}]*-webkit-app-region:\s*no-drag/s);
	assert.match(renderer, /Assistent öffnen/);
	assert.match(renderer, /closeRepetiererSettings\(aiOpen\)/);
});

test('Excel adapter reads only the requested student grade history', async () => {
	const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'repetierer-ai-test-'));
	try {
		const filePath = path.join(directory, 'class.xlsx');
		const workbook = new ExcelJS.Workbook();
		const sheet = workbook.addWorksheet('6a');
		sheet.getCell('A1').value = 'repetierer';
		sheet.getCell('A6').value = 'Anna Müller';
		sheet.getCell('B6').value = 5.5;
		sheet.getCell('K6').value = '15.09.2026';
		sheet.getCell('A7').value = 'Ben Müller';
		await workbook.xlsx.writeFile(filePath);
		await new Promise(resolve => excelReader.init(filePath, resolve));
		assert.deepEqual(excelReader.get_class_names(), ['6a']);
		assert.deepEqual(excelReader.read_grade_history('6a', 0).grades, [{ grade: 5.5, date: '15.09.2026', position: 1 }]);
		assert.deepEqual(excelReader.read_grade_history('6a', 1).grades, []);
	} finally { fs.rmSync(directory, { recursive: true, force: true }); }
});
