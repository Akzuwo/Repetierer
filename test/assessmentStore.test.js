const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { AssessmentStore, DEFAULT_CRITERIA, SCHEMA_VERSION } = require('../src/assessment/assessmentStore.js');

function fixture() {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'repetierer-assessment-'));
	return { dir, dataPath: path.join(dir, 'mitmachnoten.json'), workbook: path.join(dir, 'classes.xlsx') };
}

const students = [{ id: 0, name: 'Anna Müller' }, { id: 1, name: 'Max Muster' }];
const answers = value => Object.fromEntries(DEFAULT_CRITERIA.map(criterion => [criterion.id, value]));

test('lädt fehlende Alt-Daten rückwärtskompatibel und speichert eine Runde', () => {
	const files = fixture();
	const store = new AssessmentStore(files.dataPath);
	const round = store.createRound(files.workbook, '3a', students, { title: 'HS 2026', deadline: '2026-12-18' });
	const reloaded = new AssessmentStore(files.dataPath);
	const context = reloaded.getClassContext(files.workbook, '3a', students);
	assert.equal(reloaded.data.schemaVersion, SCHEMA_VERSION);
	assert.equal(context.rounds[0].id, round.id);
	assert.match(context.rounds[0].externalRoundId, /^assessment-3a-hs-2026-[a-f0-9]{6}$/);
	assert.equal(context.rounds[0].criteria.length, 7);
	assert.deepEqual(context.rounds[0].criteria[0].scaleLabels, ['(Fast) nie', 'Ab und zu', 'Manchmal', 'Häufig', 'Sehr häufig']);
	assert.deepEqual(context.rounds[0].criteria.at(-1).scaleLabels, ['Mangelhaft', 'Genügend', 'Recht', 'Gut', 'Sehr gut']);
	assert.equal(context.students.length, 2);
	fs.rmSync(files.dir, { recursive: true, force: true });
});

test('migriert eine alte Runde ohne Formularmetadaten oder Bewertungsverlust', () => {
	const files = fixture();
	fs.writeFileSync(files.dataPath, JSON.stringify({ schemaVersion: 1, workbooks: { old: { classes: { c: {
		id: 'class-old', name: '3a', students: [{ id: 'student-1', name: 'Anna Müller', active: true }], rounds: [{
			id: 'round-abcdef', title: 'HS 2025', createdAt: '2025-08-01T00:00:00Z', form: { formId: 'old-google-form' },
			criteria: [{ id: 'participation', label: 'Aktive Beteiligung' }], assessments: { 'student-1': {
				studentId: 'student-1', studentEmail: 'anna@example.ch', studentAnswers: { participation: 4 },
				studentComment: '', submittedAt: '2025-09-01T00:00:00Z', responseId: 'legacy-response', mailStatus: 'not_sent'
			} }
		}]
	} } } } }));
	const store = new AssessmentStore(files.dataPath);
	const migrated = store.data.workbooks.old.classes.c.rounds[0];
	assert.equal(store.data.schemaVersion, SCHEMA_VERSION);
	assert.equal(Object.hasOwn(migrated, 'form'), false);
	assert.equal(migrated.externalResponses[0].externalResponseId, 'legacy-response');
	assert.equal(migrated.assessments['student-1'].studentAnswers.participation, 4);
	assert.equal(migrated.criteria.length, 1);
	assert.equal(migrated.assessments['student-1'].studentAdditionalQuality, '');
	fs.rmSync(files.dir, { recursive: true, force: true });
});

test('stellt nur leere Runden mit altem Standardraster auf den Originalbogen um', () => {
	const files = fixture();
	const assessments = { anna: { studentId: 'anna', studentAnswers: null, teacherAnswers: null } };
	fs.writeFileSync(files.dataPath, JSON.stringify({ schemaVersion: 2, workbooks: { old: { classes: { c: {
		id: 'class-old', name: '3a', students: [{ id: 'anna', name: 'Anna Müller', active: true }], rounds: [{
			id: 'empty-round', title: 'FS 2027', criteria: [
				{ id: 'participation' }, { id: 'preparation' }, { id: 'quality' }, { id: 'reliability' }
			], assessments
		}]
	} } } } }));
	const store = new AssessmentStore(files.dataPath);
	const round = store.data.workbooks.old.classes.c.rounds[0];
	assert.deepEqual(round.criteria.map(criterion => criterion.id), DEFAULT_CRITERIA.map(criterion => criterion.id));
	fs.rmSync(files.dir, { recursive: true, force: true });
});

test('synchronisiert Antworten lokal und speichert Lehrerwerte getrennt', () => {
	const files = fixture();
	const store = new AssessmentStore(files.dataPath);
	const round = store.createRound(files.workbook, '3a', students, { title: 'HS 2026' });
	store.applyResponses(files.workbook, '3a', students, round.id, [{
		responseId: 'r1', name: 'Anna Müller', email: 'anna@example.ch', lastSubmittedTime: '2026-09-01T12:00:00Z',
		answers: answers(4), additionalQuality: 'Spezifisches Fachwissen', comment: 'Selbstkommentar'
	}]);
	const context = store.getClassContext(files.workbook, '3a', students);
	const anna = context.students.find(student => student.name === 'Anna Müller');
	store.saveTeacherAssessment(files.workbook, '3a', students, round.id, anna.id, answers(3), 'Lehrerkommentar');
	const reloaded = new AssessmentStore(files.dataPath);
	const stored = reloaded.getRound(files.workbook, '3a', students, round.id).round.assessments[anna.id];
	assert.equal(stored.studentAnswers.participation_frequency, 4);
	assert.equal(stored.teacherAnswers.participation_frequency, 3);
	assert.equal(stored.studentAdditionalQuality, 'Spezifisches Fachwissen');
	assert.equal(stored.studentEmail, 'anna@example.ch');
	fs.rmSync(files.dir, { recursive: true, force: true });
});

test('importiert Antwort-IDs nur einmal und erlaubt die Auswahl einer älteren Mehrfachantwort', () => {
	const files = fixture();
	const store = new AssessmentStore(files.dataPath);
	const round = store.createRound(files.workbook, '3a', students, { title: 'FS 2027' });
	store.applyResponses(files.workbook, '3a', students, round.id, [
		{ externalResponseId: 'old', timestamp: '2026-09-01T10:00:00Z', name: 'Anna Müller', email: 'anna@example.ch', answers: answers(2) },
		{ externalResponseId: 'new', timestamp: '2026-09-02T10:00:00Z', name: 'Anna Müller', email: 'anna@example.ch', answers: answers(4) }
	]);
	store.applyResponses(files.workbook, '3a', students, round.id, [
		{ externalResponseId: 'new', timestamp: '2026-09-02T10:00:00Z', name: 'Anna Müller', email: 'anna@example.ch', answers: answers(4) }
	]);
	let current = store.getRound(files.workbook, '3a', students, round.id);
	const anna = current.classRecord.students.find(student => student.name === 'Anna Müller');
	assert.equal(current.round.externalResponses.length, 2);
	assert.equal(current.round.assessments[anna.id].studentAnswers.participation_frequency, 4);
	store.selectResponse(files.workbook, '3a', students, round.id, 'old');
	current = store.getRound(files.workbook, '3a', students, round.id);
	assert.equal(current.round.assessments[anna.id].studentAnswers.participation_frequency, 2);
	fs.rmSync(files.dir, { recursive: true, force: true });
});

test('löscht ausschließlich die gewählte lokale Runde und bleibt nach Neustart gelöscht', () => {
	const files = fixture();
	const store = new AssessmentStore(files.dataPath);
	const keep = store.createRound(files.workbook, '3a', students, { title: 'HS 2026' });
	const remove = store.createRound(files.workbook, '3a', students, { title: 'FS 2027' });
	const contextBefore = store.getClassContext(files.workbook, '3a', students);
	const anna = contextBefore.students.find(student => student.name === 'Anna Müller');
	store.setMailStatus(files.workbook, '3a', students, remove.id, anna.id, 'sent', {});
	const result = store.deleteRound(files.workbook, '3a', students, remove.id);
	assert.equal(result.hadSentResults, true);
	const reloaded = new AssessmentStore(files.dataPath);
	const context = reloaded.getClassContext(files.workbook, '3a', students);
	assert.deepEqual(context.rounds.map(round => round.id), [keep.id]);
	assert.deepEqual(context.students.map(student => student.name), ['Anna Müller', 'Max Muster']);
	assert.equal(fs.existsSync(files.workbook), false);
	fs.rmSync(files.dir, { recursive: true, force: true });
});
