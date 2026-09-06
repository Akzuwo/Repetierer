const test = require('node:test');
const assert = require('node:assert/strict');
const { matchResponse, normalizeName, selectLatestResponses } = require('../src/assessment/matching.js');

const students = [
	{ id: '1', name: 'Anna Müller', email: 'anna@example.ch' },
	{ id: '2', name: 'Max Muster', email: '' },
	{ id: '3', name: 'Peter Meier', email: '' },
	{ id: '4', name: 'Petra Meier', email: '' }
];

test('normalisiert Namen mit Umlauten und Satzzeichen', () => {
	assert.equal(normalizeName('  Anna-Müller '), 'anna muller');
});

test('ordnet zuerst über bekannte E-Mail und danach über exakten Namen zu', () => {
	assert.equal(matchResponse({ name: 'Falscher Name', email: 'ANNA@example.ch' }, students).studentId, '1');
	assert.equal(matchResponse({ name: 'Max Muster', email: 'neu@example.ch' }, students).studentId, '2');
});

test('erzwingt bei mehrdeutig plausiblem Namen keine Zuordnung', () => {
	const result = matchResponse({ name: 'P. Meier', email: '' }, students);
	assert.equal(result.status, 'unmatched');
});

test('verwendet bei doppelter E-Mail die neueste Antwort und meldet Duplikate', () => {
	const result = selectLatestResponses([
		{ responseId: 'old', email: 'a@example.ch', lastSubmittedTime: '2026-01-01T10:00:00Z' },
		{ responseId: 'new', email: 'A@example.ch', lastSubmittedTime: '2026-01-02T10:00:00Z' }
	]);
	assert.equal(result.selected[0].responseId, 'new');
	assert.equal(result.duplicates[0].count, 2);
});
