const test = require('node:test');
const assert = require('node:assert/strict');
const { findRoundFieldId } = require('../src/assessment/assessmentForm.js');

function formHtml(items) {
	const data = [null, [null, items]];
	return `<html><script>FB_PUBLIC_LOAD_DATA_ = ${JSON.stringify(data)};</script></html>`;
}

test('ermittelt die entry-ID des Felds Beurteilungsrunde', () => {
	const html = formHtml([
		[1, 'Vorname, Nachname', null, 0, [[111, null, 1]]],
		[2, 'Beurteilungsrunde', null, 0, [[1486315442, null, 1]]]
	]);
	assert.equal(findRoundFieldId(html), 'entry.1486315442');
});

test('erzwingt genau ein eindeutig benanntes Rundenfeld', () => {
	assert.throws(() => findRoundFieldId(formHtml([])), /genau ein Kurzantwortfeld/);
	assert.throws(() => findRoundFieldId(formHtml([
		[1, 'Runde', null, 0, [[111, null, 1]]],
		[2, 'Runden-ID', null, 0, [[222, null, 1]]]
	])), /genau ein Kurzantwortfeld/);
});
