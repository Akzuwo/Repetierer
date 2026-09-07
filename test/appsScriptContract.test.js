const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

test('Apps Script bildet das Originalraster aus ausgeschriebenen Forms-Antworten ab', () => {
	const source = fs.readFileSync(path.join(__dirname, '..', 'docs', 'google-apps-script', 'Code.gs'), 'utf8');
	const context = {
		Utilities: {
			Charset: { UTF_8: 'UTF_8' },
			DigestAlgorithm: { SHA_256: 'SHA_256' },
			computeDigest: () => [1, 2, 3]
		}
	};
	vm.createContext(context);
	vm.runInContext(source, context);
	const headers = [
		'Zeitstempel', 'Beurteilungsrunde', 'Vorname, Nachname', 'E-Mail Adresse',
		'Beteiligung mit Äusserungen', 'Falsch, unbefriedigend oder nicht ausreichend',
		'Originell, aber unpassend im Lektionsverlauf', 'Korrekt, aber stichwortartig',
		'Passend und im Lektionsverlauf weiterführend',
		'Passend, eigenständig und in eine neue, interessante Richtung führend',
		'Zusätzliche Qualität meiner Äusserungen', 'Teilnahme insgesamt', 'Sonstige Bemerkungen'
	];
	const row = [
		'2026-09-07T12:00:00Z', 'assessment-3a-hs-abc123', 'Anna Müller', 'ANNA@EXAMPLE.CH',
		'Häufig', '(Fast) nie', 'Ab und zu', 'Manchmal', '4 – Häufig', 'Sehr häufig',
		'Spezifisches Fachwissen', 'Gut', 'Gutes Klassenklima'
	];
	const columns = context.mapColumns(headers);
	const response = context.mapResponse(row, columns);
	assert.equal(response.email, 'anna@example.ch');
	assert.deepEqual(JSON.parse(JSON.stringify(response.answers)), {
		participation_frequency: 4,
		insufficient_contributions: 1,
		original_unsuitable_contributions: 2,
		correct_brief_contributions: 3,
		advancing_contributions: 4,
		independent_contributions: 5,
		overall_participation: 4
	});
	assert.equal(response.additionalQuality, 'Spezifisches Fachwissen');
	assert.equal(response.comment, 'Gutes Klassenklima');
});
