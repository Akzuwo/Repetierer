const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const {
	buildRoundFormUrl,
	importIntegrationEnv,
	parseEnv,
	readAssessmentConfig,
	readBrevoConfig,
	validateAssessmentConfig,
	validateBrevoConfig
} = require('../src/assessment/integrationConfig.js');

const assessmentEnv = `ASSESSMENT_FORM_URL=https://docs.google.com/forms/d/e/example/viewform\nASSESSMENT_API_URL=https://script.google.com/macros/s/example/exec\nASSESSMENT_API_KEY=long-secret-key\nASSESSMENT_ROUND_FIELD_ID=entry.123456\n`;
const brevoEnv = `BREVO_SMTP_HOST=smtp-relay.brevo.com\nBREVO_SMTP_PORT=587\nBREVO_SMTP_USER=user@example.ch\nBREVO_SMTP_PASSWORD="secret value"\nBREVO_FROM_EMAIL=teacher@example.ch\nBREVO_FROM_NAME=Herr Schaufelberger\n`;

test('parst nur Key/Value, validiert beide Gruppen und baut einen vorbefüllten Rundenlink', () => {
	assert.equal(validateAssessmentConfig(parseEnv(assessmentEnv)).roundFieldId, 'entry.123456');
	assert.equal(validateBrevoConfig(parseEnv(brevoEnv)).port, 587);
	assert.throws(() => validateAssessmentConfig(parseEnv('ASSESSMENT_API_URL=https://example.ch')), /Fehlende/);
	assert.throws(() => parseEnv('require("child_process")'), /Ungültige Zeile/);
	const url = buildRoundFormUrl(validateAssessmentConfig(parseEnv(assessmentEnv)), 'assessment-3a-hs-a8f2d1');
	assert.equal(new URL(url).searchParams.get('entry.123456'), 'assessment-3a-hs-a8f2d1');
});

test('führt Teilkonfigurationen zusammen und löscht das Original erst nach Verifikation', async () => {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'repetierer-config-'));
	const destination = path.join(dir, 'internal', 'repetierer.env');
	const first = path.join(dir, 'assessment.env');
	const second = path.join(dir, 'brevo.env');
	fs.writeFileSync(first, assessmentEnv);
	await importIntegrationEnv(first, destination);
	fs.writeFileSync(second, brevoEnv);
	const result = await importIntegrationEnv(second, destination);
	assert.equal(result.originalDeleted, true);
	assert.equal(fs.existsSync(second), false);
	assert.equal(readAssessmentConfig(destination).apiKey, 'long-secret-key');
	assert.equal(readBrevoConfig(destination).fromEmail, 'teacher@example.ch');
	fs.rmSync(dir, { recursive: true, force: true });
});

test('liest alte Brevo-Dateien weiter und behält Originale bei Importfehlern', async () => {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'repetierer-config-legacy-'));
	const legacy = path.join(dir, 'brevo-smtp.env');
	fs.writeFileSync(legacy, brevoEnv);
	assert.equal(readBrevoConfig(path.join(dir, 'missing.env'), legacy).fromEmail, 'teacher@example.ch');
	const invalid = path.join(dir, 'invalid.env');
	fs.writeFileSync(invalid, 'ASSESSMENT_FORM_URL=http://unsafe.example');
	await assert.rejects(importIntegrationEnv(invalid, path.join(dir, 'repetierer.env')));
	assert.equal(fs.existsSync(invalid), true);
	fs.rmSync(dir, { recursive: true, force: true });
});

test('löscht das Original bei fehlgeschlagenem Schreiben nicht', async () => {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'repetierer-config-write-'));
	const source = path.join(dir, 'download.env');
	const destination = path.join(dir, 'destination.env');
	fs.writeFileSync(source, assessmentEnv);
	fs.mkdirSync(destination);
	await assert.rejects(importIntegrationEnv(source, destination));
	assert.equal(fs.existsSync(source), true);
	fs.rmSync(dir, { recursive: true, force: true });
});

test('behält bei einem Fehler nur beim abschließenden Löschen beide verifizierten Kopien', async () => {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'repetierer-config-delete-'));
	const source = path.join(dir, 'download.env');
	const destination = path.join(dir, 'internal', 'repetierer.env');
	fs.writeFileSync(source, assessmentEnv + brevoEnv);
	const originalUnlink = fs.unlinkSync;
	fs.unlinkSync = function(filePath) {
		if (path.resolve(filePath) === path.resolve(source)) throw new Error('locked');
		return originalUnlink.call(fs, filePath);
	};
	try {
		const result = await importIntegrationEnv(source, destination);
		assert.equal(result.originalDeleted, false);
		assert.equal(fs.existsSync(source), true);
		assert.equal(readAssessmentConfig(destination).configured, undefined);
	} finally {
		fs.unlinkSync = originalUnlink;
		fs.rmSync(dir, { recursive: true, force: true });
	}
});
