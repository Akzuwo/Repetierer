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
const {exportIntegrationFile,parsePortableConfig,readIntegrationValues} = require('../src/assessment/integrationConfig.js');

const teacherToken = 'rt_' + 'a'.repeat(43);
const assessmentEnv = `ASSESSMENT_API_URL=https://timolu.ch/app/services/repetierer/api\nASSESSMENT_TEACHER_TOKEN=${teacherToken}\n`;
const brevoEnv = `BREVO_SMTP_HOST=smtp-relay.brevo.com\nBREVO_SMTP_PORT=587\nBREVO_SMTP_USER=user@example.ch\nBREVO_SMTP_PASSWORD="secret value"\nBREVO_FROM_EMAIL=teacher@example.ch\nBREVO_FROM_NAME=Herr Schaufelberger\n`;

test('ein portables Paket überträgt SMTP und Cloudflare zusammen und behält die Quelldatei', async () => {
	const dir=fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-bundle-'));
	const file=path.join(dir,'teacher.repetierer'),internal=path.join(dir,'internal.env');
	const values=parseEnv(assessmentEnv+brevoEnv); values.ASSESSMENT_TEACHER_TOKEN='encrypted:bG9jYWw=';
	exportIntegrationFile(file,values,{revealTeacherToken:()=>teacherToken});
	const before=fs.readFileSync(file); assert.equal(parsePortableConfig(before).ASSESSMENT_TEACHER_TOKEN,teacherToken);
	let verified=false;
	const result=await importIntegrationEnv(file,internal,{protectTeacherToken:token=>'encrypted:'+Buffer.from(token).toString('base64'),verify:async ({assessment,brevo})=>{assert.equal(assessment.teacherToken,teacherToken);assert.equal(brevo.password,'secret value');verified=true;}});
	assert.ok(verified);assert.equal(result.originalDeleted,false);assert.deepEqual(fs.readFileSync(file),before);
	assert.equal(readBrevoConfig(internal).fromEmail,'teacher@example.ch');assert.match(readAssessmentConfig(internal).teacherToken,/^encrypted:/);
	assert.equal(Object.keys(readIntegrationValues(internal)).length,8);
});
test('ungültige und gerätegebundene Pakete verändern die bestehende Konfiguration nicht', async () => {
	const dir=fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-bundle-invalid-')),file=path.join(dir,'invalid.repetierer'),internal=path.join(dir,'internal.env');
	fs.writeFileSync(internal,assessmentEnv+brevoEnv);const before=fs.readFileSync(internal);
	for(const bundle of [{format:'repetierer-settings',version:2,values:parseEnv(assessmentEnv+brevoEnv)},{format:'repetierer-settings',version:1,values:parseEnv(assessmentEnv)},{format:'repetierer-settings',version:1,values:{...parseEnv(assessmentEnv+brevoEnv),ASSESSMENT_TEACHER_TOKEN:'encrypted:abc='}}]) {
		fs.writeFileSync(file,JSON.stringify(bundle));await assert.rejects(importIntegrationEnv(file,internal));assert.deepEqual(fs.readFileSync(internal),before);assert.ok(fs.existsSync(file));
	}
});
test('Paketexport benötigt beide Verbindungen und exportiert keine fremden Schlüssel',()=>{
	assert.throws(()=>exportIntegrationFile('unused.repetierer',parseEnv(assessmentEnv)),/Fehlende/);
	assert.throws(()=>parsePortableConfig(JSON.stringify({format:'repetierer-settings',version:1,values:{...parseEnv(assessmentEnv+brevoEnv),COMMAND:'run me'}})),/unbekannte/);
});
test('fehlgeschlagene API-Prüfung oder lokaler Schlüsselspeicher bewahren vorhandene Daten',async()=>{
	const dir=fs.mkdtempSync(path.join(os.tmpdir(),'repetierer-bundle-verify-')),file=path.join(dir,'teacher.repetierer'),internal=path.join(dir,'internal.env');
	exportIntegrationFile(file,{...parseEnv(assessmentEnv+brevoEnv),PRIVATE_ADMIN_KEY:'never-export'});
	assert.equal(fs.readFileSync(file,'utf8').includes('never-export'),false);
	fs.writeFileSync(internal,brevoEnv);const before=fs.readFileSync(internal),source=fs.readFileSync(file);
	for(const options of [{verify:async()=>{throw new Error('offline');}},{protectTeacherToken:()=>{throw new Error('keyring unavailable');}}]) {
		await assert.rejects(importIntegrationEnv(file,internal,options));assert.deepEqual(fs.readFileSync(internal),before);assert.deepEqual(fs.readFileSync(file),source);
	}
});

test('parst nur Key/Value, validiert beide Gruppen und baut einen vorbefüllten Rundenlink', () => {
	assert.equal(validateAssessmentConfig(parseEnv(assessmentEnv)).fallbackApiUrl, 'https://akzuwo.ch/app/services/repetierer/api');
	assert.equal(validateBrevoConfig(parseEnv(brevoEnv)).port, 587);
	assert.throws(() => validateAssessmentConfig(parseEnv('ASSESSMENT_API_URL=https://example.ch')), /Fehlende/);
	assert.throws(() => parseEnv('require("child_process")'), /Ungültige Zeile/);
	const url = buildRoundFormUrl(validateAssessmentConfig(parseEnv(assessmentEnv)), 'assessment-3a-hs-a8f2d1');
	assert.equal(new URL(url).searchParams.get('round'), 'assessment-3a-hs-a8f2d1');
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
	assert.equal(readAssessmentConfig(destination).teacherToken, teacherToken);
	assert.equal(readBrevoConfig(destination).fromEmail, 'teacher@example.ch');
	fs.rmSync(dir, { recursive: true, force: true });
});

test('liest alte Brevo-Dateien weiter und behält Originale bei Importfehlern', async () => {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'repetierer-config-legacy-'));
	const legacy = path.join(dir, 'brevo-smtp.env');
	fs.writeFileSync(legacy, brevoEnv);
	assert.equal(readBrevoConfig(path.join(dir, 'missing.env'), legacy).fromEmail, 'teacher@example.ch');
	const invalid = path.join(dir, 'invalid.env');
	fs.writeFileSync(invalid, 'ASSESSMENT_API_URL=http://unsafe.example');
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
