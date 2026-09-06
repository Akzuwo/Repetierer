const test = require('node:test');
const assert = require('node:assert/strict');
const { deliverRound } = require('../src/assessment/assessmentDelivery.js');
const { buildResultMessage, escapeHtml } = require('../src/assessment/mailTemplates.js');
const { createSmtpMailer } = require('../src/assessment/smtpMailer.js');

function data() {
	const students = [{ id: 'a', name: 'Anna <Müller>' }, { id: 'b', name: 'Max Muster' }];
	const base = email => ({ studentEmail: email, studentAnswers: { c: 4 }, teacherAnswers: { c: 3 }, studentComment: '<script>', teacherComment: '& gut', mailStatus: 'not_sent' });
	return {
		classRecord: { students },
		round: { id: 'r', title: 'HS 2026', criteria: [{ id: 'c', label: 'Qualität' }], assessments: { a: base('anna@example.ch'), b: base('max@example.ch') } }
	};
}

test('escaped HTML enthält keine ungefilterten Schülerdaten', () => {
	const { classRecord, round } = data();
	const message = buildResultMessage({ student: classRecord.students[0], assessment: round.assessments.a, round, sender: { fromName: 'Lehrer', fromEmail: 'lehrer@example.ch' } });
	assert.match(message.html, /Anna &lt;Müller&gt;/);
	assert.doesNotMatch(message.html, /<script>/);
	assert.equal(escapeHtml('A&B'), 'A&amp;B');
	assert.throws(() => buildResultMessage({ student: classRecord.students[0], assessment: { studentEmail: 'x' }, round, sender: {} }), /ungültig/);
});

test('partieller SMTP-Fehler bleibt pro Schüler gespeichert', async () => {
	const { classRecord, round } = data();
	const statuses = [];
	const result = await deliverRound({
		classRecord, round,
		buildMessage: student => ({ to: student.id }),
		send: async message => { if (message.to === 'b') { const error = new Error('secret'); error.code = 'ETIMEDOUT'; throw error; } },
		updateStatus: async (studentId, status) => { round.assessments[studentId].mailStatus = status; statuses.push([studentId, status]); }
	});
	assert.equal(result.sent, 1);
	assert.equal(result.failed, 1);
	assert.equal(round.assessments.a.mailStatus, 'sent');
	assert.equal(round.assessments.b.mailStatus, 'failed');
	assert.deepEqual(statuses.filter(entry => entry[0] === 'a').map(entry => entry[1]), ['sending', 'sent']);
});

test('Retry sendet nur fehlgeschlagene Mails erneut', async () => {
	const { classRecord, round } = data();
	round.assessments.a.mailStatus = 'sent';
	round.assessments.b.mailStatus = 'failed';
	const sent = [];
	await deliverRound({ classRecord, round, mode: 'failed_only', buildMessage: student => student, send: async student => sent.push(student.id), updateStatus: async (id, status) => { round.assessments[id].mailStatus = status; } });
	assert.deepEqual(sent, ['b']);
});

test('SMTP-Adapter übergibt Brevo-Konfiguration und Nachrichten an Nodemailer', async () => {
	let transportOptions;
	let sentMessage;
	const fakeNodemailer = { createTransport(options) { transportOptions = options; return { verify: async () => true, sendMail: async message => { sentMessage = message; return { accepted: [message.to] }; } }; } };
	const mailer = createSmtpMailer({ host: 'smtp-relay.brevo.com', port: 587, secure: false, user: 'u', password: 'p' }, fakeNodemailer);
	assert.equal(await mailer.verify(), true);
	await mailer.send({ to: 'anna@example.ch' });
	assert.equal(transportOptions.auth.pass, 'p');
	assert.equal(sentMessage.to, 'anna@example.ch');
});
