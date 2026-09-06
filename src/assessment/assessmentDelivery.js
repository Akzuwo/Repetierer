const { isValidEmail } = require('./matching.js');

function validateRoundForDelivery(classRecord, round) {
	const incomplete = [];
	for (const student of classRecord.students.filter(item => item.active !== false)) {
		const assessment = round.assessments[student.id];
		const reasons = [];
		if (!assessment || !assessment.studentAnswers) reasons.push('Selbstbeurteilung fehlt');
		if (!assessment || !assessment.teacherAnswers) reasons.push('Lehrerbeurteilung fehlt');
		if (!assessment || !isValidEmail(assessment.studentEmail)) reasons.push('E-Mail ungültig');
		if (reasons.length) incomplete.push({ studentId: student.id, name: student.name, reasons });
	}
	return { valid: incomplete.length === 0, incomplete };
}

async function deliverRound({ classRecord, round, mode = 'pending', buildMessage, send, updateStatus }) {
	const validation = validateRoundForDelivery(classRecord, round);
	if (mode !== 'failed_only' && !validation.valid) {
		const error = new Error('Der Klassensatz ist noch nicht vollständig versandbereit.');
		error.code = 'ROUND_INCOMPLETE';
		error.incomplete = validation.incomplete;
		throw error;
	}
	const targets = classRecord.students.filter(student => {
		if (student.active === false) return false;
		const status = round.assessments[student.id] && round.assessments[student.id].mailStatus;
		return mode === 'failed_only' ? status === 'failed' : status !== 'sent';
	});
	const results = [];
	for (const student of targets) {
		const assessment = round.assessments[student.id];
		await updateStatus(student.id, 'sending', { incrementAttempt: true });
		try {
			await send(buildMessage(student, assessment));
			await updateStatus(student.id, 'sent', {});
			results.push({ studentId: student.id, status: 'sent' });
		} catch (error) {
			const reason = sanitizeMailError(error);
			await updateStatus(student.id, 'failed', { error: reason });
			results.push({ studentId: student.id, status: 'failed', error: reason });
		}
	}
	return {
		total: targets.length,
		sent: results.filter(result => result.status === 'sent').length,
		failed: results.filter(result => result.status === 'failed').length,
		results
	};
}

function sanitizeMailError(error) {
	const code = String(error && error.code || '').toUpperCase();
	if (code.includes('AUTH') || code === 'EAUTH') return 'SMTP-Anmeldung fehlgeschlagen.';
	if (code.includes('TIMEOUT') || code === 'ETIMEDOUT') return 'SMTP-Zeitüberschreitung.';
	if (code === 'ECONNECTION' || code === 'ECONNREFUSED' || code === 'ENETUNREACH') return 'SMTP-Verbindung fehlgeschlagen.';
	if (error && error.code === 'MAIL_VALIDATION_ERROR') return error.message;
	return 'Die E-Mail konnte nicht versendet werden.';
}

module.exports = { deliverRound, sanitizeMailError, validateRoundForDelivery };
