const { isValidEmail } = require('./matching.js');

function buildInvitation(round, responderUrl, teacherName = 'Herr Schaufelberger') {
	const deadline = round.deadline ? ` bis am ${formatDate(round.deadline)}` : '';
	const subject = `Selbstbeurteilung Mitmachnote – ${round.title}`;
	const text = [
		'Liebe Schülerinnen und Schüler',
		'',
		`Bitte füllt${deadline} eure Selbstbeurteilung für die Mitmachnote aus:`,
		'',
		responderUrl,
		'',
		'Bitte verwendet eure eigene E-Mail-Adresse.',
		'',
		'Freundliche Grüsse',
		teacherName
	].join('\n');
	return { subject, text, complete: `Betreff: ${subject}\n\n${text}` };
}

function buildResultMessage({ student, assessment, round, sender }) {
	if (!assessment || !isValidEmail(assessment.studentEmail)) throw mailError('Die E-Mail-Adresse des Schülers ist ungültig.');
	if (!assessment.studentAnswers || !assessment.teacherAnswers) throw mailError('Die Beurteilung ist noch nicht vollständig.');
	const rows = round.criteria.map(criterion => {
		const self = assessment.studentAnswers[criterion.id];
		const teacher = assessment.teacherAnswers[criterion.id];
		return { label: criterion.label, self, teacher, difference: teacher - self };
	});
	const subject = `Mitmachnote – ${round.title} – ${student.name}`;
	const textRows = rows.map(row => `${row.label}: Selbstbeurteilung ${row.self}, Lehrerbeurteilung ${row.teacher}, Differenz ${signed(row.difference)}`);
	const text = [
		`Guten Tag ${student.name}`,
		'',
		`Hier ist deine persönliche Auswertung zur Mitmachnote (${round.title}).`,
		'',
		...textRows,
		'',
		`Kommentar Schüler: ${assessment.studentComment || '–'}`,
		`Kommentar Lehrer: ${assessment.teacherComment || '–'}`,
		'',
		'Freundliche Grüsse',
		sender.fromName
	].join('\n');
	const tableRows = rows.map(row => `<tr><td>${escapeHtml(row.label)}</td><td>${escapeHtml(row.self)}</td><td>${escapeHtml(row.teacher)}</td><td>${escapeHtml(signed(row.difference))}</td></tr>`).join('');
	const html = `<!doctype html><html><body style="margin:0;background:#f3f4f5;color:#202125;font-family:Arial,sans-serif"><div style="max-width:680px;margin:0 auto;padding:28px"><div style="background:#202125;color:#f5f5f5;border-radius:12px;padding:26px"><h1 style="margin:0 0 8px;color:#18a890;font-size:24px">Mitmachnote</h1><p style="margin:0 0 24px">${escapeHtml(round.title)} · ${escapeHtml(student.name)}</p><table style="border-collapse:collapse;width:100%;background:#fff;color:#202125;border-radius:8px;overflow:hidden"><thead><tr style="background:#18a890;color:#fff"><th style="padding:10px;text-align:left">Kriterium</th><th style="padding:10px">Schüler</th><th style="padding:10px">Lehrer</th><th style="padding:10px">Differenz</th></tr></thead><tbody>${tableRows}</tbody></table><h2 style="font-size:16px;margin:24px 0 6px;color:#18a890">Kommentar Schüler</h2><p style="white-space:pre-wrap;margin:0">${escapeHtml(assessment.studentComment || '–')}</p><h2 style="font-size:16px;margin:20px 0 6px;color:#18a890">Kommentar Lehrer</h2><p style="white-space:pre-wrap;margin:0">${escapeHtml(assessment.teacherComment || '–')}</p></div></div></body></html>`;
	return {
		to: assessment.studentEmail,
		from: { name: sender.fromName, address: sender.fromEmail },
		subject,
		text,
		html
	};
}

function escapeHtml(value) {
	return String(value === null || value === undefined ? '' : value)
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&#39;');
}

function signed(value) {
	const numeric = Number(value);
	return numeric > 0 ? `+${numeric}` : String(numeric);
}

function formatDate(value) {
	const [year, month, day] = String(value).split('-');
	return `${day}.${month}.${year}`;
}

function mailError(message) {
	const error = new Error(message);
	error.code = 'MAIL_VALIDATION_ERROR';
	return error;
}

module.exports = { buildInvitation, buildResultMessage, escapeHtml };
