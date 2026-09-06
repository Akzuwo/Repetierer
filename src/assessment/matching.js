const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function normalizeEmail(value) {
	return String(value || '').trim().toLowerCase();
}

function isValidEmail(value) {
	return EMAIL_PATTERN.test(normalizeEmail(value));
}

function normalizeName(value) {
	return String(value || '')
		.normalize('NFKD')
		.replace(/[\u0300-\u036f]/g, '')
		.toLowerCase()
		.replace(/[^a-z0-9]+/g, ' ')
		.trim()
		.replace(/\s+/g, ' ');
}

function levenshtein(left, right) {
	const a = normalizeName(left);
	const b = normalizeName(right);
	if (a === b) return 0;
	if (!a.length) return b.length;
	if (!b.length) return a.length;

	let previous = Array.from({ length: b.length + 1 }, (_, index) => index);
	for (let i = 1; i <= a.length; i++) {
		const current = [i];
		for (let j = 1; j <= b.length; j++) {
			current[j] = Math.min(
			current[j - 1] + 1,
			previous[j] + 1,
			previous[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)
			);
		}
		previous = current;
	}
	return previous[b.length];
}

function similarity(left, right) {
	const a = normalizeName(left);
	const b = normalizeName(right);
	const length = Math.max(a.length, b.length);
	return length ? 1 - (levenshtein(a, b) / length) : 1;
}

function matchResponse(response, students) {
	const email = normalizeEmail(response && response.email);
	const name = normalizeName(response && response.name);
	const candidates = Array.isArray(students) ? students : [];

	if (email) {
		const emailMatches = candidates.filter(student => normalizeEmail(student.email) === email);
		if (emailMatches.length === 1) return matched(emailMatches[0], 'email', 1);
	}

	if (name) {
		const exactMatches = candidates.filter(student => normalizeName(student.name) === name);
		if (exactMatches.length === 1) return matched(exactMatches[0], 'exact_name', 1);
	}

	const scored = candidates
		.map(student => ({ student, score: similarity(name, student.name) }))
		.filter(entry => entry.score >= 0.82)
		.sort((a, b) => b.score - a.score);

	if (scored.length && (scored.length === 1 || scored[0].score - scored[1].score >= 0.12)) {
		return matched(scored[0].student, 'plausible_name', scored[0].score);
	}

	return {
		status: 'unmatched',
		studentId: null,
		method: null,
		confidence: scored.length ? scored[0].score : 0,
		candidateIds: scored.slice(0, 3).map(entry => entry.student.id)
	};
}

function matched(student, method, confidence) {
	return {
		status: 'matched',
		studentId: student.id,
		method,
		confidence,
		candidateIds: [student.id]
	};
}

function selectLatestResponses(responses, preferredResponseIds = {}) {
	const grouped = new Map();
	for (const response of Array.isArray(responses) ? responses : []) {
		const email = normalizeEmail(response.email);
		const responseId = getResponseId(response);
		const key = email || `response:${responseId || Math.random()}`;
		if (!grouped.has(key)) grouped.set(key, []);
		grouped.get(key).push(response);
	}

	const selected = [];
	const duplicates = [];
	for (const [key, group] of grouped.entries()) {
		group.sort((a, b) => responseTime(b) - responseTime(a));
		const preferredId = preferredResponseIds[key];
		const chosen = group.find(response => getResponseId(response) === preferredId) || group[0];
		selected.push(chosen);
		if (group.length > 1) {
			duplicates.push({
				email: key.startsWith('response:') ? '' : key,
				selectedResponseId: getResponseId(chosen),
				responses: group.map(response => ({
					externalResponseId: getResponseId(response),
					timestamp: response.timestamp || response.lastSubmittedTime || response.submittedAt || response.createTime || null
				})),
				count: group.length
			});
		}
	}
	return { selected, duplicates };
}

function responseTime(response) {
	const value = Date.parse(response.timestamp || response.lastSubmittedTime || response.submittedAt || response.createTime || 0);
	return Number.isFinite(value) ? value : 0;
}

function getResponseId(response) {
	return String(response && (response.externalResponseId || response.responseId) || '');
}

module.exports = {
	isValidEmail,
	levenshtein,
	matchResponse,
	normalizeEmail,
	normalizeName,
	selectLatestResponses,
	similarity
};
