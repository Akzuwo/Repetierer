const ROUND_FIELD_TITLES = new Set([
	'beurteilungsrunde',
	'runden id',
	'runde',
	'round id',
	'roundid'
]);

async function discoverRoundFieldId(formUrl, options = {}) {
	const fetch = options.fetch || global.fetch;
	const timeoutMs = options.timeoutMs || 15000;
	const controller = new AbortController();
	const timer = setTimeout(() => controller.abort(), timeoutMs);
	let response;
	try {
		response = await fetch(formUrl, { signal: controller.signal, redirect: 'follow' });
	} catch (error) {
		if (error && error.name === 'AbortError') throw formError('Zeitüberschreitung beim Prüfen des Umfrageformulars.', 'ASSESSMENT_FORM_TIMEOUT');
		throw formError('Das Umfrageformular ist derzeit nicht erreichbar.', 'ASSESSMENT_FORM_NETWORK_ERROR');
	} finally {
		clearTimeout(timer);
	}
	if (!response.ok) throw formError('Das Umfrageformular konnte nicht geladen werden.', 'ASSESSMENT_FORM_HTTP_ERROR');
	const html = await response.text();
	if (html.length > 2 * 1024 * 1024) throw formError('Das Umfrageformular ist unerwartet groß.', 'ASSESSMENT_FORM_SCHEMA_INVALID');
	return findRoundFieldId(html);
}

function findRoundFieldId(html) {
	const match = String(html || '').match(/FB_PUBLIC_LOAD_DATA_\s*=\s*(\[[\s\S]*?\]);\s*<\/script>/);
	if (!match) throw formError('Die Formularfelder konnten nicht gelesen werden.', 'ASSESSMENT_FORM_SCHEMA_INVALID');
	let data;
	try { data = JSON.parse(match[1]); }
	catch (error) { throw formError('Die Formularfelder enthalten ungültige Daten.', 'ASSESSMENT_FORM_SCHEMA_INVALID'); }
	const items = data && data[1] && data[1][1];
	if (!Array.isArray(items)) throw formError('Im Formular wurden keine Fragen gefunden.', 'ASSESSMENT_FORM_SCHEMA_INVALID');
	const candidates = items.filter(item => Array.isArray(item) && ROUND_FIELD_TITLES.has(normalizeTitle(item[1])));
	if (candidates.length !== 1) {
		throw formError('Im Formular muss genau ein Kurzantwortfeld „Beurteilungsrunde“ vorhanden sein.', 'ASSESSMENT_ROUND_FIELD_MISSING');
	}
	const questionId = candidates[0][4] && candidates[0][4][0] && candidates[0][4][0][0];
	if (!Number.isInteger(questionId) || questionId <= 0) throw formError('Die Feld-ID der Beurteilungsrunde ist ungültig.', 'ASSESSMENT_FORM_SCHEMA_INVALID');
	return `entry.${questionId}`;
}

function normalizeTitle(value) {
	return String(value || '').normalize('NFKD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

function formError(message, code) {
	const error = new Error(message);
	error.code = code;
	return error;
}

module.exports = { discoverRoundFieldId, findRoundFieldId };
