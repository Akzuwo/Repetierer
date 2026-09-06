const { isValidEmail } = require('./matching.js');

const RESPONSE_FIELDS = Object.freeze({
	responseId: 'responseId',
	timestamp: 'timestamp',
	roundId: 'roundId',
	name: 'name',
	email: 'email',
	answers: 'answers',
	comment: 'comment'
});
const transportByEndpoint = new Map();

class AssessmentApiClient {
	constructor(config, options = {}) {
		this.config = config;
		this.fetch = options.fetch || global.fetch;
		this.timeoutMs = options.timeoutMs || 15000;
	}

	async checkStatus() {
		await this.request({ action: 'responses', roundId: '__repetierer_connection_test__', knownResponseIds: [] });
		return { reachable: true };
	}

	async fetchResponses({ externalRoundId, criteria, knownResponseIds = [] }) {
		const body = await this.request({
			action: 'responses',
			roundId: externalRoundId,
			knownResponseIds: Array.from(new Set(knownResponseIds.map(String))).slice(-5000)
		});
		if (!Array.isArray(body.responses)) throw apiError('Die Antwort der Mitmachnoten-API enthält keine Antwortliste.', 'ASSESSMENT_API_SCHEMA_INVALID');
		const responses = [];
		const invalidResponses = [];
		let ignoredRoundCount = 0;
		for (let index = 0; index < body.responses.length; index++) {
			const raw = body.responses[index];
			if (raw && String(raw[RESPONSE_FIELDS.roundId] || '') !== externalRoundId) {
				ignoredRoundCount++;
				continue;
			}
			try {
				responses.push(validateResponse(raw, externalRoundId, criteria));
			} catch (error) {
				invalidResponses.push({ index, responseId: safeResponseId(raw), reason: error.message });
			}
		}
		return { responses, invalidResponses, ignoredRoundCount };
	}

	async request(payload) {
		const preferred = transportByEndpoint.get(this.config.apiUrl);
		if (preferred === 'get') return this.performRequest('get', payload);
		try {
			const result = await this.performRequest('post', payload);
			transportByEndpoint.set(this.config.apiUrl, 'post');
			return result;
		} catch (error) {
			if (!isPostUnsupported(error)) throw error;
			const result = await this.performRequest('get', payload);
			transportByEndpoint.set(this.config.apiUrl, 'get');
			return result;
		}
	}

	async performRequest(method, payload) {
		const controller = new AbortController();
		const timer = setTimeout(() => controller.abort(), this.timeoutMs);
		let response;
		try {
			const request = buildRequest(this.config, method, payload);
			response = await this.fetch(request.url, {
				method: request.method,
				headers: request.headers,
				body: request.body,
				signal: controller.signal,
				redirect: 'follow'
			});
		} catch (error) {
			if (error && error.name === 'AbortError') throw apiError('Zeitüberschreitung beim Abrufen der Mitmachnoten.', 'ASSESSMENT_API_TIMEOUT');
			throw apiError('Die Mitmachnoten-API ist derzeit nicht erreichbar.', 'ASSESSMENT_API_NETWORK_ERROR');
		} finally {
			clearTimeout(timer);
		}
		let body;
		let rawBody = '';
		try {
			rawBody = await response.text();
			body = JSON.parse(rawBody);
		} catch (error) {
			const invalid = apiError('Die Mitmachnoten-API hat keine gültige JSON-Antwort geliefert.', 'ASSESSMENT_API_INVALID_JSON');
			invalid.allowGetFallback = method === 'post' && /(?:Script-Funktion nicht gefunden|Script function not found):\s*doPost/i.test(rawBody);
			throw invalid;
		}
		if (response.status === 401 || response.status === 403) throw apiError('Der API-Schlüssel wurde abgelehnt.', 'ASSESSMENT_API_UNAUTHORIZED');
		if (!response.ok) throw apiError('Die Mitmachnoten-API konnte die Anfrage nicht verarbeiten.', 'ASSESSMENT_API_HTTP_ERROR');
		if (!body || body.success !== true) throw apiError(publicApiMessage(body), 'ASSESSMENT_API_REJECTED');
		return body;
	}
}

function buildRequest(config, method, payload) {
	if (method === 'get') {
		const url = new URL(config.apiUrl);
		url.searchParams.set('apiKey', config.apiKey);
		if (payload.roundId) url.searchParams.set('roundId', payload.roundId);
		return { url: url.toString(), method: 'GET', headers: { Accept: 'application/json' } };
	}
	return {
		url: config.apiUrl,
		method: 'POST',
		headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
		body: JSON.stringify(Object.assign({}, payload, { apiKey: config.apiKey }))
	};
}

function isPostUnsupported(error) {
	return !!(error && error.allowGetFallback);
}

function validateResponse(raw, externalRoundId, criteria) {
	if (!raw || typeof raw !== 'object' || Array.isArray(raw)) throw schemaError('Antwort ist kein Objekt.');
	const externalResponseId = String(raw[RESPONSE_FIELDS.responseId] || '').trim();
	if (!externalResponseId || externalResponseId.length > 200) throw schemaError('responseId fehlt oder ist ungültig.');
	const timestamp = String(raw[RESPONSE_FIELDS.timestamp] || '').trim();
	if (!timestamp || Number.isNaN(Date.parse(timestamp))) throw schemaError('timestamp fehlt oder ist ungültig.');
	if (String(raw[RESPONSE_FIELDS.roundId] || '') !== externalRoundId) throw schemaError('roundId stimmt nicht mit der Runde überein.');
	const name = String(raw[RESPONSE_FIELDS.name] || '').trim();
	if (!name || name.length > 200) throw schemaError('Name fehlt oder ist ungültig.');
	const email = String(raw[RESPONSE_FIELDS.email] || '').trim();
	if (!isValidEmail(email)) throw schemaError('E-Mail-Adresse fehlt oder ist ungültig.');
	const sourceAnswers = raw[RESPONSE_FIELDS.answers];
	if (!sourceAnswers || typeof sourceAnswers !== 'object' || Array.isArray(sourceAnswers)) throw schemaError('Bewertungswerte fehlen.');
	const answers = {};
	for (const criterion of Array.isArray(criteria) ? criteria : []) {
		const value = Number(sourceAnswers[criterion.id]);
		if (!Number.isInteger(value) || value < 1 || value > 5) throw schemaError(`Bewertungswert für ${criterion.id} ist ungültig.`);
		answers[criterion.id] = value;
	}
	const comment = raw[RESPONSE_FIELDS.comment];
	if (comment !== undefined && comment !== null && typeof comment !== 'string') throw schemaError('Kommentar ist ungültig.');
	return { externalResponseId, timestamp, externalRoundId, name, email, answers, comment: String(comment || '').trim() };
}

function safeResponseId(raw) {
	return raw && typeof raw === 'object' ? String(raw[RESPONSE_FIELDS.responseId] || '').slice(0, 200) : '';
}

function publicApiMessage(body) {
	const code = body && typeof body.error === 'string' ? body.error : '';
	if (code === 'unauthorized') return 'Der API-Schlüssel wurde abgelehnt.';
	return 'Die Mitmachnoten-API hat die Anfrage abgelehnt.';
}

function schemaError(message) {
	return apiError(message, 'ASSESSMENT_RESPONSE_INVALID');
}

function apiError(message, code) {
	const error = new Error(message);
	error.code = code;
	return error;
}

module.exports = { AssessmentApiClient, RESPONSE_FIELDS, buildRequest, validateResponse };
