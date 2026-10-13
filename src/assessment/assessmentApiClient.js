const { isValidEmail } = require('./matching.js');

const RESPONSE_FIELDS = Object.freeze({
	responseId: 'responseId',
	timestamp: 'timestamp',
	roundId: 'roundId',
	name: 'name',
	email: 'email',
	answers: 'answers',
	additionalQuality: 'additionalQuality',
	comment: 'comment'
});
const reachableEndpoints = new Map();

class AssessmentApiClient {
	constructor(config, options = {}) {
		this.config = config;
		this.fetch = options.fetch || global.fetch;
		this.timeoutMs = options.timeoutMs || 15000;
	}

	async checkStatus() {
		await this.request('GET', 'me');
		return { reachable: true, apiUrl: this.getApiUrl(), fallback: this.getApiUrl() !== this.config.apiUrl };
	}

	getApiUrl() {
		const preference = reachableEndpoints.get(this.config.apiUrl);
		return preference && preference.until > Date.now() ? preference.url : this.config.apiUrl;
	}
	getFormUrl() { return this.getApiUrl().replace(/api$/, ''); }

	async createRound(round, maxResponses) {
		return this.request('PUT', `rounds/${round.externalRoundId}`, {
			title: round.title,
			criteria: round.criteria,
			deadline: round.deadline ? deadlineEndZurich(round.deadline) : null,
			maxResponses
		});
	}

	async closeRound(id) { return this.request('POST', `rounds/${id}/close`); }
	async deleteRound(id) {
		try { return await this.request('DELETE', `rounds/${id}`); }
		catch (error) { if (error.code === 'ASSESSMENT_ROUND_NOT_FOUND') return { deleted: true }; throw error; }
	}

	async fetchResponses({ externalRoundId, criteria, knownResponseIds = [] }) {
		const body = { responses: [] };
		let cursor = 0;
		for (let page = 0; page < 10; page++) {
			const result = await this.request('GET', `rounds/${externalRoundId}/responses?cursor=${cursor}`);
			if (!Array.isArray(result.responses)) throw apiError('Ungültige Antwortliste.', 'ASSESSMENT_API_SCHEMA_INVALID');
			body.responses.push(...result.responses);
			body.closed = result.closed;
			if (!result.nextCursor) break;
			if (!Number.isSafeInteger(result.nextCursor) || result.nextCursor <= cursor || page === 9) throw apiError('Ungültige Seitennavigation.', 'ASSESSMENT_API_SCHEMA_INVALID');
			cursor = result.nextCursor;
		}
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
		const known = new Set(knownResponseIds.map(String));
		return { responses: responses.filter(r => !known.has(r.externalResponseId)), invalidResponses, ignoredRoundCount, closed: body.closed };
	}

	async request(method, resource, payload) {
		const preferred = this.getApiUrl();
		try { return await this.performRequest(preferred, method, resource, payload); }
		catch (error) {
			const recoverable = ['ASSESSMENT_API_NETWORK_ERROR', 'ASSESSMENT_API_TIMEOUT', 'ASSESSMENT_API_INVALID_JSON', 'ASSESSMENT_API_UNAVAILABLE'].includes(error.code);
			const alternative = preferred === this.config.apiUrl ? this.config.fallbackApiUrl : this.config.apiUrl;
			if (!recoverable || !alternative || alternative === preferred) throw error;
			const result = await this.performRequest(alternative, method, resource, payload);
			reachableEndpoints.set(this.config.apiUrl, {url:alternative, until:Date.now()+10*60*1000});
			return result;
		}
	}

	async performRequest(apiUrl, method, resource, payload) {
		const controller = new AbortController();
		const timer = setTimeout(() => controller.abort(), this.timeoutMs);
		let response;
		try {
			const request = buildRequest({...this.config, apiUrl}, method, resource, payload);
			response = await this.fetch(request.url, {
				method: request.method,
				headers: request.headers,
				body: request.body,
				signal: controller.signal,
				redirect: 'error'
			});
		} catch (error) {
			if (error && error.name === 'AbortError') throw apiError('Zeitüberschreitung beim Abrufen der Mitmachnoten.', 'ASSESSMENT_API_TIMEOUT');
			throw apiError('Die Mitmachnoten-API ist derzeit nicht erreichbar.', 'ASSESSMENT_API_NETWORK_ERROR');
		} finally {
			clearTimeout(timer);
		}
		let body;
		try {
			body = await response.json();
		} catch (error) {
			throw apiError('Die Mitmachnoten-API hat keine gültige JSON-Antwort geliefert.', 'ASSESSMENT_API_INVALID_JSON');
		}
		if (response.status === 401 || response.status === 403) throw apiError('Der Lehrerzugang ist abgelaufen oder wurde widerrufen.', 'ASSESSMENT_API_UNAUTHORIZED');
		if (response.status === 404) throw apiError('Die Online-Runde wurde nicht gefunden oder bereits gelöscht.', 'ASSESSMENT_ROUND_NOT_FOUND');
		if (response.status >= 500) throw apiError('Die Mitmachnoten-API ist vorübergehend nicht verfügbar.', 'ASSESSMENT_API_UNAVAILABLE');
		if (!response.ok) throw apiError('Die Mitmachnoten-API konnte die Anfrage nicht verarbeiten.', 'ASSESSMENT_API_HTTP_ERROR');
		if (!body || body.success !== true) throw apiError(publicApiMessage(body), 'ASSESSMENT_API_REJECTED');
		return body;
	}
}

function buildRequest(config, method, resource, payload) {
	return {
		url: `${config.apiUrl.replace(/\/$/, '')}/teacher/${resource}`,
		method,
		headers: { 'Content-Type': 'application/json', Accept: 'application/json', Authorization: `Bearer ${config.teacherToken}` },
		body: payload === undefined ? undefined : JSON.stringify(payload)
	};
}

function deadlineEndZurich(date) {
	const midnight = Date.parse(`${date}T23:59:59Z`);
	const formatter = new Intl.DateTimeFormat('en-US', { timeZone: 'Europe/Zurich', timeZoneName: 'longOffset' });
	const offset = formatter.formatToParts(new Date(midnight)).find(p => p.type === 'timeZoneName').value.match(/GMT\+(\d+)/);
	return new Date(midnight - Number(offset[1]) * 3600000).toISOString();
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
	const additionalQuality = raw[RESPONSE_FIELDS.additionalQuality];
	if (additionalQuality !== undefined && additionalQuality !== null && typeof additionalQuality !== 'string') throw schemaError('Zusätzliche Qualität ist ungültig.');
	const comment = raw[RESPONSE_FIELDS.comment];
	if (comment !== undefined && comment !== null && typeof comment !== 'string') throw schemaError('Kommentar ist ungültig.');
	return { externalResponseId, timestamp, externalRoundId, name, email, answers, additionalQuality: String(additionalQuality || '').trim(), comment: String(comment || '').trim() };
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
