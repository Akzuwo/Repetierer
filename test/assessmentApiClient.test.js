const test = require('node:test');
const assert = require('node:assert/strict');
const { AssessmentApiClient, validateResponse } = require('../src/assessment/assessmentApiClient.js');

const config = { apiUrl: 'https://script.google.com/macros/s/test/exec', apiKey: 'super-secret-key' };
const criteria = [
	{ id: 'participation_frequency' },
	{ id: 'insufficient_contributions' },
	{ id: 'original_unsuitable_contributions' },
	{ id: 'correct_brief_contributions' },
	{ id: 'advancing_contributions' },
	{ id: 'independent_contributions' },
	{ id: 'overall_participation' }
];
const answers = value => Object.fromEntries(criteria.map(criterion => [criterion.id, value]));

function validResponse(overrides = {}) {
	return Object.assign({
		responseId: 'response-1',
		timestamp: '2026-09-04T10:00:00.000Z',
		roundId: 'assessment-3a-hs-abc123',
		name: 'Anna Müller',
		email: 'anna@example.ch',
		answers: answers(4),
		additionalQuality: 'Spezifisches Fachwissen',
		comment: 'Kommentar'
	}, overrides);
}

test('ruft die API per POST auf und überträgt den Schlüssel nicht in der URL', async () => {
	let call;
	const client = new AssessmentApiClient(config, { fetch: async (url, options) => {
		call = { url, options };
		return new Response(JSON.stringify({ success: true, responses: [validResponse()] }), { status: 200 });
	} });
	const result = await client.fetchResponses({ externalRoundId: 'assessment-3a-hs-abc123', criteria, knownResponseIds: ['old'] });
	const body = JSON.parse(call.options.body);
	assert.equal(call.options.method, 'POST');
	assert.equal(call.url.includes(config.apiKey), false);
	assert.equal(body.apiKey, config.apiKey);
	assert.deepEqual(body.knownResponseIds, ['old']);
	assert.equal(result.responses[0].externalResponseId, 'response-1');
});

test('filtert Runden exakt und markiert ungültige Antworten ohne sie zu übernehmen', async () => {
	const client = new AssessmentApiClient(config, { fetch: async () => new Response(JSON.stringify({ success: true, responses: [
		validResponse(),
		validResponse({ responseId: 'other', roundId: 'assessment-3a-hs-abc123-extra' }),
		validResponse({ responseId: 'bad', email: 'keine-mail' })
	] }), { status: 200 }) });
	const result = await client.fetchResponses({ externalRoundId: 'assessment-3a-hs-abc123', criteria });
	assert.equal(result.responses.length, 1);
	assert.equal(result.ignoredRoundCount, 1);
	assert.equal(result.invalidResponses.length, 1);
});

test('normalisiert Authentifizierungs-, Netzwerk- und JSON-Fehler', async () => {
	const unauthorized = new AssessmentApiClient(config, { fetch: async () => new Response('{}', { status: 403 }) });
	await assert.rejects(unauthorized.checkStatus(), error => error.code === 'ASSESSMENT_API_UNAUTHORIZED');
	const offline = new AssessmentApiClient(config, { fetch: async () => { throw new Error('offline'); } });
	await assert.rejects(offline.checkStatus(), error => error.code === 'ASSESSMENT_API_NETWORK_ERROR');
	const invalidJson = new AssessmentApiClient(config, { fetch: async () => new Response('<html>', { status: 200 }) });
	await assert.rejects(invalidJson.checkStatus(), error => error.code === 'ASSESSMENT_API_INVALID_JSON');
});

test('fällt bei einer bestehenden doGet-API kontrolliert auf GET zurück', async () => {
	const calls = [];
	const client = new AssessmentApiClient(config, { fetch: async (url, options) => {
		calls.push({ url: String(url), options });
		if (options.method === 'POST') return new Response('<title>Fehler</title>Script-Funktion nicht gefunden: doPost', { status: 200 });
		return new Response(JSON.stringify({ success: true, roundId: 'assessment-3a-hs-abc123', count: 0, responses: [] }), { status: 200 });
	} });
	const result = await client.fetchResponses({ externalRoundId: 'assessment-3a-hs-abc123', criteria });
	assert.equal(result.responses.length, 0);
	assert.deepEqual(calls.map(call => call.options.method), ['POST', 'GET']);
	const fallbackUrl = new URL(calls[1].url);
	assert.equal(fallbackUrl.searchParams.get('apiKey'), config.apiKey);
	assert.equal(fallbackUrl.searchParams.get('roundId'), 'assessment-3a-hs-abc123');
});

test('validiert Pflichtfelder und Bewertungswerte', () => {
	const response = validateResponse(validResponse(), 'assessment-3a-hs-abc123', criteria);
	assert.equal(response.answers.overall_participation, 4);
	assert.equal(response.additionalQuality, 'Spezifisches Fachwissen');
	assert.throws(() => validateResponse(validResponse({ answers: { participation_frequency: 6 } }), 'assessment-3a-hs-abc123', criteria), /Bewertungswert/);
});
