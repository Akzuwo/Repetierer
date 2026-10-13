const test = require('node:test');
const assert = require('node:assert/strict');
const { AssessmentApiClient, validateResponse } = require('../src/assessment/assessmentApiClient.js');

const config = { apiUrl: 'https://assessment.example.ch/api', teacherToken: 'rt_' + 'a'.repeat(43) };
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

test('ruft die API mit Bearer-Token auf und überträgt ihn nicht in URL oder Body', async () => {
	let call;
	const client = new AssessmentApiClient(config, { fetch: async (url, options) => {
		call = { url, options };
		return new Response(JSON.stringify({ success: true, responses: [validResponse()] }), { status: 200 });
	} });
	const result = await client.fetchResponses({ externalRoundId: 'assessment-3a-hs-abc123', criteria, knownResponseIds: ['old'] });
	assert.equal(call.options.method, 'GET');
	assert.equal(call.url.includes(config.teacherToken), false);
	assert.equal(call.options.headers.Authorization, `Bearer ${config.teacherToken}`);
	assert.equal(call.options.body, undefined);
	assert.equal(call.options.redirect, 'error');
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

test('fällt bei einer blockierten Hauptdomain auf die zweite Domain zurück und verwendet deren Einladungslink', async () => {
	const calls = [];
	const dual = {...config, apiUrl:'https://primary.example.ch/api', fallbackApiUrl:'https://secondary.example.ch/api'};
	const client = new AssessmentApiClient(dual, { fetch: async (url, options) => {
		calls.push({ url: String(url), options });
		if (String(url).startsWith(dual.apiUrl)) return new Response('<html>Network filter</html>', { status: 403 });
		return new Response(JSON.stringify({ success: true, roundId: 'assessment-3a-hs-abc123', count: 0, responses: [] }), { status: 200 });
	} });
	const result = await client.fetchResponses({ externalRoundId: 'assessment-3a-hs-abc123', criteria });
	assert.equal(result.responses.length, 0);
	assert.deepEqual(calls.map(call => call.options.method), ['GET', 'GET']);
	assert.equal(client.getFormUrl(), 'https://secondary.example.ch/');
	assert.equal(calls[1].url.startsWith(dual.fallbackApiUrl), true);
	assert.equal(calls[1].url.includes(config.teacherToken), false);
});

test('ein abgelehnter Lehrerzugang löst keinen Domain-Fallback aus', async () => {
	let calls = 0;
	const client = new AssessmentApiClient({...config,apiUrl:'https://auth.example.ch/api',fallbackApiUrl:'https://unused.example.ch/api'}, {fetch:async()=>{calls++;return new Response('{"success":false,"error":"unauthorized"}',{status:401});}});
	await assert.rejects(client.checkStatus(), e=>e.code==='ASSESSMENT_API_UNAUTHORIZED');
	assert.equal(calls,1);
});

test('liest mehrere Antwortseiten und filtert lokal bekannte IDs', async () => {
	let calls=0;
	const client=new AssessmentApiClient({...config,apiUrl:'https://pages.example.ch/api'}, {fetch:async()=>new Response(JSON.stringify(++calls===1 ? {success:true,responses:[validResponse()],nextCursor:100} : {success:true,responses:[validResponse({responseId:'response-2'})],nextCursor:null}))});
	const result=await client.fetchResponses({externalRoundId:'assessment-3a-hs-abc123',criteria,knownResponseIds:['response-1']});
	assert.equal(calls,2);assert.deepEqual(result.responses.map(r=>r.externalResponseId),['response-2']);
});

test('validiert Pflichtfelder und Bewertungswerte', () => {
	const response = validateResponse(validResponse(), 'assessment-3a-hs-abc123', criteria);
	assert.equal(response.answers.overall_participation, 4);
	assert.equal(response.additionalQuality, 'Spezifisches Fachwissen');
	assert.throws(() => validateResponse(validResponse({ answers: { participation_frequency: 6 } }), 'assessment-3a-hs-abc123', criteria), /Bewertungswert/);
});
