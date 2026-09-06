const fs = require('fs');
const path = require('path');
const { app, net } = require('electron');
const { AssessmentApiClient } = require('../src/assessment/assessmentApiClient.js');
const { discoverRoundFieldId } = require('../src/assessment/assessmentForm.js');
const { buildRoundFormUrl, readAssessmentConfig } = require('../src/assessment/integrationConfig.js');

async function run() {
	const sourceUserData = process.env.ASSESSMENT_SMOKE_USER_DATA || app.getPath('userData');
	const config = readAssessmentConfig(path.join(sourceUserData, 'repetierer.env'), path.join(sourceUserData, 'brevo-smtp.env'));
	if (!config) throw new Error('Mitmachnoten-Konfiguration fehlt.');
	const roundFieldId = config.roundFieldId || await discoverRoundFieldId(config.formUrl, { fetch: net.fetch });
	const resolvedConfig = Object.assign({}, config, { roundFieldId });
	const assessmentPath = path.join(sourceUserData, 'mitmachnoten.json');
	const data = fs.existsSync(assessmentPath) ? JSON.parse(fs.readFileSync(assessmentPath, 'utf8')) : { workbooks: {} };
	const rounds = [];
	for (const workbook of Object.values(data.workbooks || {})) {
		for (const classRecord of Object.values(workbook.classes || {})) rounds.push(...(classRecord.rounds || []));
	}
	const client = new AssessmentApiClient(resolvedConfig, { fetch: net.fetch, timeoutMs: 20000 });
	await client.checkStatus();
	const results = [];
	for (const round of rounds) {
		const result = await client.fetchResponses({ externalRoundId: round.externalRoundId, criteria: round.criteria, knownResponseIds: [] });
		results.push({ roundId: round.externalRoundId, valid: result.responses.length, invalid: result.invalidResponses.length, ignored: result.ignoredRoundCount });
	}
	const sampleRoundId = rounds[0] && rounds[0].externalRoundId || 'assessment-smoke-test';
	const invitationUrl = new URL(buildRoundFormUrl(resolvedConfig, sampleRoundId));
	console.log(JSON.stringify({
		ok: true,
		apiReachable: true,
		roundFieldId,
		invitationRoundId: invitationUrl.searchParams.get(roundFieldId),
		roundResults: results
	}));
}

app.whenReady().then(run).then(() => app.quit()).catch(error => {
	console.error(JSON.stringify({ ok: false, code: error.code || 'SMOKE_FAILED', message: error.message }));
	app.exit(1);
});
