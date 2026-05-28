const fs = require('fs');
const path = require('path');

const REPORT_ENDPOINT = 'https://issues.akzuwo.ch/report';

function getReportApiKey(appRoot) {
	if (process.env.REPETIERER_REPORT_API_KEY) return process.env.REPETIERER_REPORT_API_KEY;
	if (process.env.REPORT_API_KEY) return process.env.REPORT_API_KEY;

	const configDirs = [appRoot, process.cwd()].filter(Boolean);
	for (const configDir of configDirs) {
		try {
			const configPath = path.join(configDir, 'report-config.local.json');
			if (!fs.existsSync(configPath)) continue;
			const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
			return typeof config.reportApiKey === 'string' ? config.reportApiKey : '';
		} catch (error) {
			return '';
		}
	}

	return '';
}

module.exports = {
	REPORT_ENDPOINT,
	getReportApiKey
};
