const fs = require('fs');
const path = require('path');

const ANALYTICS_ENDPOINT = 'https://analytics.akzuwo.ch';
const ANALYTICS_APP_ID = 'Repetierer';

function getAnalyticsApiKey(appRoot) {
	if (process.env.REPETIERER_ANALYTICS_KEY) return process.env.REPETIERER_ANALYTICS_KEY;
	if (process.env.ANALYTICS_KEY) return process.env.ANALYTICS_KEY;

	const configDirs = [appRoot, process.cwd()].filter(Boolean);
	for (const configDir of configDirs) {
		const key = readKeyFromConfig(path.join(configDir, 'analytics-config.local.json'), 'analyticsKey');
		if (key) return key;

		const legacyKey = readKeyFromConfig(path.join(configDir, 'report-config.local.json'), 'reportApiKey');
		if (legacyKey) return legacyKey;
	}

	return '';
}

function readKeyFromConfig(configPath, keyName) {
	try {
		if (!fs.existsSync(configPath)) return '';
		const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
		return typeof config[keyName] === 'string' ? config[keyName] : '';
	} catch (error) {
		return '';
	}
}

module.exports = {
	ANALYTICS_ENDPOINT,
	ANALYTICS_APP_ID,
	getAnalyticsApiKey
};
