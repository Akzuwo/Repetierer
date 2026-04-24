const { app } = require('electron');
const fs = require('fs');
const path = require('path');

const defaultAppSettings = {
	extraJokerAfterThreeGrades: false,
	probabilityDecreaseFactor: 3,
	boostNeverSelected: false,
	neverSelectedBoostFactor: 3
};

function getPaths() {
	const storageDir = app.getPath('userData');
	return {
		storageDir: storageDir,
		settingsPath: path.join(storageDir, 'settings.json'),
		backupPath: path.join(storageDir, 'backup.json')
	};
}

function ensureStorageDir() {
	const storageDir = getPaths().storageDir;
	if (!fs.existsSync(storageDir)) {
		fs.mkdirSync(storageDir, { recursive: true });
	}
}

function readJson(filePath, fallback) {
	try {
		ensureStorageDir();
		if (!fs.existsSync(filePath)) return fallback;
		return JSON.parse(fs.readFileSync(filePath, 'utf8'));
	} catch (error) {
		return fallback;
	}
}

function writeJson(filePath, data) {
	ensureStorageDir();
	fs.writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf8');
}

function getSettings() {
	return readJson(getPaths().settingsPath, {});
}

function saveSettings(settings) {
	writeJson(getPaths().settingsPath, Object.assign({}, getSettings(), settings));
}

function getAppSettings() {
	const settings = getSettings();
	return Object.assign({}, defaultAppSettings, {
		extraJokerAfterThreeGrades: !!settings.extraJokerAfterThreeGrades,
		probabilityDecreaseFactor: normalizeFactor(settings.probabilityDecreaseFactor, defaultAppSettings.probabilityDecreaseFactor),
		boostNeverSelected: !!settings.boostNeverSelected,
		neverSelectedBoostFactor: normalizeFactor(settings.neverSelectedBoostFactor, defaultAppSettings.neverSelectedBoostFactor)
	});
}

function saveAppSettings(settings) {
	saveSettings({
		extraJokerAfterThreeGrades: !!settings.extraJokerAfterThreeGrades,
		probabilityDecreaseFactor: normalizeFactor(settings.probabilityDecreaseFactor, defaultAppSettings.probabilityDecreaseFactor),
		boostNeverSelected: !!settings.boostNeverSelected,
		neverSelectedBoostFactor: normalizeFactor(settings.neverSelectedBoostFactor, defaultAppSettings.neverSelectedBoostFactor)
	});
}

function normalizeFactor(value, fallback) {
	const factor = Number(value);
	if (!Number.isFinite(factor)) return fallback;
	return Math.max(1.1, Math.min(10, factor));
}

function getExcelFilePath() {
	return getSettings().excelFilePath;
}

function saveExcelFilePath(filePath) {
	saveSettings({ excelFilePath: filePath });
}

function getBackup() {
	const backup = readJson(getPaths().backupPath, { entries: [] });
	if (!backup.entries) backup.entries = [];
	return backup;
}

function addBackupEntry(entry) {
	const backup = getBackup();
	backup.entries.push(Object.assign({
		createdAt: new Date().toISOString()
	}, entry));
	writeJson(getPaths().backupPath, backup);
}

module.exports = {
	getBackup,
	getAppSettings,
	getExcelFilePath,
	saveExcelFilePath,
	saveAppSettings,
	addBackupEntry,
	getPaths
};
