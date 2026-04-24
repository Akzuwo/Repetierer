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
		backupPath: path.join(storageDir, 'backup.json'),
		logPath: path.join(storageDir, 'repetierer.log')
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
	if (!backup.pendingExcelEntries) backup.pendingExcelEntries = [];
	return backup;
}

function addBackupEntry(entry) {
	const backup = getBackup();
	backup.entries.push(Object.assign({
		createdAt: new Date().toISOString()
	}, entry));
	writeJson(getPaths().backupPath, backup);
}

function addPendingExcelEntry(entry) {
	const backup = getBackup();
	const pendingEntry = Object.assign({
		id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
		createdAt: new Date().toISOString()
	}, entry);
	backup.pendingExcelEntries.push(pendingEntry);
	writeJson(getPaths().backupPath, backup);
	return pendingEntry;
}

function getPendingExcelEntries() {
	return getBackup().pendingExcelEntries;
}

function removePendingExcelEntries(ids) {
	const idSet = new Set(ids || []);
	const backup = getBackup();
	backup.pendingExcelEntries = backup.pendingExcelEntries.filter(entry => !idSet.has(entry.id));
	writeJson(getPaths().backupPath, backup);
}

function logEvent(message, details) {
	try {
		ensureStorageDir();
		const payload = details ? ` ${JSON.stringify(details)}` : '';
		const line = `[${new Date().toISOString()}] ${message}${payload}\n`;
		fs.appendFileSync(getPaths().logPath, line, 'utf8');
	} catch (error) {
		console.error('Logging fehlgeschlagen:', error);
	}
}

function getLogs() {
	try {
		ensureStorageDir();
		const logPath = getPaths().logPath;
		if (!fs.existsSync(logPath)) return '';
		return fs.readFileSync(logPath, 'utf8');
	} catch (error) {
		return `Logs konnten nicht gelesen werden: ${error && error.message ? error.message : String(error)}`;
	}
}

module.exports = {
	getBackup,
	getAppSettings,
	getPendingExcelEntries,
	getExcelFilePath,
	saveExcelFilePath,
	saveAppSettings,
	addBackupEntry,
	addPendingExcelEntry,
	removePendingExcelEntries,
	logEvent,
	getLogs,
	getPaths
};
