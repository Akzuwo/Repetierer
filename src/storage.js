const { app } = require('electron');
const fs = require('fs');
const path = require('path');

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
	getExcelFilePath,
	saveExcelFilePath,
	addBackupEntry,
	getPaths
};
