const { app } = require('electron');
const fs = require('fs');
const path = require('path');

const defaultAppSettings = {
	extraJokerAfterThreeGrades: false,
	probabilityDecreaseFactor: 3,
	boostNeverSelected: false,
	neverSelectedBoostFactor: 3,
	logoAnimationEnabled: false,
	wiggersRuleEnabled: true,
	wiggersRuleDurationMinutes: 120
};
const MAX_BACKUP_PREVIEW_ENTRIES = 100;
const MAX_LOG_READ_BYTES = 200 * 1024;

function getPaths() {
	const storageDir = app.getPath('userData');
	return {
		storageDir: storageDir,
		settingsPath: path.join(storageDir, 'settings.json'),
		backupPath: path.join(storageDir, 'backup.json'),
		migrationsPath: path.join(storageDir, 'migrations.json'),
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

function getWiggersRulePenalties() {
	return getSettings().wiggersRulePenalties || {};
}

function saveWiggersRulePenalty(key, expiresAt) {
	const penalties = getWiggersRulePenalties();
	penalties[key] = expiresAt;
	saveSettings({ wiggersRulePenalties: penalties });
}

function removeWiggersRulePenalty(key) {
	const penalties = getWiggersRulePenalties();
	delete penalties[key];
	saveSettings({ wiggersRulePenalties: penalties });
}

function clearWiggersRulePenalties() {
	saveSettings({ wiggersRulePenalties: {} });
}

function getMigrations() {
	const migrations = readJson(getPaths().migrationsPath, { files: {} });
	if (!migrations.files) migrations.files = {};
	return migrations;
}

function getFileMigrationStatus(filePath, migrationId) {
	const fileId = getFileIdentifier(filePath);
	if (!fileId || !migrationId) {
		return {
			hasFile: false,
			fileId: '',
			completed: false
		};
	}

	const migrations = getMigrations();
	const fileMigrations = migrations.files[fileId] || {};
	return {
		hasFile: true,
		fileId: fileId,
		completed: !!fileMigrations[migrationId],
		completedAt: fileMigrations[migrationId] && fileMigrations[migrationId].completedAt
	};
}

function markFileMigrationCompleted(filePath, migrationId, details) {
	const fileId = getFileIdentifier(filePath);
	if (!fileId || !migrationId) return false;

	const migrations = getMigrations();
	if (!migrations.files[fileId]) migrations.files[fileId] = {};
	migrations.files[fileId][migrationId] = Object.assign({}, details || {}, {
		filePath: path.resolve(filePath),
		completedAt: new Date().toISOString()
	});
	writeJson(getPaths().migrationsPath, migrations);
	return true;
}

function getFileIdentifier(filePath) {
	if (!filePath) return '';
	let resolvedPath;
	try {
		resolvedPath = fs.realpathSync.native ? fs.realpathSync.native(filePath) : fs.realpathSync(filePath);
	} catch (error) {
		resolvedPath = path.resolve(filePath);
	}
	return process.platform === 'win32' ? resolvedPath.toLowerCase() : resolvedPath;
}

function getAppSettings() {
	const settings = getSettings();
	return Object.assign({}, defaultAppSettings, {
		extraJokerAfterThreeGrades: !!settings.extraJokerAfterThreeGrades,
		probabilityDecreaseFactor: normalizeFactor(settings.probabilityDecreaseFactor, defaultAppSettings.probabilityDecreaseFactor),
		boostNeverSelected: !!settings.boostNeverSelected,
		neverSelectedBoostFactor: normalizeFactor(settings.neverSelectedBoostFactor, defaultAppSettings.neverSelectedBoostFactor),
		logoAnimationEnabled: !!settings.logoAnimationEnabled,
		wiggersRuleEnabled: settings.wiggersRuleEnabled !== false,
		wiggersRuleDurationMinutes: normalizeDuration(settings.wiggersRuleDurationMinutes, defaultAppSettings.wiggersRuleDurationMinutes)
	});
}

function saveAppSettings(settings) {
	const wiggersRuleEnabled = settings.wiggersRuleEnabled !== false;
	saveSettings({
		extraJokerAfterThreeGrades: !!settings.extraJokerAfterThreeGrades,
		probabilityDecreaseFactor: normalizeFactor(settings.probabilityDecreaseFactor, defaultAppSettings.probabilityDecreaseFactor),
		boostNeverSelected: !!settings.boostNeverSelected,
		neverSelectedBoostFactor: normalizeFactor(settings.neverSelectedBoostFactor, defaultAppSettings.neverSelectedBoostFactor),
		logoAnimationEnabled: !!settings.logoAnimationEnabled,
		wiggersRuleEnabled: wiggersRuleEnabled,
		wiggersRuleDurationMinutes: normalizeDuration(settings.wiggersRuleDurationMinutes, defaultAppSettings.wiggersRuleDurationMinutes),
		...(wiggersRuleEnabled ? {} : { wiggersRulePenalties: {} })
	});
}

function normalizeFactor(value, fallback) {
	const factor = Number(value);
	if (!Number.isFinite(factor)) return fallback;
	return Math.max(1.1, Math.min(10, factor));
}

function normalizeDuration(value, fallback) {
	const duration = Number(value);
	if (!Number.isFinite(duration)) return fallback;
	return Math.max(1, Math.min(480, Math.round(duration)));
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

function getBackupPreview(limit) {
	const backup = getBackup();
	const entryLimit = Number.isInteger(limit) && limit > 0 ? limit : MAX_BACKUP_PREVIEW_ENTRIES;
	return {
		entries: backup.entries.slice(-entryLimit),
		pendingExcelEntries: backup.pendingExcelEntries,
		totalEntries: backup.entries.length
	};
}

function addBackupEntry(entry) {
	const backup = getBackup();
	const backupEntry = Object.assign({}, entry, {
		id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
		createdAt: new Date().toISOString()
	});
	backup.entries.push(backupEntry);
	writeJson(getPaths().backupPath, backup);
	return backupEntry;
}

function addPendingExcelEntry(entry) {
	const backup = getBackup();
	const pendingEntry = Object.assign({}, entry, {
		id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
		backupEntryId: entry && entry.id,
		createdAt: new Date().toISOString()
	});
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

function removeBackupEntries(ids) {
	const idSet = new Set(ids || []);
	const backup = getBackup();
	backup.entries = backup.entries.filter(entry => !idSet.has(entry.id));
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
		const stats = fs.statSync(logPath);
		if (stats.size <= MAX_LOG_READ_BYTES) return fs.readFileSync(logPath, 'utf8');

		const fd = fs.openSync(logPath, 'r');
		try {
			const buffer = Buffer.alloc(MAX_LOG_READ_BYTES);
			fs.readSync(fd, buffer, 0, MAX_LOG_READ_BYTES, stats.size - MAX_LOG_READ_BYTES);
			const text = buffer.toString('utf8');
			const firstNewline = text.indexOf('\n');
			const trimmedText = firstNewline >= 0 ? text.slice(firstNewline + 1) : text;
			return `[Log gekuerzt: zeige die letzten ${Math.round(MAX_LOG_READ_BYTES / 1024)} KB]\n${trimmedText}`;
		} finally {
			fs.closeSync(fd);
		}
	} catch (error) {
		return `Logs konnten nicht gelesen werden: ${error && error.message ? error.message : String(error)}`;
	}
}

module.exports = {
	getBackup,
	getBackupPreview,
	getAppSettings,
	getFileMigrationStatus,
	getWiggersRulePenalties,
	getPendingExcelEntries,
	getExcelFilePath,
	saveExcelFilePath,
	saveAppSettings,
	markFileMigrationCompleted,
	saveWiggersRulePenalty,
	removeWiggersRulePenalty,
	clearWiggersRulePenalties,
	addBackupEntry,
	addPendingExcelEntry,
	removePendingExcelEntries,
	removeBackupEntries,
	logEvent,
	getLogs,
	getPaths
};
