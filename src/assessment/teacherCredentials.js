// Only the Electron main process may access the teacher token.
function protectTeacherToken(token, safeStorage) {
	if (!/^rt_[A-Za-z0-9_-]{43}$/.test(token || '')) throw new Error('Ungültiger Lehrerzugang. Importiere die ursprüngliche Zugangsdatei.');
	if (!safeStorage.isEncryptionAvailable() || safeStorage.getSelectedStorageBackend?.() === 'basic_text') throw new Error('Der Betriebssystem-Schlüsselspeicher ist nicht verfügbar.');
	return 'encrypted:' + safeStorage.encryptString(token).toString('base64');
}
function revealTeacherToken(token, safeStorage) {
	if (!token.startsWith('encrypted:')) throw new Error('Bitte den Lehrerzugang erneut importieren, damit er sicher gespeichert wird.');
	if (!safeStorage.isEncryptionAvailable()) throw new Error('Der Betriebssystem-Schlüsselspeicher ist nicht verfügbar.');
	return safeStorage.decryptString(Buffer.from(token.slice(10), 'base64'));
}
module.exports = { protectTeacherToken, revealTeacherToken };
