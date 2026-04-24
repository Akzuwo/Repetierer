const {BrowserWindow, app, ipcMain} = require('electron');
const {autoUpdater} = require("electron-updater");
const fs = require('fs');
const path = require('path');
require('./src/mainProcess.js');

let win;
let updateDownloadApproved = false;
let updateCheckStarted = false;
let updateListenersRegistered = false;
let availableUpdateInfo;

// initialize the window
function initWindow() {
	win = new BrowserWindow({
		width: 1100,
		height: 800,
		resizable: false,
		transparent: true,
		frame: false,
		center: true,
		webPreferences: {
			nodeIntegration: true,
			contextIsolation: false
		},
	});
}

// load the index.html file and display it
function loadHTML() {
	win.loadFile('public/index.html')
}

// close the window
function closeWindow() {
	win = null;
}

function createApp() {
	initWindow();
	loadHTML();
	win.webContents.on('did-finish-load', () => {
		notifyInstalledUpdateIfNeeded();
		checkForUpdatesWithConsent();
	});
	win.on('close', () => {
		closeWindow();
	})
}

function checkForUpdatesWithConsent() {
	if (updateCheckStarted) return;
	updateCheckStarted = true;

	if (!app.isPackaged) {
		console.log('Update check skipped in development mode.');
		return;
	}

	autoUpdater.autoDownload = false;
	autoUpdater.autoInstallOnAppQuit = false;
	registerUpdaterListeners();

	sendUpdateStatus('checking');
	autoUpdater.checkForUpdates().catch(error => {
		console.error('Update check failed:', error);
		sendUpdateStatus('error', {
			message: 'Update-Prüfung fehlgeschlagen.',
			error: getErrorMessage(error)
		});
	});
}

function registerUpdaterListeners() {
	if (updateListenersRegistered) return;
	updateListenersRegistered = true;

	autoUpdater.on('error', error => {
		console.error('Update error:', error);
		sendUpdateStatus('error', {
			message: 'Beim Update ist ein Fehler aufgetreten.',
			error: getErrorMessage(error)
		});
	});

	autoUpdater.on('update-available', async info => {
		console.log('Update available:', info.version);
		availableUpdateInfo = info;
		sendUpdateStatus('available', {
			version: info.version
		});
	});

	autoUpdater.on('update-not-available', info => {
		console.log('No update available:', info.version);
		sendUpdateStatus('not-available', {
			version: info.version
		});
	});

	autoUpdater.on('download-progress', progress => {
		const percent = Number(progress.percent || 0).toFixed(1);
		console.log(`Update download progress: ${percent}%`);
		sendUpdateStatus('progress', {
			percent: Number(percent)
		});
	});

	autoUpdater.on('update-downloaded', async info => {
		if (!updateDownloadApproved) {
			console.log('Update downloaded event ignored because no user approval was recorded.');
			return;
		}

		console.log('Update downloaded:', info.version);
		sendUpdateStatus('downloaded', {
			version: info.version
		});

		setTimeout(() => installDownloadedUpdate(info), 700);
	});
}

ipcMain.on('update-download-approved', async event => {
	if (!app.isPackaged) {
		sendUpdateStatus('error', {
			message: 'Updates sind im Entwicklungsmodus deaktiviert.'
		});
		return;
	}

	if (!availableUpdateInfo) {
		sendUpdateStatus('error', {
			message: 'Es ist kein Update zum Herunterladen verfügbar.'
		});
		return;
	}

	updateDownloadApproved = true;
	sendUpdateStatus('downloading', {
		version: availableUpdateInfo.version,
		percent: 0
	});

	try {
		await autoUpdater.downloadUpdate();
	} catch (error) {
		console.error('Update download failed:', error);
		sendUpdateStatus('error', {
			message: 'Update-Download fehlgeschlagen.',
			error: getErrorMessage(error)
		});
	}
});

function installDownloadedUpdate(info) {
	sendUpdateStatus('installing', {
		version: info.version
	});

	try {
		writeUpdateState({
			pendingInstalledVersion: info.version
		});
		autoUpdater.quitAndInstall(true, true);
	} catch (error) {
		console.error('Silent update install failed:', error);
		writeUpdateState({});
		sendUpdateStatus('error', {
			message: 'Das Update konnte nicht automatisch installiert werden. Bitte installiere es über den heruntergeladenen Installer.',
			error: getErrorMessage(error)
		});
	}
}

function notifyInstalledUpdateIfNeeded() {
	const updateState = readUpdateState();
	if (updateState.pendingInstalledVersion && updateState.pendingInstalledVersion === app.getVersion()) {
		sendUpdateStatus('installed', {
			version: app.getVersion(),
			message: 'Update erfolgreich installiert'
		});
		writeUpdateState({});
	}
}

function sendUpdateStatus(status, data) {
	const payload = Object.assign({ status: status }, data || {});
	if (win && !win.isDestroyed() && win.webContents) {
		win.webContents.send('update-status', payload);
	}
}

function getUpdateStatePath() {
	return path.join(app.getPath('userData'), 'update-state.json');
}

function readUpdateState() {
	try {
		const updateStatePath = getUpdateStatePath();
		if (!fs.existsSync(updateStatePath)) return {};
		return JSON.parse(fs.readFileSync(updateStatePath, 'utf8'));
	} catch (error) {
		console.error('Could not read update state:', error);
		return {};
	}
}

function writeUpdateState(updateState) {
	try {
		const updateStatePath = getUpdateStatePath();
		fs.mkdirSync(path.dirname(updateStatePath), { recursive: true });
		fs.writeFileSync(updateStatePath, JSON.stringify(updateState || {}, null, 2), 'utf8');
	} catch (error) {
		console.error('Could not write update state:', error);
	}
}

function getErrorMessage(error) {
	if (!error) return '';
	return error.message || String(error);
}

// main method
app.whenReady().then(createApp);

app.on('window-all-closed', () => {
	if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
	if (BrowserWindow.getAllWindows().length === 0) createApp();
});
