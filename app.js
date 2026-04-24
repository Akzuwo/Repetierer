const {BrowserWindow, app, dialog} = require('electron');
const {autoUpdater} = require("electron-updater");
require('./src/mainProcess.js');

let win;
let updateDownloadApproved = false;
let updateCheckStarted = false;

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
	win.on('close', () => {
		closeWindow();
	})
	checkForUpdatesWithConsent();
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

	autoUpdater.on('error', error => {
		console.error('Update error:', error);
	});

	autoUpdater.on('update-available', async info => {
		console.log('Update available:', info.version);
		const result = await dialog.showMessageBox(win, {
			type: 'info',
			title: 'Neue Version verfügbar',
			message: 'Neue Version verfügbar',
			detail: `Version ${info.version} ist verfügbar. Möchtest du sie herunterladen und installieren?`,
			buttons: ['Installieren', 'Später'],
			defaultId: 0,
			cancelId: 1
		});

		if (result.response !== 0) {
			console.log('Update download declined by user.');
			return;
		}

		updateDownloadApproved = true;
		try {
			await autoUpdater.downloadUpdate();
		} catch (error) {
			console.error('Update download failed:', error);
		}
	});

	autoUpdater.on('update-not-available', info => {
		console.log('No update available:', info.version);
	});

	autoUpdater.on('download-progress', progress => {
		const percent = Number(progress.percent || 0).toFixed(1);
		console.log(`Update download progress: ${percent}%`);
	});

	autoUpdater.on('update-downloaded', async info => {
		if (!updateDownloadApproved) {
			console.log('Update downloaded event ignored because no user approval was recorded.');
			return;
		}

		console.log('Update downloaded:', info.version);
		const result = await dialog.showMessageBox(win, {
			type: 'info',
			title: 'Update bereit',
			message: 'Update bereit',
			detail: 'Das Update wurde heruntergeladen. Soll Repetierer jetzt neu gestartet und aktualisiert werden?',
			buttons: ['Jetzt neu starten', 'Später'],
			defaultId: 0,
			cancelId: 1
		});

		if (result.response === 0) {
			autoUpdater.quitAndInstall();
		} else {
			console.log('Update installation postponed by user.');
		}
	});

	autoUpdater.checkForUpdates().catch(error => {
		console.error('Update check failed:', error);
	});
}

// main method
app.whenReady().then(createApp);

app.on('window-all-closed', () => {
	if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
	if (BrowserWindow.getAllWindows().length === 0) createApp();
});
