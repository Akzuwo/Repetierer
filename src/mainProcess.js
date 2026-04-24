const { ipcMain, app, BrowserWindow, dialog } = require('electron');
const fs = require('fs');
const { setFile, setClass, selectPerson, saveGrade, setJoker, selectSpecificPerson, getPersons, getProbabilities, applyPendingExcelEntries, reloadExcel} = require('./program.js');
const { getBackup, getAppSettings, getPendingExcelEntries, getExcelFilePath, saveExcelFilePath, saveAppSettings, addBackupEntry, addPendingExcelEntry, removePendingExcelEntries, logEvent, getLogs, getPaths } = require('./storage.js');

logEvent('App gestartet', { version: app.getVersion() });

/*
 * events
 */

// minimize
ipcMain.on('minimize', (event, args) => {
	BrowserWindow.getFocusedWindow().minimize();
});

// quit
ipcMain.on('quit', (event, args) => {
	app.quit();
});

// app version
ipcMain.on('get-version', (event, args) => {
	event.sender.send('version', app.getVersion());
});

// settings
ipcMain.on('get-settings', (event, args) => {
	event.sender.send('settings-data', getAppSettings(), getPaths());
});

ipcMain.on('save-settings', (event, settings) => {
	saveAppSettings(settings || {});
	logEvent('Einstellungen gespeichert');
	event.sender.send('settings-saved', getAppSettings());
});

// file
ipcMain.on('file', (event, args) => {
	let file = dialog.showOpenDialogSync(BrowserWindow.getFocusedWindow(), {
		properties: ['openFile'],
		filters: [{ name: 'Excel Dateien (*.xlsx)', extensions: ['xlsx'] }]
	});
	if (file) {
		saveExcelFilePath(file[0]);
		logEvent('Datei ausgewählt', { filePath: file[0] });
		setFile(file[0], result => {
			logEvent(result ? 'Excel eingelesen' : 'Fehler beim Lesen', { filePath: file[0] });
			event.sender.send('classes', result, file[0]);
		});
	}
});

// load the last selected file automatically
ipcMain.on('load-saved-file', (event, args) => {
	const filePath = getExcelFilePath();
	if (!filePath || !fs.existsSync(filePath)) {
		event.sender.send('saved-file-missing', filePath);
		return;
	}

	setFile(filePath, result => {
		logEvent(result ? 'Excel eingelesen' : 'Fehler beim Lesen', { filePath: filePath, automatic: true });
		event.sender.send('classes', result, filePath);
		event.sender.send('pending-excel-status', getPendingExcelEntries().length);
	});
});

// class
ipcMain.on('class', (event, args) => {
	setClass(args, result => {
		event.sender.send('ready', result);
	});
});

// start
ipcMain.on('start', (event, args) => {
	const selectedPerson = selectPerson();
	if (selectedPerson) {
		event.sender.send('name', selectedPerson);
	} else {
		event.sender.send('no-person-available');
	}
});


// get persons list for manual selection
ipcMain.on('get-persons', (event, args) => {
	const persons = getPersons();
	event.sender.send('persons-list', persons);
});

// get probability list
ipcMain.on('get-probabilities', (event, args) => {
	event.sender.send('probability-data', getProbabilities());
});

// select specific person
ipcMain.on('select-person', (event, personId) => {
	const result = selectSpecificPerson(personId);
	if (result) {
		event.sender.send('name', result);
	}
});

// ok
ipcMain.on('ok', (event, args) => {
	saveGrade(args, (result, backupEntry) => {
		handleExcelBackupEntry(event, backupEntry);
		event.sender.send('finished', result);
	});
});

// joker
ipcMain.on('joker', (event, args) => {
    setJoker((result, backupEntry) => {
		handleExcelBackupEntry(event, backupEntry);
		event.sender.send('finished', result);
	});
});

// backup
ipcMain.on('get-backup', (event, args) => {
	logEvent('Backup geöffnet');
	event.sender.send('backup-data', getBackup(), getPaths());
});

ipcMain.on('get-logs', (event, args) => {
	logEvent('Log geöffnet');
	event.sender.send('log-data', getLogs(), getPaths());
});

// pending excel entries
ipcMain.on('get-pending-excel-status', (event, args) => {
	event.sender.send('pending-excel-status', getPendingExcelEntries().length);
});

ipcMain.on('flush-pending-excel', (event, args) => {
	const pendingEntries = getPendingExcelEntries();
	if (pendingEntries.length === 0) {
		event.sender.send('pending-excel-status', 0);
		return;
	}

	applyPendingExcelEntries(pendingEntries, result => {
		if (result && result.success) {
			removePendingExcelEntries(result.applied || []);
			logEvent('Excel aktualisiert', { applied: result.applied || [] });
		} else {
			logEvent(result && result.reason === 'excel-locked' ? 'Konflikt erkannt' : 'Fehler beim Schreiben', result || {});
		}

		const remaining = getPendingExcelEntries().length;
		event.sender.send('pending-excel-status', remaining);
		event.sender.send('pending-excel-flushed', Object.assign({}, result, { remaining: remaining }));
	});
});

ipcMain.on('reload-excel', (event, args) => {
	const pendingEntries = getPendingExcelEntries();
	if (pendingEntries.length > 0) {
		logEvent('Konflikt erkannt', { action: 'reload-excel', pendingEntries: pendingEntries.length });
		event.sender.send('reload-excel-conflict', pendingEntries.length);
		return;
	}

	reloadExcelFile(event);
});

ipcMain.on('reload-excel-force', (event, args) => {
	logEvent('Excel neu laden trotz lokaler Änderungen bestätigt');
	reloadExcelFile(event);
});

ipcMain.on('discard-pending-excel-entry', (event, pendingEntryId) => {
	removePendingExcelEntries([pendingEntryId]);
	event.sender.send('pending-excel-status', getPendingExcelEntries().length);
});

function handleExcelBackupEntry(event, backupEntry) {
	if (!backupEntry) return;
	addBackupEntry(backupEntry);
	logEvent('Backup erstellt', {
		type: backupEntry.type,
		personName: backupEntry.personName,
		excelWriteSucceeded: backupEntry.excelWriteSucceeded,
		reason: backupEntry.excelWriteReason
	});

	if (!backupEntry.excelWriteSucceeded && ['excel-write-failed', 'excel-locked'].includes(backupEntry.excelWriteReason)) {
		const pendingEntry = addPendingExcelEntry(backupEntry);
		logEvent(backupEntry.excelWriteReason === 'excel-locked' ? 'Konflikt erkannt' : 'Fehler beim Schreiben', {
			type: backupEntry.type,
			personName: backupEntry.personName,
			reason: backupEntry.excelWriteReason
		});
		event.sender.send('excel-write-pending', pendingEntry);
	}

	event.sender.send('pending-excel-status', getPendingExcelEntries().length);
}

function reloadExcelFile(event) {
	reloadExcel(result => {
		if (result && result.success) {
			logEvent('Excel neu geladen', { filePath: result.filePath, className: result.className });
			event.sender.send('classes', result.worksheets, result.filePath);
			event.sender.send('excel-reloaded', result);
			event.sender.send('pending-excel-status', getPendingExcelEntries().length);
		} else {
			logEvent('Fehler beim Lesen', result || {});
			event.sender.send('excel-reload-failed', result || { success: false, reason: 'unknown' });
		}
	});
}
