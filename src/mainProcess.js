const { ipcMain, app, BrowserWindow, dialog } = require('electron');
const fs = require('fs');
const { setFile, setClass, selectPerson, saveGrade, setJoker, selectSpecificPerson, getPersons, getProbabilities, applyPendingExcelEntries} = require('./program.js');
const { getBackup, getAppSettings, getPendingExcelEntries, getExcelFilePath, saveExcelFilePath, saveAppSettings, addBackupEntry, addPendingExcelEntry, removePendingExcelEntries, getPaths } = require('./storage.js');

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
		setFile(file[0], result => {
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
	event.sender.send('backup-data', getBackup(), getPaths());
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
		}

		const remaining = getPendingExcelEntries().length;
		event.sender.send('pending-excel-status', remaining);
		event.sender.send('pending-excel-flushed', Object.assign({}, result, { remaining: remaining }));
	});
});

ipcMain.on('discard-pending-excel-entry', (event, pendingEntryId) => {
	removePendingExcelEntries([pendingEntryId]);
	event.sender.send('pending-excel-status', getPendingExcelEntries().length);
});

function handleExcelBackupEntry(event, backupEntry) {
	if (!backupEntry) return;
	addBackupEntry(backupEntry);

	if (!backupEntry.excelWriteSucceeded && backupEntry.excelWriteReason === 'excel-write-failed') {
		const pendingEntry = addPendingExcelEntry(backupEntry);
		event.sender.send('excel-write-pending', pendingEntry);
	}

	event.sender.send('pending-excel-status', getPendingExcelEntries().length);
}
