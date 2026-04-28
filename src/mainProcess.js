const { ipcMain, app, BrowserWindow, dialog } = require('electron');
const fs = require('fs');
const path = require('path');
const { getBackup, getAppSettings, getPendingExcelEntries, getExcelFilePath, saveExcelFilePath, saveAppSettings, addBackupEntry, addPendingExcelEntry, removePendingExcelEntries, removeBackupEntries, logEvent, getLogs, getPaths } = require('./storage.js');
const isDebugMode = process.argv.includes('--dev-mode');
const sessionEntries = [];
let lastUndoEntry;
let lastRedoEntry;
let programModule;
let ExcelJSModule;

function getProgram() {
	if (!programModule) programModule = require('./program.js');
	return programModule;
}

function createWorkbook() {
	if (!ExcelJSModule) ExcelJSModule = require('exceljs');
	return new ExcelJSModule.Workbook();
}

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

ipcMain.on('get-debug-mode', (event, args) => {
	event.sender.send('debug-mode', isDebugMode);
});

// settings
ipcMain.on('get-settings', (event, args) => {
	event.sender.send('settings-data', getAppSettings(), getPaths());
});

ipcMain.on('get-ui-settings', (event, args) => {
	event.sender.send('ui-settings', getAppSettings());
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
		getProgram().setFile(file[0], result => {
			logEvent(result ? 'Excel eingelesen' : 'Fehler beim Lesen', { filePath: file[0] });
			sendClassesOrJokerMigration(event, result, file[0]);
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

	getProgram().setFile(filePath, result => {
		logEvent(result ? 'Excel eingelesen' : 'Fehler beim Lesen', { filePath: filePath, automatic: true });
		sendClassesOrJokerMigration(event, result, filePath);
		event.sender.send('pending-excel-status', getPendingExcelEntries().length);
	});
});

// class
ipcMain.on('class', (event, args) => {
	getProgram().setClass(args, result => {
		event.sender.send('ready', result);
	});
});

// start
ipcMain.on('start', (event, args) => {
	const selectedPerson = getProgram().selectPerson();
	if (selectedPerson) {
		event.sender.send('name', selectedPerson);
	} else {
		event.sender.send('no-person-available');
	}
});


// get persons list for manual selection
ipcMain.on('get-persons', (event, args) => {
	const persons = getProgram().getPersons();
	event.sender.send('persons-list', persons);
});

ipcMain.on('get-editor-persons', (event, args) => {
	const persons = getProgram().getEditorPersons();
	event.sender.send('editor-persons-data', persons);
});

ipcMain.on('get-absence-persons', (event, args) => {
	event.sender.send('absence-persons-data', getProgram().getAbsencePersons());
});

ipcMain.on('save-absences', (event, absentIds) => {
	getProgram().setAbsences(absentIds, result => {
		if (result && result.success) {
			logEvent('Abwesenheiten gespeichert', { count: result.count });
			event.sender.send('absences-saved', result);
		} else {
			event.sender.send('absences-save-failed', result || { success: false, reason: 'unknown' });
		}
	});
});

ipcMain.on('save-editor-persons', (event, editorPersons) => {
	getProgram().saveEditorPersons(editorPersons, result => {
		if (result && result.success) {
			logEvent('Excel Personen bearbeitet', { count: editorPersons ? editorPersons.length : 0 });
			event.sender.send('editor-persons-saved', result);
		} else {
			logEvent('Excel Personen bearbeiten fehlgeschlagen', result || {});
			event.sender.send('editor-persons-save-failed', result || { success: false, reason: 'unknown' });
		}
	});
});

// get probability list
ipcMain.on('get-probabilities', (event, args) => {
	event.sender.send('probability-data', getProgram().getProbabilities());
});

ipcMain.on('get-undo-status', (event, args) => {
	event.sender.send('undo-status', getUndoStatus());
	event.sender.send('redo-status', getRedoStatus());
});

ipcMain.on('undo-last-action', (event, args) => {
	undoLastAction(event);
});

ipcMain.on('redo-last-action', (event, args) => {
	redoLastAction(event);
});

ipcMain.on('export-session-protocol', (event, format) => {
	exportSessionProtocol(event, format);
});

ipcMain.on('import-class-list', (event, args) => {
	importClassList(event);
});

ipcMain.on('confirm-class-list-import', (event, args) => {
	confirmClassListImport(event, args || {});
});

// select specific person
ipcMain.on('select-person', (event, personId) => {
	const result = getProgram().selectSpecificPerson(personId);
	if (result) {
		event.sender.send('name', result);
	}
});

// ok
ipcMain.on('ok', (event, args) => {
	getProgram().saveGrade(args, (result, backupEntry) => {
		handleExcelBackupEntry(event, backupEntry);
		event.sender.send('finished', result);
	});
});

// joker
ipcMain.on('joker', (event, args) => {
    getProgram().setJoker((result, backupEntry) => {
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

	getProgram().applyPendingExcelEntries(pendingEntries, result => {
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

ipcMain.on('run-joker-migration', (event, args) => {
	logEvent('Joker-Migration gestartet');
	getProgram().migrateJokers(result => {
		if (result && result.success) {
			logEvent('Joker-Migration abgeschlossen', { migratedCells: result.migratedCells });
			event.sender.send('joker-migration-done', result);
			event.sender.send('classes', result.worksheets, getExcelFilePath());
		} else {
			logEvent('Joker-Migration fehlgeschlagen', result || {});
			event.sender.send('joker-migration-failed', result || { success: false, reason: 'unknown' });
		}
	});
});

ipcMain.on('cancel-joker-migration', (event, args) => {
	logEvent('Joker-Migration abgebrochen');
});

function handleExcelBackupEntry(event, backupEntry) {
	if (!backupEntry) return;
	const storedBackupEntry = addBackupEntry(backupEntry);
	let pendingEntry;
	logEvent('Backup erstellt', {
		type: storedBackupEntry.type,
		personName: storedBackupEntry.personName,
		excelWriteSucceeded: storedBackupEntry.excelWriteSucceeded,
		reason: storedBackupEntry.excelWriteReason
	});

	if (!storedBackupEntry.excelWriteSucceeded && ['excel-write-failed', 'excel-locked'].includes(storedBackupEntry.excelWriteReason)) {
		pendingEntry = addPendingExcelEntry(storedBackupEntry);
		logEvent(storedBackupEntry.excelWriteReason === 'excel-locked' ? 'Konflikt erkannt' : 'Fehler beim Schreiben', {
			type: storedBackupEntry.type,
			personName: storedBackupEntry.personName,
			reason: storedBackupEntry.excelWriteReason
		});
		event.sender.send('excel-write-pending', pendingEntry);
	}

	if (storedBackupEntry.excelWriteSucceeded || pendingEntry) {
		const undoEntry = Object.assign({}, storedBackupEntry, {
			pendingEntryId: pendingEntry && pendingEntry.id
		});
		sessionEntries.push(undoEntry);
		lastUndoEntry = undoEntry;
		lastRedoEntry = null;
	}

	event.sender.send('pending-excel-status', getPendingExcelEntries().length);
	event.sender.send('undo-status', getUndoStatus());
	event.sender.send('redo-status', getRedoStatus());
}

function getUndoStatus() {
	return {
		available: !!lastUndoEntry,
		entry: lastUndoEntry ? {
			type: lastUndoEntry.type,
			personName: lastUndoEntry.personName,
			grade: lastUndoEntry.grade,
			excelWriteSucceeded: lastUndoEntry.excelWriteSucceeded
		} : null
	};
}

function getRedoStatus() {
	return {
		available: !!lastRedoEntry,
		entry: lastRedoEntry ? {
			type: lastRedoEntry.type,
			personName: lastRedoEntry.personName,
			grade: lastRedoEntry.grade,
			excelWriteSucceeded: lastRedoEntry.excelWriteSucceeded
		} : null
	};
}

function undoLastAction(event) {
	if (!lastUndoEntry) {
		event.sender.send('undo-failed', { success: false, reason: 'nothing-to-undo' });
		return;
	}

	const entry = lastUndoEntry;
	if (!entry.excelWriteSucceeded) {
		removePendingExcelEntries([entry.pendingEntryId].filter(Boolean));
		removeBackupEntries([entry.id]);
		getProgram().deactivateWiggersRuleForEntry(entry);
		removeSessionEntry(entry.id);
		lastUndoEntry = sessionEntries[sessionEntries.length - 1];
		lastRedoEntry = entry;
		logEvent('Letzte Aktion rückgängig gemacht', { type: entry.type, personName: entry.personName, pendingOnly: true });
		event.sender.send('pending-excel-status', getPendingExcelEntries().length);
		event.sender.send('undo-completed', { success: true, entry: entry, pendingOnly: true });
		event.sender.send('undo-status', getUndoStatus());
		event.sender.send('redo-status', getRedoStatus());
		return;
	}

	getProgram().undoLastExcelEntry(entry, result => {
		if (result && result.success) {
			removeBackupEntries([entry.id]);
			removeSessionEntry(entry.id);
			lastUndoEntry = sessionEntries[sessionEntries.length - 1];
			lastRedoEntry = entry;
			logEvent('Letzte Aktion rückgängig gemacht', { type: entry.type, personName: entry.personName });
			event.sender.send('undo-completed', { success: true, entry: entry });
			event.sender.send('undo-status', getUndoStatus());
			event.sender.send('redo-status', getRedoStatus());
			return;
		}

		logEvent('Rückgängig fehlgeschlagen', result || {});
		event.sender.send('undo-failed', result || { success: false, reason: 'unknown' });
		event.sender.send('undo-status', getUndoStatus());
		event.sender.send('redo-status', getRedoStatus());
	});
}

function redoLastAction(event) {
	if (!lastRedoEntry) {
		event.sender.send('redo-failed', { success: false, reason: 'nothing-to-redo' });
		return;
	}

	const entry = Object.assign({}, lastRedoEntry);
	if (!entry.excelWriteSucceeded) {
		const storedBackupEntry = addBackupEntry(Object.assign({}, entry, {
			excelWriteSucceeded: false,
			excelWriteReason: entry.excelWriteReason || 'redo-pending'
		}));
		const pendingEntry = addPendingExcelEntry(storedBackupEntry);
		const sessionEntry = Object.assign({}, storedBackupEntry, {
			pendingEntryId: pendingEntry.id
		});
		sessionEntries.push(sessionEntry);
		lastUndoEntry = sessionEntry;
		lastRedoEntry = null;
		getProgram().activateWiggersRuleForEntry(sessionEntry);
		logEvent('Letzte Aktion erneut angewendet', { type: entry.type, personName: entry.personName, pendingOnly: true });
		event.sender.send('pending-excel-status', getPendingExcelEntries().length);
		event.sender.send('redo-completed', { success: true, entry: sessionEntry, pendingOnly: true });
		event.sender.send('undo-status', getUndoStatus());
		event.sender.send('redo-status', getRedoStatus());
		return;
	}

	getProgram().redoExcelEntry(entry, result => {
		if (result && result.success && (!result.applied || result.applied.length > 0)) {
			const storedBackupEntry = addBackupEntry(Object.assign({}, entry, {
				excelWriteSucceeded: true,
				excelWriteError: undefined,
				excelWriteReason: undefined
			}));
			sessionEntries.push(storedBackupEntry);
			lastUndoEntry = storedBackupEntry;
			lastRedoEntry = null;
			logEvent('Letzte Aktion erneut angewendet', { type: entry.type, personName: entry.personName });
			event.sender.send('redo-completed', { success: true, entry: storedBackupEntry });
			event.sender.send('undo-status', getUndoStatus());
			event.sender.send('redo-status', getRedoStatus());
			return;
		}

		if (result && ['excel-locked', 'excel-write-failed'].includes(result.reason)) {
			const storedBackupEntry = addBackupEntry(Object.assign({}, entry, {
				excelWriteSucceeded: false,
				excelWriteError: result.error,
				excelWriteReason: result.reason
			}));
			const pendingEntry = addPendingExcelEntry(storedBackupEntry);
			const sessionEntry = Object.assign({}, storedBackupEntry, {
				pendingEntryId: pendingEntry.id
			});
			sessionEntries.push(sessionEntry);
			lastUndoEntry = sessionEntry;
			lastRedoEntry = null;
			getProgram().activateWiggersRuleForEntry(sessionEntry);
			logEvent('Redo als Pending gespeichert', { type: entry.type, personName: entry.personName, reason: result.reason });
			event.sender.send('excel-write-pending', pendingEntry);
			event.sender.send('pending-excel-status', getPendingExcelEntries().length);
			event.sender.send('redo-completed', { success: true, entry: sessionEntry, pendingOnly: true });
			event.sender.send('undo-status', getUndoStatus());
			event.sender.send('redo-status', getRedoStatus());
			return;
		}

		logEvent('Redo fehlgeschlagen', result || {});
		event.sender.send('redo-failed', result || { success: false, reason: 'unknown' });
		event.sender.send('undo-status', getUndoStatus());
		event.sender.send('redo-status', getRedoStatus());
	});
}

function removeSessionEntry(id) {
	const index = sessionEntries.findIndex(entry => entry.id === id);
	if (index >= 0) sessionEntries.splice(index, 1);
}

async function exportSessionProtocol(event, format) {
	const normalizedFormat = format === 'pdf' ? 'pdf' : format === 'both' ? 'both' : 'csv';
	if (sessionEntries.length === 0) {
		event.sender.send('session-protocol-empty');
		return;
	}

	if (normalizedFormat === 'both') {
		const folderPaths = dialog.showOpenDialogSync(BrowserWindow.getFocusedWindow(), {
			title: 'Sitzungsprotokoll exportieren',
			properties: ['openDirectory', 'createDirectory']
		});
		if (!folderPaths || folderPaths.length === 0) return;

		const basePath = path.join(folderPaths[0], `Repetierer-Sitzungsprotokoll-${getIsoDate()}`);
		const csvPath = `${basePath}.csv`;
		const pdfPath = `${basePath}.pdf`;
		try {
			fs.writeFileSync(csvPath, buildProtocolCsv(sessionEntries), 'utf8');
			await writeProtocolPdf(pdfPath);
			logEvent('Sitzungsprotokoll exportiert', { format: 'both', filePaths: [pdfPath, csvPath], count: sessionEntries.length });
			event.sender.send('session-protocol-exported', { success: true, filePaths: [pdfPath, csvPath], format: 'both' });
		} catch (error) {
			logEvent('Sitzungsprotokoll Export fehlgeschlagen', { format: 'both', error: getErrorMessage(error) });
			event.sender.send('session-protocol-export-failed', { success: false, reason: 'combined-write-failed', error: getErrorMessage(error) });
		}
		return;
	}

	const extension = normalizedFormat === 'pdf' ? 'pdf' : 'csv';
	const filePath = dialog.showSaveDialogSync(BrowserWindow.getFocusedWindow(), {
		title: 'Sitzungsprotokoll exportieren',
		defaultPath: `Repetierer-Sitzungsprotokoll-${getIsoDate()}.${extension}`,
		filters: [{ name: extension.toUpperCase(), extensions: [extension] }]
	});
	if (!filePath) return;

	if (normalizedFormat === 'pdf') {
		writeProtocolPdf(filePath)
			.then(() => {
				logEvent('Sitzungsprotokoll exportiert', { format: 'pdf', filePath: filePath, count: sessionEntries.length });
				event.sender.send('session-protocol-exported', { success: true, filePath: filePath, format: 'pdf' });
			})
			.catch(error => {
				logEvent('Sitzungsprotokoll Export fehlgeschlagen', { format: 'pdf', error: getErrorMessage(error) });
				event.sender.send('session-protocol-export-failed', { success: false, reason: 'pdf-write-failed', error: getErrorMessage(error) });
			});
		return;
	}

	try {
		fs.writeFileSync(filePath, buildProtocolCsv(sessionEntries), 'utf8');
		logEvent('Sitzungsprotokoll exportiert', { format: 'csv', filePath: filePath, count: sessionEntries.length });
		event.sender.send('session-protocol-exported', { success: true, filePath: filePath, format: 'csv' });
	} catch (error) {
		logEvent('Sitzungsprotokoll Export fehlgeschlagen', { format: 'csv', error: getErrorMessage(error) });
		event.sender.send('session-protocol-export-failed', { success: false, reason: 'csv-write-failed', error: getErrorMessage(error) });
	}
}

function buildProtocolCsv(entries) {
	const rows = [
		['Zeitpunkt', 'Datum', 'Klasse', 'Person', 'Aktion', 'Note', 'Excel-Status', 'Datei']
	].concat(entries.map(entry => [
		formatDateTime(entry.createdAt),
		entry.date || '',
		entry.className || '',
		entry.personName || '',
		entry.type === 'joker' ? 'Joker' : 'Note',
		entry.type === 'grade' ? entry.grade : '',
		entry.excelWriteSucceeded ? 'Excel gespeichert' : 'Excel offen',
		entry.filePath || ''
	]));

	return `\ufeff${rows.map(row => row.map(csvCell).join(';')).join('\n')}`;
}

function csvCell(value) {
	const text = String(value === null || value === undefined ? '' : value);
	return `"${text.replace(/"/g, '""')}"`;
}

function writeProtocolPdf(filePath) {
	const html = buildProtocolHtml(sessionEntries);
	const pdfWindow = new BrowserWindow({
		show: false,
		webPreferences: {
			nodeIntegration: false,
			contextIsolation: true
		}
	});

	return pdfWindow.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`)
		.then(() => pdfWindow.webContents.printToPDF({ printBackground: true, pageSize: 'A4' }))
		.then(data => {
			fs.writeFileSync(filePath, data);
			pdfWindow.close();
		})
		.catch(error => {
			if (!pdfWindow.isDestroyed()) pdfWindow.close();
			throw error;
		});
}

function buildProtocolHtml(entries) {
	const rows = entries.map(entry => `
		<tr>
			<td>${escapeHtml(formatDateTime(entry.createdAt))}</td>
			<td>${escapeHtml(entry.date || '')}</td>
			<td>${escapeHtml(entry.className || '')}</td>
			<td>${escapeHtml(entry.personName || '')}</td>
			<td>${escapeHtml(entry.type === 'joker' ? 'Joker' : 'Note')}</td>
			<td>${escapeHtml(entry.type === 'grade' ? entry.grade : '')}</td>
			<td>${escapeHtml(entry.excelWriteSucceeded ? 'Excel gespeichert' : 'Excel offen')}</td>
		</tr>
	`).join('');

	return `<!doctype html>
	<html lang="de">
	<head>
		<meta charset="utf-8">
		<style>
			body { color: #202125; font-family: Arial, sans-serif; margin: 32px; }
			h1 { font-size: 24px; margin: 0 0 8px; }
			p { color: #555; margin: 0 0 20px; }
			table { border-collapse: collapse; width: 100%; }
			th, td { border: 1px solid #d9d9d9; font-size: 11px; padding: 8px; text-align: left; }
			th { background: #18a890; color: white; font-weight: 700; }
			tr:nth-child(even) td { background: #f5f7f7; }
		</style>
	</head>
	<body>
		<h1>Repetierer Sitzungsprotokoll</h1>
		<p>Exportiert am ${escapeHtml(formatDateTime(new Date().toISOString()))}</p>
		<table>
			<thead>
				<tr>
					<th>Zeitpunkt</th>
					<th>Datum</th>
					<th>Klasse</th>
					<th>Person</th>
					<th>Aktion</th>
					<th>Note</th>
					<th>Excel-Status</th>
				</tr>
			</thead>
			<tbody>${rows}</tbody>
		</table>
	</body>
	</html>`;
}

function importClassList(event) {
	const targetPath = getExcelFilePath();
	if (!targetPath) {
		event.sender.send('class-list-import-failed', { success: false, reason: 'no-target-file' });
		return;
	}
	if (getPendingExcelEntries().length > 0) {
		event.sender.send('class-list-import-failed', { success: false, reason: 'pending-changes' });
		return;
	}

	const selectedFiles = dialog.showOpenDialogSync(BrowserWindow.getFocusedWindow(), {
		title: 'Klassenliste importieren',
		properties: ['openFile'],
		filters: [
			{ name: 'Excel/CSV', extensions: ['xlsx', 'csv'] },
			{ name: 'Excel Dateien (*.xlsx)', extensions: ['xlsx'] },
			{ name: 'CSV Dateien (*.csv)', extensions: ['csv'] }
		]
	});
	if (!selectedFiles || selectedFiles.length === 0) return;

	const sourcePath = selectedFiles[0];
	readClassListImportData(sourcePath)
		.then(importData => {
			const normalizedNames = normalizeImportedNames(importData.names);
			const uniqueNames = normalizedNames.slice(0, 25);
			if (uniqueNames.length === 0) {
				event.sender.send('class-list-import-failed', { success: false, reason: 'no-names-found' });
				return;
			}

			event.sender.send('class-list-import-preview', {
				success: true,
				sourcePath: sourcePath,
				targetPath: targetPath,
				detectedClassName: sanitizeWorksheetName(importData.className || path.basename(sourcePath, path.extname(sourcePath))),
				names: uniqueNames,
				count: uniqueNames.length,
				truncated: normalizedNames.length > uniqueNames.length
			});
		})
		.catch(error => {
			logEvent('Klassenliste Import fehlgeschlagen', { error: getErrorMessage(error), sourcePath: sourcePath });
			event.sender.send('class-list-import-failed', { success: false, reason: 'import-failed', error: getErrorMessage(error) });
		});
}

function confirmClassListImport(event, args) {
	const sourcePath = args.sourcePath;
	const targetPath = getExcelFilePath();
	const className = sanitizeWorksheetName(args.className || '');
	if (!targetPath) {
		event.sender.send('class-list-import-failed', { success: false, reason: 'no-target-file' });
		return;
	}
	if (getPendingExcelEntries().length > 0) {
		event.sender.send('class-list-import-failed', { success: false, reason: 'pending-changes' });
		return;
	}
	if (!sourcePath || !className) {
		event.sender.send('class-list-import-failed', { success: false, reason: 'missing-class-name' });
		return;
	}

	readClassListImportData(sourcePath)
		.then(importData => {
			const normalizedNames = normalizeImportedNames(importData.names);
			const uniqueNames = normalizedNames.slice(0, 25);
			if (uniqueNames.length === 0) {
				event.sender.send('class-list-import-failed', { success: false, reason: 'no-names-found' });
				return;
			}

			return appendImportedClassToWorkbook(targetPath, className, uniqueNames)
				.then(() => {
					saveExcelFilePath(targetPath);
					logEvent('Klassenliste importiert', { sourcePath: sourcePath, targetPath: targetPath, className: className, count: uniqueNames.length });
					getProgram().setFile(targetPath, worksheets => {
						sendClassesOrJokerMigration(event, worksheets, targetPath);
						event.sender.send('class-list-imported', {
							success: true,
							sourcePath: sourcePath,
							targetPath: targetPath,
							className: className,
							count: uniqueNames.length,
							truncated: normalizedNames.length > uniqueNames.length
						});
					});
				});
		})
		.catch(error => {
			const reason = error && error.code ? error.code : 'import-failed';
			logEvent('Klassenliste Import fehlgeschlagen', { error: getErrorMessage(error), sourcePath: sourcePath, targetPath: targetPath });
			event.sender.send('class-list-import-failed', { success: false, reason: reason, error: getErrorMessage(error) });
		});
}

function readClassListImportData(filePath) {
	if (path.extname(filePath).toLowerCase() === '.csv') {
		return Promise.resolve(extractImportDataFromRows(fs.readFileSync(filePath, 'utf8').split(/\r?\n/).map(line => parseCsvLine(line))));
	}

	const workbook = createWorkbook();
	return workbook.xlsx.readFile(filePath).then(() => extractImportDataFromWorksheet(workbook.worksheets[0]));
}

function extractNamesFromCsv(content) {
	const rows = content.split(/\r?\n/).map(line => parseCsvLine(line));
	return extractBestTextColumn(rows);
}

function parseCsvLine(line) {
	const cells = [];
	let current = '';
	let quoted = false;
	for (let i = 0; i < line.length; i++) {
		const char = line[i];
		const next = line[i + 1];
		if (char === '"' && quoted && next === '"') {
			current += '"';
			i++;
		} else if (char === '"') {
			quoted = !quoted;
		} else if ((char === ';' || char === ',') && !quoted) {
			cells.push(current);
			current = '';
		} else {
			current += char;
		}
	}
	cells.push(current);
	return cells;
}

function extractNamesFromWorksheet(worksheet) {
	if (!worksheet) return [];
	const rows = [];
	worksheet.eachRow({ includeEmpty: false }, row => {
		rows.push(row.values.slice(1).map(value => getPlainCellValue(value)));
	});
	return extractBestTextColumn(rows);
}

function extractImportDataFromWorksheet(worksheet) {
	if (!worksheet) return { names: [], className: '' };
	const rows = [];
	worksheet.eachRow({ includeEmpty: false }, row => {
		rows.push(row.values.slice(1).map(value => getPlainCellValue(value)));
	});
	return extractImportDataFromRows(rows);
}

function extractImportDataFromRows(rows) {
	const headerData = extractSchulnetzRows(rows);
	if (headerData.names.length > 0) return headerData;
	return { names: extractBestTextColumn(rows), className: '' };
}

function extractSchulnetzRows(rows) {
	for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
		const headers = rows[rowIndex].map(value => normalizeHeader(value));
		const lastNameColumn = headers.indexOf('nachname');
		const firstNameColumn = headers.indexOf('vorname');
		const classColumn = headers.indexOf('klasse');
		if (lastNameColumn < 0 || firstNameColumn < 0) continue;

		const names = [];
		const classCounts = new Map();
		for (let i = rowIndex + 1; i < rows.length; i++) {
			const firstName = String(rows[i][firstNameColumn] || '').trim();
			const lastName = String(rows[i][lastNameColumn] || '').trim();
			if (!firstName && !lastName) continue;
			names.push(`${firstName} ${lastName}`.trim());

			if (classColumn >= 0) {
				const className = String(rows[i][classColumn] || '').trim();
				if (className) classCounts.set(className, (classCounts.get(className) || 0) + 1);
			}
		}

		return {
			names: names,
			className: getMostCommonValue(classCounts)
		};
	}
	return { names: [], className: '' };
}

function normalizeHeader(value) {
	return String(value || '').trim().toLowerCase().replace(/ü/g, 'ue').replace(/ä/g, 'ae').replace(/ö/g, 'oe');
}

function getMostCommonValue(counts) {
	let bestValue = '';
	let bestCount = 0;
	counts.forEach((count, value) => {
		if (count > bestCount) {
			bestValue = value;
			bestCount = count;
		}
	});
	return bestValue;
}

function extractBestTextColumn(rows) {
	const maxColumns = rows.reduce((max, row) => Math.max(max, row.length), 0);
	let bestNames = [];
	for (let column = 0; column < maxColumns; column++) {
		const names = rows.map(row => row[column]).filter(Boolean);
		if (names.length > bestNames.length) bestNames = names;
	}
	return bestNames;
}

function normalizeImportedNames(names) {
	const seen = new Set();
	return (names || [])
		.map(name => String(name || '').trim())
		.filter(name => name && !isLikelyNameHeader(name))
		.filter(name => {
			const key = name.toLowerCase();
			if (seen.has(key)) return false;
			seen.add(key);
			return true;
		});
}

function isLikelyNameHeader(value) {
	return ['name', 'namen', 'schüler', 'schueler', 'student', 'students', 'person', 'klasse'].includes(String(value).trim().toLowerCase());
}

function getPlainCellValue(value) {
	if (value === null || value === undefined) return '';
	if (typeof value === 'object') {
		if (value.text) return value.text;
		if (value.result) return value.result;
		if (Array.isArray(value.richText)) return value.richText.map(part => part.text || '').join('');
	}
	return value;
}

function writeImportedWorkbook(filePath, className, names) {
	const workbook = createWorkbook();
	const worksheet = workbook.addWorksheet(sanitizeWorksheetName(className || 'Klasse'));
	fillImportedWorksheet(worksheet, names);
	return workbook.xlsx.writeFile(filePath);
}

function appendImportedClassToWorkbook(filePath, className, names) {
	const workbook = createWorkbook();
	return workbook.xlsx.readFile(filePath).then(() => {
		const worksheetName = sanitizeWorksheetName(className || 'Klasse');
		if (workbook.getWorksheet(worksheetName)) {
			const error = new Error(`Klasse existiert bereits: ${worksheetName}`);
			error.code = 'class-exists';
			throw error;
		}
		const worksheet = workbook.addWorksheet(worksheetName);
		fillImportedWorksheet(worksheet, names);
		return workbook.xlsx.writeFile(filePath);
	});
}

function fillImportedWorksheet(worksheet, names) {
	worksheet.getCell('A1').value = 'repetierer';
	const headers = [
		'Name',
		'Note 1',
		'Note 2',
		'Note 3',
		'Note 4',
		'Note 5',
		'Note 6',
		'Joker',
		'',
		'',
		'Datum 1',
		'Datum 2',
		'Datum 3',
		'Datum 4',
		'Datum 5',
		'Datum 6',
		'Joker Datum'
	];
	headers.forEach((header, index) => {
		worksheet.getCell(5, index + 1).value = header;
	});
	names.forEach((name, index) => {
		const row = index + 6;
		worksheet.getCell('A' + row).value = name;
		worksheet.getCell('H' + row).value = 1;
	});
	worksheet.columns.forEach(column => {
		column.width = 16;
	});
	worksheet.getColumn(1).width = 28;
}

function sanitizeWorksheetName(name) {
	const cleaned = String(name || 'Klasse').replace(/[\\/*?:[\]]/g, ' ').trim();
	return (cleaned || 'Klasse').slice(0, 31);
}

function getIsoDate() {
	return new Date().toISOString().slice(0, 10);
}

function formatDateTime(value) {
	try {
		return new Date(value).toLocaleString('de-CH');
	} catch (error) {
		return '';
	}
}

function escapeHtml(value) {
	return String(value === null || value === undefined ? '' : value)
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&#39;');
}

function getErrorMessage(error) {
	if (!error) return '';
	return error.message || String(error);
}

function reloadExcelFile(event) {
	getProgram().reloadExcel(result => {
		if (result && result.success) {
			logEvent('Excel neu geladen', { filePath: result.filePath, className: result.className });
			sendClassesOrJokerMigration(event, result.worksheets, result.filePath);
			if (!getProgram().getJokerMigrationStatus().needsMigration) {
				event.sender.send('excel-reloaded', result);
			}
			event.sender.send('pending-excel-status', getPendingExcelEntries().length);
		} else {
			logEvent('Fehler beim Lesen', result || {});
			event.sender.send('excel-reload-failed', result || { success: false, reason: 'unknown' });
		}
	});
}

function sendClassesOrJokerMigration(event, worksheets, filePath) {
	if (!worksheets) {
		event.sender.send('classes', worksheets, filePath);
		return;
	}

	const migrationStatus = getProgram().getJokerMigrationStatus();
	if (migrationStatus && migrationStatus.needsMigration) {
		logEvent('Joker-Migration erforderlich', migrationStatus);
		event.sender.send('joker-migration-required', migrationStatus, filePath);
		return;
	}

	event.sender.send('classes', worksheets, filePath);
}
