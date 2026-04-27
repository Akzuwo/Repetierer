const {ipcRenderer, clipboard} = require('electron');
let state = 0; // 0: select file, 1: select class, 2: click start, 3: set grade

// elements
const _container = document.getElementById('container');
const _minimize = document.getElementById('minimize');
const _quit = document.getElementById('quit');
const _startupSplash = document.getElementById('startup-splash');
const _drawerToggle = document.getElementById('drawer-toggle');
const _drawer = document.getElementById('drawer');
const _drawerScrim = document.getElementById('drawer-scrim');
const _drawerStatus = document.getElementById('drawer-status');
const _file = document.getElementById('drawer-file-btn');
const _filePath = document.getElementById('file-path');
const _backupBtn = document.getElementById('backup-btn');
const _logBtn = document.getElementById('log-btn');
const _reloadExcelBtn = document.getElementById('reload-excel-btn');
const _excelEditorBtn = document.getElementById('excel-editor-btn');
const _debugActions = document.getElementById('debug-actions');
const _debugTestAlertBtn = document.getElementById('debug-test-alert-btn');
const _debugFakeUpdateBtn = document.getElementById('debug-fake-update-btn');
const _settingsBtn = document.getElementById('settings-btn');
const _classes = document.getElementById('class-list').children.item(0);
const _start = document.getElementById('start-btn');
const _probabilitiesBtn = document.getElementById('probabilities-btn');
const _flushExcelBtn = document.getElementById('flush-excel-btn');
const _repetition = document.getElementById('repetition');
const _name = document.getElementById('name');
const _label = document.getElementById('grade').children.item(0);
const _grade = document.getElementById('grade').children.item(1);
const _cancel = document.getElementById('res').children.item(1);
const _ok = document.getElementById('res').children.item(2);
const _joker = document.getElementById('res').children.item(3);
const _manualSelectBtn = document.getElementById('manual-select-btn');
const _personModal = document.getElementById('person-modal');
const _personList = document.getElementById('person-list');
const _closeModal = document.getElementById('close-modal');
const _randomSelectBtn = document.getElementById('random-select-btn');
const _className = document.getElementById('class-name');
const _backupModal = document.getElementById('backup-modal');
const _backupList = document.getElementById('backup-list');
const _backupLocation = document.getElementById('backup-location');
const _closeBackupModal = document.getElementById('close-backup-modal');
const _logModal = document.getElementById('log-modal');
const _logList = document.getElementById('log-list');
const _logLocation = document.getElementById('log-location');
const _copyLogBtn = document.getElementById('copy-log-btn');
const _closeLogModal = document.getElementById('close-log-modal');
const _reloadConflictModal = document.getElementById('reload-conflict-modal');
const _conflictFlushBtn = document.getElementById('conflict-flush-btn');
const _conflictForceReloadBtn = document.getElementById('conflict-force-reload-btn');
const _conflictCancelBtn = document.getElementById('conflict-cancel-btn');
const _jokerMigrationModal = document.getElementById('joker-migration-modal');
const _jokerMigrationSummary = document.getElementById('joker-migration-summary');
const _jokerMigrationHelpBtn = document.getElementById('joker-migration-help-btn');
const _runJokerMigrationBtn = document.getElementById('run-joker-migration-btn');
const _cancelJokerMigrationBtn = document.getElementById('cancel-joker-migration-btn');
const _jokerMigrationHelpModal = document.getElementById('joker-migration-help-modal');
const _closeJokerMigrationHelpBtn = document.getElementById('close-joker-migration-help-btn');
const _closeJokerMigrationHelpBottomBtn = document.getElementById('close-joker-migration-help-bottom-btn');
const _probabilitiesModal = document.getElementById('probabilities-modal');
const _probabilityList = document.getElementById('probability-list');
const _closeProbabilitiesModal = document.getElementById('close-probabilities-modal');
const _excelEditorModal = document.getElementById('excel-editor-modal');
const _excelEditorList = document.getElementById('excel-editor-list');
const _excelEditorName = document.getElementById('excel-editor-name');
const _excelEditorJoker = document.getElementById('excel-editor-joker');
const _excelEditorAddBtn = document.getElementById('excel-editor-add-btn');
const _excelEditorSaveBtn = document.getElementById('excel-editor-save-btn');
const _excelEditorCloseBtn = document.getElementById('excel-editor-close-btn');
const _settingsModal = document.getElementById('settings-modal');
const _settingsHelpBtn = document.getElementById('settings-help-btn');
const _settingsHelpModal = document.getElementById('settings-help-modal');
const _closeSettingsHelpBtn = document.getElementById('close-settings-help-btn');
const _closeSettingsHelpBottomBtn = document.getElementById('close-settings-help-bottom-btn');
const _settingsLocation = document.getElementById('settings-location');
const _extraJokerSetting = document.getElementById('extra-joker-setting');
const _probabilityFactorSetting = document.getElementById('probability-factor-setting');
const _boostNeverSelectedSetting = document.getElementById('boost-never-selected-setting');
const _neverSelectedBoostFactorSetting = document.getElementById('never-selected-boost-factor-setting');
const _saveSettingsBtn = document.getElementById('save-settings-btn');
const _closeSettingsModal = document.getElementById('close-settings-modal');
const _updatePanel = document.getElementById('update-panel');
const _updateTitle = document.getElementById('update-title');
const _updateDetail = document.getElementById('update-detail');
const _updateProgress = document.getElementById('update-progress');
const _updateProgressBar = document.getElementById('update-progress-bar');
const _updatePrimaryBtn = document.getElementById('update-primary-btn');
const _updateSecondaryBtn = document.getElementById('update-secondary-btn');

let selectedPersonIds = [];
let excelEditorPersons = [];
let _currentClass;
let _nameSize;
let updatePrimaryAction = requestUpdateDownload;
let restoreAnimationActive = false;

/*
 * events to main
 */

// minimize
_minimize.addEventListener('click', () => {
	animateMinimize();
});

// quit
_quit.addEventListener('click', () => {
	ipcRenderer.send('quit');
});

// file
_file.addEventListener('click', () => {
	ipcRenderer.send('file');
});

_drawerToggle.addEventListener('click', () => {
	toggleDrawer();
});

_drawerScrim.addEventListener('click', () => {
	closeDrawer();
});

// backup
_backupBtn.addEventListener('click', () => {
	ipcRenderer.send('get-backup');
});

_logBtn.addEventListener('click', () => {
	ipcRenderer.send('get-logs');
});

_closeBackupModal.addEventListener('click', () => {
	_backupModal.style.display = 'none';
});

_backupModal.addEventListener('click', (e) => {
	if (e.target === _backupModal) {
		_backupModal.style.display = 'none';
	}
});

_closeLogModal.addEventListener('click', () => {
	_logModal.style.display = 'none';
});

_copyLogBtn.addEventListener('click', () => {
	clipboard.writeText(_logList.innerText || '');
	showUpdatePanel({
		title: 'Logs kopiert',
		detail: 'Die Log-Datei liegt jetzt in der Zwischenablage.',
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true
	});
});

_logModal.addEventListener('click', (e) => {
	if (e.target === _logModal) {
		_logModal.style.display = 'none';
	}
});

// probabilities
_probabilitiesBtn.addEventListener('click', () => {
	ipcRenderer.send('get-probabilities');
});

_flushExcelBtn.addEventListener('click', () => {
	ipcRenderer.send('flush-pending-excel');
});

_reloadExcelBtn.addEventListener('click', () => {
	ipcRenderer.send('reload-excel');
});

_excelEditorBtn.addEventListener('click', () => {
	ipcRenderer.send('get-editor-persons');
});

_conflictFlushBtn.addEventListener('click', () => {
	_reloadConflictModal.style.display = 'none';
	ipcRenderer.send('flush-pending-excel');
});

_conflictForceReloadBtn.addEventListener('click', () => {
	_reloadConflictModal.style.display = 'none';
	ipcRenderer.send('reload-excel-force');
});

_conflictCancelBtn.addEventListener('click', () => {
	_reloadConflictModal.style.display = 'none';
});

_runJokerMigrationBtn.addEventListener('click', () => {
	closeModal(_jokerMigrationModal);
	ipcRenderer.send('run-joker-migration');
});

_cancelJokerMigrationBtn.addEventListener('click', () => {
	closeModal(_jokerMigrationModal);
	ipcRenderer.send('cancel-joker-migration');
	state = 0;
	_filePath.innerText = '';
	_drawerStatus.innerText = 'Migration abgebrochen. Datei wurde nicht geladen.';
	updateState();
});

_jokerMigrationHelpBtn.addEventListener('click', () => {
	closeModal(_jokerMigrationModal, () => openModal(_jokerMigrationHelpModal));
});

_closeJokerMigrationHelpBtn.addEventListener('click', () => {
	closeJokerMigrationHelp();
});

_closeJokerMigrationHelpBottomBtn.addEventListener('click', () => {
	closeJokerMigrationHelp();
});

_jokerMigrationHelpModal.addEventListener('click', (e) => {
	if (e.target === _jokerMigrationHelpModal) {
		closeJokerMigrationHelp();
	}
});

_closeProbabilitiesModal.addEventListener('click', () => {
	_probabilitiesModal.style.display = 'none';
});

_probabilitiesModal.addEventListener('click', (e) => {
	if (e.target === _probabilitiesModal) {
		_probabilitiesModal.style.display = 'none';
	}
});

_excelEditorAddBtn.addEventListener('click', () => {
	addExcelEditorPerson();
});

_excelEditorSaveBtn.addEventListener('click', () => {
	saveExcelEditorPersons();
});

_excelEditorCloseBtn.addEventListener('click', () => {
	closeModal(_excelEditorModal);
});

_excelEditorModal.addEventListener('click', (e) => {
	if (e.target === _excelEditorModal) {
		closeModal(_excelEditorModal);
	}
});

// settings
_settingsBtn.addEventListener('click', () => {
	ipcRenderer.send('get-settings');
});

_saveSettingsBtn.addEventListener('click', () => {
	const probabilityDecreaseFactor = parseFloat(_probabilityFactorSetting.value.replace(',', '.'));
	const neverSelectedBoostFactor = parseFloat(_neverSelectedBoostFactorSetting.value.replace(',', '.'));
	if (!probabilityDecreaseFactor || probabilityDecreaseFactor < 1.1 || probabilityDecreaseFactor > 10 ||
		!neverSelectedBoostFactor || neverSelectedBoostFactor < 1.1 || neverSelectedBoostFactor > 10) {
		error(_saveSettingsBtn);
		return;
	}

	ipcRenderer.send('save-settings', {
		extraJokerAfterThreeGrades: _extraJokerSetting.checked,
		probabilityDecreaseFactor: probabilityDecreaseFactor,
		boostNeverSelected: _boostNeverSelectedSetting.checked,
		neverSelectedBoostFactor: neverSelectedBoostFactor
	});
	closeModal(_settingsModal);
});

_closeSettingsModal.addEventListener('click', () => {
	closeModal(_settingsModal);
});

_settingsModal.addEventListener('click', (e) => {
	if (e.target === _settingsModal) {
		closeModal(_settingsModal);
	}
});

_settingsHelpBtn.addEventListener('click', () => {
	closeModal(_settingsModal, () => openModal(_settingsHelpModal));
});

_closeSettingsHelpBtn.addEventListener('click', () => {
	closeSettingsHelp();
});

_closeSettingsHelpBottomBtn.addEventListener('click', () => {
	closeSettingsHelp();
});

_settingsHelpModal.addEventListener('click', (e) => {
	if (e.target === _settingsHelpModal) {
		closeSettingsHelp();
	}
});

_boostNeverSelectedSetting.addEventListener('change', () => {
	updateNeverSelectedBoostField();
});

_updatePrimaryBtn.addEventListener('click', () => {
	updatePrimaryAction();
});

_updateSecondaryBtn.addEventListener('click', () => {
	hideUpdatePanel();
});

_debugTestAlertBtn.addEventListener('click', () => {
	showUpdatePanel({
		title: 'Testalert',
		detail: 'Debug läuft. Diese Nachricht ist absichtlich oben rechts.',
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true,
		placement: 'toast'
	});
});

_debugFakeUpdateBtn.addEventListener('click', () => {
	showUpdatePanel({
		title: 'Version 99.9.9 ist verfügbar',
		detail: 'Das ist ein Fake-Update aus dem Debug-Menü.',
		primaryText: 'Aktualisieren auf Version 99.9.9',
		secondaryText: 'Später',
		primaryVisible: true,
		secondaryVisible: true,
		primaryAction: () => {
			showUpdatePanel({
				title: 'Fake Update',
				detail: 'Debug-Test erfolgreich. Es wird nichts heruntergeladen.',
				primaryVisible: false,
				secondaryText: 'OK',
				secondaryVisible: true,
				placement: 'toast'
			});
		},
		placement: 'center'
	});
});

// class
function classEvent(e) {
	e.addEventListener('click', () => {
		setActiveClassButton(e);
		ipcRenderer.send('class', e.innerText);
	});
}

// start
_start.addEventListener('click', () => {
	ipcRenderer.send('start');
});

// cancel
_cancel.addEventListener('click', () => {
	state = 2;
	updateState();
});

// ok
_ok.addEventListener('click', () => {
	let v = parseGradeInput(_grade.value);
	if (isValidGrade(v)) {
		ipcRenderer.send('ok', v);
	} else {
		error(_ok);
		showInvalidGradeAlert();
	}

});
_joker.addEventListener('click', () => {
	ipcRenderer.send('joker');

});

// manual select button
_manualSelectBtn.addEventListener('click', () => {
	selectedPersonIds = [];
	ipcRenderer.send('get-persons');
});

// close modal
_closeModal.addEventListener('click', () => {
	_personModal.style.display = 'none';
	selectedPersonIds = [];
});

// close modal when clicking outside
_personModal.addEventListener('click', (e) => {
	if (e.target === _personModal) {
		_personModal.style.display = 'none';
		selectedPersonIds = [];
	}
});

// random select from checked
_randomSelectBtn.addEventListener('click', () => {
	if (selectedPersonIds.length === 0) {
		error(_randomSelectBtn);
		return;
	}
	randomId = selectedPersonIds[Math.floor(Math.random() * selectedPersonIds.length)];
	ipcRenderer.send('select-person', randomId);
	_personModal.style.display = 'none';
	selectedPersonIds = [];
});

/*
 * events from main
 */

// state
/*ipcRenderer.on('state', (e, a) => {
	state = a;
	updateState();
})*/

// classes
ipcRenderer.on('classes', (event, args, filePath) => {
	if (args) {
		state = 1;
		if (filePath) _filePath.innerText = filePath;
		_drawerStatus.innerText = 'Datei geladen. Wähle jetzt eine Klasse.';
		updateState();

		_classes.innerHTML = '';
		_currentClass = null;
		for (let i = 0; i < args.length; i++) {
			let x = document.createElement('button')
			x.className = 'btn-2';
			x.innerText = args[i];
			classEvent(x);
			_classes.appendChild(x);
		}
	} else {
		state = 0;
		_filePath.innerText = '';
		_drawerStatus.innerText = 'Die Datei konnte nicht gelesen werden.';
		updateState();
		error(_file);
	}
})

ipcRenderer.on('saved-file-missing', (event, filePath) => {
	if (filePath) _filePath.innerText = `Gespeicherte Datei nicht gefunden: ${filePath}`;
	_drawerStatus.innerText = 'Die zuletzt gespeicherte Datei wurde nicht gefunden.';
});

ipcRenderer.on('joker-migration-required', (event, status, filePath) => {
	state = 0;
	_classes.innerHTML = '';
	_className.innerText = '';
	if (filePath) _filePath.innerText = filePath;
	const emptyText = `${status.emptyCells || 0} leere Joker-Felder`;
	const usedText = `${status.usedJokerCells || 0} bisherige 1-Werte`;
	_jokerMigrationSummary.innerText = `${emptyText} werden erkannt. Bei der Migration werden leere Felder zu 1 und bisherige 1-Werte zu 0.`;
	_drawerStatus.innerText = 'Joker-Migration erforderlich.';
	updateState();
	openModal(_jokerMigrationModal);
});

ipcRenderer.on('joker-migration-done', (event, result) => {
	showUpdatePanel({
		title: 'Joker angepasst',
		detail: `${result.migratedCells || 0} Joker-Felder wurden migriert.`,
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true
	});
});

ipcRenderer.on('joker-migration-failed', (event, result) => {
	showUpdatePanel({
		title: 'Migration nicht möglich',
		detail: result && result.reason === 'excel-locked'
			? 'Die Excel-Datei scheint geöffnet zu sein. Bitte schließe sie und versuche es erneut.'
			: 'Die Jokerwerte konnten nicht angepasst werden.',
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true
	});
});

ipcRenderer.on('version', (event, appVersion) => {
	document.getElementById("version").innerText = "Repetierer v" + appVersion;
});

ipcRenderer.on('debug-mode', (event, isDebugMode) => {
	if (isDebugMode) {
		_debugActions.classList.remove('update-hidden');
	}
});

ipcRenderer.on('settings-data', (event, settings, paths) => {
	_extraJokerSetting.checked = !!settings.extraJokerAfterThreeGrades;
	_probabilityFactorSetting.value = settings.probabilityDecreaseFactor;
	_boostNeverSelectedSetting.checked = !!settings.boostNeverSelected;
	_neverSelectedBoostFactorSetting.value = settings.neverSelectedBoostFactor;
	updateNeverSelectedBoostField();
	_settingsLocation.innerText = paths ? paths.settingsPath : '';
	openModal(_settingsModal);
});

ipcRenderer.on('update-status', (event, payload) => {
	handleUpdateStatus(payload || {});
});

ipcRenderer.on('window-restored', () => {
	animateRestore();
});

ipcRenderer.on('pending-excel-status', (event, count) => {
	if (count > 0) {
		enableElement(_flushExcelBtn);
		_flushExcelBtn.innerText = `Excel aktualisieren (${count})`;
		_drawerStatus.innerText = `${count} Eintrag${count === 1 ? '' : 'e'} warten auf Excel.`;
	} else {
		disableElement(_flushExcelBtn);
		_flushExcelBtn.innerText = 'Excel aktualisieren';
		if (_filePath.innerText) _drawerStatus.innerText = 'Excel und Backup sind synchron.';
	}
});

// ready
ipcRenderer.on('ready', (event, args) => {
	if (!args) {
		state = 2;
		_className.innerText = `${_currentClass.innerText}`;
		setActiveClassButton(_currentClass);
		updateState();
	} else {
		state = 1;
		_className.innerText = '';
		clearActiveClassButton();
		updateState();
		error(_currentClass);
	}
})

// finished
ipcRenderer.on('finished', (event, args) => {
	if (!args) {
		state = 2;
		updateState();
	} else
		error(_ok);

})

ipcRenderer.on('excel-write-pending', (event, entry) => {
	showUpdatePanel({
		title: 'Excel-Datei ist geöffnet',
		detail: 'Der Eintrag wurde sicher im Programm gespeichert. Schließe Excel und klicke danach im Menü auf "Excel aktualisieren".',
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true
	});
});

ipcRenderer.on('pending-excel-flushed', (event, result) => {
	if (result && result.success && result.remaining === 0) {
		showUpdatePanel({
			title: 'Excel aktualisiert',
			detail: 'Alle zwischengespeicherten Einträge wurden erfolgreich in die Excel-Datei geschrieben.',
			primaryVisible: false,
			secondaryText: 'OK',
			secondaryVisible: true
		});
	} else if (result && result.success) {
		showUpdatePanel({
			title: 'Excel teilweise aktualisiert',
			detail: `Ein Teil der Einträge wurde geschrieben. Noch offen: ${result.remaining}`,
			primaryVisible: false,
			secondaryText: 'OK',
			secondaryVisible: true
		});
	} else {
		showUpdatePanel({
			title: 'Excel-Datei ist geöffnet',
			detail: 'Die Excel-Datei scheint aktuell geöffnet zu sein. Bitte schließe sie zuerst und versuche es erneut.',
			primaryVisible: false,
			secondaryText: 'OK',
			secondaryVisible: true
		});
	}
});

ipcRenderer.on('reload-excel-conflict', (event, pendingCount) => {
	_drawerStatus.innerText = `${pendingCount} lokale Änderung${pendingCount === 1 ? '' : 'en'} blockieren das Neuladen.`;
	_reloadConflictModal.style.display = 'block';
});

ipcRenderer.on('excel-reloaded', (event, result) => {
	if (result && result.hasClass && result.className) {
		state = 2;
		_className.innerText = result.className;
		setActiveClassByName(result.className);
		updateState();
	}
	_drawerStatus.innerText = 'Excel wurde neu geladen.';
	showUpdatePanel({
		title: 'Excel neu geladen',
		detail: 'Die aktuell ausgewählte Excel-Datei wurde erneut eingelesen.',
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true
	});
});

ipcRenderer.on('excel-reload-failed', (event, result) => {
	_drawerStatus.innerText = 'Excel konnte nicht neu geladen werden.';
	showUpdatePanel({
		title: 'Excel konnte nicht geladen werden',
		detail: 'Bitte prüfe, ob die Datei existiert und eine gültige Repetierer-Datei ist.',
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true
	});
});

// name
ipcRenderer.on('name', (event, args) => {
	
	const [name, joker] = args
    
    // check if joker is available
    if (joker) {
        state = 3
    } else {
        state = 4
    }
    
    updateState();
	_grade.value = '';
	_name.innerText = name;
	_nameSize = 10;
	scaleName();

})

// no person available
ipcRenderer.on('no-person-available', (event, args) => {
	error(_start);
});

/// persons list
ipcRenderer.on('persons-list', (event, persons) => {
	_personList.innerHTML = '';
	selectedPersonIds = [];
	
	if (persons && persons.length > 0) {
		persons.forEach(person => {
			const personDiv = document.createElement('div');
			personDiv.className = 'person-item';
			
			const checkbox = document.createElement('input');
			checkbox.type = 'checkbox';
			checkbox.id = `person-${person.id}`;
			checkbox.addEventListener('change', (e) => {
				if (e.target.checked) {
					selectedPersonIds.push(person.id);
				} else {
					selectedPersonIds = selectedPersonIds.filter(id => id !== person.id);
				}
			});
			
			const label = document.createElement('label');
			label.htmlFor = `person-${person.id}`;
			label.textContent = `${person.name}`;
			label.style.marginLeft = '10px';
			label.style.cursor = 'pointer';
			label.style.flexGrow = '1';
			
			personDiv.appendChild(checkbox);
			personDiv.appendChild(label);
			_personList.appendChild(personDiv);
		});
		_personModal.style.display = 'block';
	} else {
		error(_manualSelectBtn);
	}
});

ipcRenderer.on('backup-data', (event, backup, paths) => {
	const entries = backup && backup.entries ? backup.entries.slice().reverse() : [];
	_backupList.innerHTML = '';
	_backupLocation.innerText = paths ? paths.backupPath : '';

	if (entries.length === 0) {
		const empty = document.createElement('p');
		empty.className = 'backup-empty';
		empty.innerText = 'Noch keine Backup-Einträge vorhanden.';
		_backupList.appendChild(empty);
	} else {
		entries.forEach(entry => {
			const item = document.createElement('div');
			item.className = 'backup-item';

			const title = document.createElement('strong');
			title.innerText = entry.type === 'joker' ? 'Joker' : `Note ${entry.grade}`;

			const details = document.createElement('p');
			const createdAt = entry.createdAt ? new Date(entry.createdAt).toLocaleString('de-CH') : 'Unbekanntes Datum';
			const status = entry.excelWriteSucceeded ? 'Excel gespeichert' : 'Excel nicht bestätigt';
			details.innerText = `${createdAt} | ${entry.className || 'Keine Klasse'} | ${entry.personName || 'Unbekannt'} | ${status}`;

			const file = document.createElement('small');
			file.innerText = entry.filePath || '';

			item.appendChild(title);
			item.appendChild(details);
			item.appendChild(file);
			_backupList.appendChild(item);
		});
	}

	_backupModal.style.display = 'block';
});

ipcRenderer.on('log-data', (event, logs, paths) => {
	_logLocation.innerText = paths ? paths.logPath : '';
	_logList.innerText = logs && logs.trim() ? logs : 'Noch keine Logs vorhanden.';
	_logModal.style.display = 'block';
});

ipcRenderer.on('probability-data', (event, probabilities) => {
	_probabilityList.innerHTML = '';

	if (!probabilities || probabilities.length === 0) {
		const empty = document.createElement('p');
		empty.className = 'backup-empty';
		empty.innerText = 'Keine Personen verfügbar.';
		_probabilityList.appendChild(empty);
	} else {
		probabilities
			.slice()
			.sort((a, b) => b.probability - a.probability)
			.forEach(person => {
				const item = document.createElement('div');
				item.className = 'probability-item';

				const header = document.createElement('div');
				header.className = 'probability-header';

				const name = document.createElement('strong');
				name.innerText = person.name;

				const value = document.createElement('span');
				value.innerText = `${(person.probability * 100).toFixed(1)}%`;

				const bar = document.createElement('div');
				bar.className = 'probability-bar';
				const fill = document.createElement('div');
				fill.style.width = `${Math.max(1, person.probability * 100)}%`;
				bar.appendChild(fill);

				const meta = document.createElement('small');
				const jokerText = ` | ${person.joker} Joker übrig`;
				const boostText = person.boostedNeverSelected ? ' | Boost aktiv' : '';
				meta.innerText = `${person.grades} Noten${jokerText} | Gewicht ${person.weight.toFixed(2)}${boostText}`;

				header.appendChild(name);
				header.appendChild(value);
				item.appendChild(header);
				item.appendChild(bar);
				item.appendChild(meta);
				_probabilityList.appendChild(item);
			});
	}

	_probabilitiesModal.style.display = 'block';
});

ipcRenderer.on('editor-persons-data', (event, persons) => {
	if (!_className.innerText) {
		error(_excelEditorBtn);
		showUpdatePanel({
			title: 'Keine Klasse ausgewählt',
			detail: 'Wähle zuerst eine Klasse aus, bevor du die Excel-Datei bearbeitest.',
			primaryVisible: false,
			secondaryText: 'OK',
			secondaryVisible: true,
			placement: 'toast'
		});
		return;
	}

	excelEditorPersons = (persons || []).map(person => ({
		id: person.id,
		name: person.name,
		grades: person.grades,
		joker: Number(person.joker || 0),
		isNew: false,
		deleted: false
	}));
	_excelEditorName.value = '';
	_excelEditorJoker.value = '1';
	renderExcelEditor();
	openModal(_excelEditorModal);
});

ipcRenderer.on('editor-persons-saved', (event, result) => {
	closeModal(_excelEditorModal);
	state = 2;
	updateState();
	showUpdatePanel({
		title: 'Excel gespeichert',
		detail: 'Personen und Joker wurden in der Excel-Datei aktualisiert.',
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true
	});
});

ipcRenderer.on('editor-persons-save-failed', (event, result) => {
	const reason = result && result.reason;
	showUpdatePanel({
		title: 'Excel nicht gespeichert',
		detail: reason === 'excel-locked'
			? 'Die Excel-Datei scheint geöffnet zu sein. Bitte schließe sie und versuche es erneut.'
			: 'Die Personen konnten nicht in die Excel-Datei geschrieben werden.',
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true
	});
});

    
/*
 * other
 */

// update the buttons based on the state
function updateState() {
	switch (state) {
		case 0:
			_classes.innerHTML = '';
			_className.innerText = '';
		case 1:
			disable(_start);
			disable(_name);
			disable(_label);
			disable(_grade);
			disable(_cancel);
			disable(_ok);
			disable(_joker);
			disable(_manualSelectBtn);
			disable(_probabilitiesBtn);
			disable(_excelEditorBtn);
			text();
			break;
		case 2:
			enable(_start);
			disable(_name);
			disable(_label);
			disable(_grade);
			disable(_cancel);
			disable(_ok);
            disable(_joker);
			enable(_manualSelectBtn);
			enable(_probabilitiesBtn);
			enable(_excelEditorBtn);
			ipcRenderer.send('get-pending-excel-status');
			text();
			break;
		case 3:
			disable(_start);
			enable(_name);
			enable(_label);
			enable(_grade);
			enable(_cancel);
			enable(_ok);
            disable(_joker);
			disable(_manualSelectBtn);
			disable(_probabilitiesBtn);
			disable(_excelEditorBtn);
			break;
        case 4:
			disable(_start);
			enable(_name);
			enable(_label);
			enable(_grade);
			enable(_cancel);
			enable(_ok);
            enable(_joker);
			disable(_manualSelectBtn);
			disable(_probabilitiesBtn);
			disable(_excelEditorBtn);
			break;

	}

	function disable(e) {
		disableElement(e);
	}

	function enable(e) {
		enableElement(e);
	}

	function text() {
		_grade.value = '';
		_name.innerText = 'name';
		_nameSize = 10;
		scaleName();
	}
}

// error "handling"
function error(x) {
	x.classList.add('error');

	new Promise(r => setTimeout(r, 600)).then(() => {
		x.classList.remove('error');
	});
}

function showInvalidGradeAlert() {
	showUpdatePanel({
		title: 'Ungültige Note',
		detail: 'Bitte gib eine Note zwischen 1 und 6 ein.',
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true,
		placement: 'toast'
	});
}

function parseGradeInput(value) {
	return Number(String(value).trim().replace(',', '.'));
}

function isValidGrade(value) {
	return Number.isFinite(value) && value >= 1 && value <= 6;
}

function renderExcelEditor() {
	_excelEditorList.innerHTML = '';

	const activePersons = excelEditorPersons.filter(person => !person.deleted);
	if (excelEditorPersons.length === 0) {
		const empty = document.createElement('p');
		empty.className = 'backup-empty';
		empty.innerText = 'Noch keine Personen in dieser Klasse.';
		_excelEditorList.appendChild(empty);
	}

	excelEditorPersons.forEach((person, index) => {
		const item = document.createElement('div');
		item.className = person.deleted ? 'excel-editor-item excel-editor-item-deleted' : 'excel-editor-item';

		const nameInput = document.createElement('input');
		nameInput.type = 'text';
		nameInput.value = person.name;
		nameInput.disabled = person.deleted;
		nameInput.addEventListener('input', () => {
			person.name = nameInput.value;
		});

		const jokerInput = document.createElement('input');
		jokerInput.type = 'number';
		jokerInput.min = '0';
		jokerInput.step = '1';
		jokerInput.value = person.joker;
		jokerInput.disabled = person.deleted;
		jokerInput.addEventListener('input', () => {
			person.joker = parseJokerInput(jokerInput.value);
		});

		const meta = document.createElement('small');
		meta.innerText = person.isNew ? 'Neu' : `${person.grades} Noten`;

		const removeButton = document.createElement('button');
		removeButton.className = person.deleted ? 'btn-1' : 'btn-2';
		removeButton.innerText = person.deleted ? 'Zurück' : 'Löschen';
		removeButton.addEventListener('click', () => {
			person.deleted = !person.deleted;
			renderExcelEditor();
		});

		item.appendChild(nameInput);
		item.appendChild(jokerInput);
		item.appendChild(meta);
		item.appendChild(removeButton);
		_excelEditorList.appendChild(item);
	});

	_excelEditorSaveBtn.innerText = `Speichern (${activePersons.length}/25)`;
}

function addExcelEditorPerson() {
	const name = _excelEditorName.value.trim();
	const joker = parseJokerInput(_excelEditorJoker.value);
	const activeCount = excelEditorPersons.filter(person => !person.deleted).length;

	if (!name || !Number.isInteger(joker) || joker < 0 || activeCount >= 25) {
		error(_excelEditorAddBtn);
		return;
	}

	excelEditorPersons.push({
		id: null,
		name: name,
		grades: 0,
		joker: joker,
		isNew: true,
		deleted: false
	});
	_excelEditorName.value = '';
	_excelEditorJoker.value = '1';
	renderExcelEditor();
}

function saveExcelEditorPersons() {
	const persons = excelEditorPersons
		.filter(person => !person.deleted)
		.map(person => ({
			id: person.id,
			name: String(person.name || '').trim(),
			joker: parseJokerInput(person.joker)
		}));

	const hasInvalidPerson = persons.some(person =>
		!person.name ||
		!Number.isInteger(person.joker) ||
		person.joker < 0
	);

	if (hasInvalidPerson || persons.length > 25) {
		error(_excelEditorSaveBtn);
		showUpdatePanel({
			title: 'Excel nicht gespeichert',
			detail: 'Bitte prüfe Namen und Joker. Joker müssen ganze Zahlen ab 0 sein.',
			primaryVisible: false,
			secondaryText: 'OK',
			secondaryVisible: true,
			placement: 'toast'
		});
		return;
	}

	ipcRenderer.send('save-editor-persons', persons);
}

function parseJokerInput(value) {
	return Number(String(value).trim().replace(',', '.'));
}

// display the version
function version() {
	ipcRenderer.send('get-version');
}

function disableElement(e) {
	if (!e.classList.contains('disabled')) e.classList.add('disabled');
}

function enableElement(e) {
	if (e.classList.contains('disabled')) e.classList.remove('disabled');
}

function toggleDrawer() {
	if (_drawer.classList.contains('drawer-open')) {
		closeDrawer();
	} else {
		_drawer.classList.add('drawer-open');
		_drawerScrim.classList.add('scrim-visible');
		_drawerToggle.classList.add('drawer-toggle-open');
	}
}

function closeDrawer() {
	_drawer.classList.remove('drawer-open');
	_drawerScrim.classList.remove('scrim-visible');
	_drawerToggle.classList.remove('drawer-toggle-open');
}

function setActiveClassButton(button) {
	clearActiveClassButton();
	_currentClass = button;
	if (_currentClass) _currentClass.classList.add('class-selected');
}

function setActiveClassByName(className) {
	const buttons = Array.from(_classes.querySelectorAll('button'));
	const matchingButton = buttons.find(button => button.innerText === className);
	if (matchingButton) setActiveClassButton(matchingButton);
}

function clearActiveClassButton() {
	Array.from(_classes.querySelectorAll('.class-selected')).forEach(button => {
		button.classList.remove('class-selected');
	});
}

function handleUpdateStatus(payload) {
	switch (payload.status) {
		case 'checking':
		case 'not-available':
			break;
		case 'available':
			showUpdatePanel({
				title: `Version ${payload.version} ist verfügbar`,
				detail: 'Eine neue Version von Repetierer kann installiert werden.',
				primaryText: `Aktualisieren auf Version ${payload.version}`,
				secondaryText: 'Später',
				primaryVisible: true,
				secondaryVisible: true,
				primaryAction: requestUpdateDownload,
				placement: 'center'
			});
			break;
		case 'downloading':
			showUpdatePanel({
				title: 'Update wird heruntergeladen',
				detail: `Version ${payload.version} wird heruntergeladen.`,
				progress: payload.percent || 0,
				primaryVisible: false,
				secondaryVisible: false,
				placement: 'center'
			});
			break;
		case 'progress':
			showUpdatePanel({
				title: 'Update wird heruntergeladen',
				detail: `${Number(payload.percent || 0).toFixed(1)}% abgeschlossen.`,
				progress: payload.percent || 0,
				primaryVisible: false,
				secondaryVisible: false,
				placement: 'center'
			});
			break;
		case 'downloaded':
			showUpdatePanel({
				title: 'Update bereit',
				detail: 'Repetierer wird gleich geschlossen und aktualisiert.',
				progress: 100,
				primaryVisible: false,
				secondaryVisible: false,
				placement: 'center'
			});
			break;
		case 'installing':
			showUpdatePanel({
				title: 'Update wird installiert',
				detail: 'Die App startet nach der Installation automatisch neu.',
				progress: 100,
				primaryVisible: false,
				secondaryVisible: false,
				placement: 'center'
			});
			break;
		case 'installed':
			showUpdatePanel({
				title: 'Update erfolgreich installiert',
				detail: `Repetierer läuft jetzt mit Version ${payload.version}.`,
				primaryText: 'OK',
				primaryVisible: false,
				secondaryText: 'Schließen',
				secondaryVisible: true
			});
			break;
		case 'error':
			showUpdatePanel({
				title: 'Update nicht möglich',
				detail: payload.error ? `${payload.message} ${payload.error}` : payload.message,
				primaryVisible: false,
				secondaryText: 'Schließen',
				secondaryVisible: true
			});
			break;
		default:
			break;
	}
}

function showUpdatePanel(options) {
	updatePrimaryAction = options.primaryAction || requestUpdateDownload;
	_updateTitle.innerText = options.title || 'Update';
	_updateDetail.innerText = options.detail || '';

	if (typeof options.progress === 'number') {
		_updateProgress.classList.remove('update-hidden');
		_updateProgressBar.style.width = `${Math.max(0, Math.min(100, options.progress))}%`;
	} else {
		_updateProgress.classList.add('update-hidden');
	}

	if (options.primaryVisible) {
		_updatePrimaryBtn.classList.remove('update-hidden');
		_updatePrimaryBtn.innerText = options.primaryText || 'Aktualisieren';
	} else {
		_updatePrimaryBtn.classList.add('update-hidden');
	}

	if (options.secondaryVisible) {
		_updateSecondaryBtn.classList.remove('update-hidden');
		_updateSecondaryBtn.innerText = options.secondaryText || 'Später';
	} else {
		_updateSecondaryBtn.classList.add('update-hidden');
	}

	_updatePanel.classList.toggle('update-panel-center', options.placement === 'center');
	_updatePanel.classList.remove('update-hidden');
}

function hideUpdatePanel() {
	_updatePanel.classList.add('update-panel-closing');
	setTimeout(() => {
		_updatePanel.classList.add('update-hidden');
		_updatePanel.classList.remove('update-panel-center', 'update-panel-closing');
		updatePrimaryAction = requestUpdateDownload;
	}, 180);
}

function requestUpdateDownload() {
	ipcRenderer.send('update-download-approved');
	showUpdatePanel({
		title: 'Update wird heruntergeladen',
		detail: 'Bitte warte kurz. Repetierer bereitet die neue Version vor.',
		progress: 0,
		primaryVisible: false,
		secondaryVisible: false,
		placement: 'center'
	});
}

function openModal(modal) {
	modal.classList.remove('modal-closing');
	modal.style.display = 'block';
}

function closeModal(modal, afterClose) {
	modal.classList.add('modal-closing');
	setTimeout(() => {
		modal.style.display = 'none';
		modal.classList.remove('modal-closing');
		if (afterClose) afterClose();
	}, 180);
}

function closeSettingsHelp() {
	closeModal(_settingsHelpModal, () => openModal(_settingsModal));
}

function closeJokerMigrationHelp() {
	closeModal(_jokerMigrationHelpModal, () => openModal(_jokerMigrationModal));
}

function updateNeverSelectedBoostField() {
	const isEnabled = _boostNeverSelectedSetting.checked;
	_neverSelectedBoostFactorSetting.disabled = !isEnabled;
	_neverSelectedBoostFactorSetting.closest('.setting-field').classList.toggle('setting-disabled', !isEnabled);
}

function hideStartupSplash() {
	if (!_startupSplash) return;
	setTimeout(() => {
		_startupSplash.classList.add('startup-splash-hidden');
		setTimeout(() => {
			ipcRenderer.send('startup-splash-finished');
			if (_container) {
				_container.classList.remove('startup-active');
				requestAnimationFrame(() => animateRestore());
			}
		}, 650);
	}, 2600);
}

function animateMinimize() {
	if (!_container) {
		ipcRenderer.send('minimize');
		return;
	}

	_container.classList.remove('window-restoring');
	_container.classList.add('window-minimizing');
	setTimeout(() => {
		ipcRenderer.send('minimize');
	}, 180);
}

function animateRestore() {
	if (!_container) return;
	if (restoreAnimationActive) return;
	restoreAnimationActive = true;
	_container.classList.remove('window-minimizing');
	_container.classList.remove('window-restoring');
	void _container.offsetWidth;
	_container.classList.add('window-restoring');
	setTimeout(() => {
		_container.classList.remove('window-restoring');
		restoreAnimationActive = false;
	}, 260);
}

// scale the name to fit the screen
function scaleName() {
	_name.style.fontSize = _nameSize + 'rem';
	let a = parseInt(window.getComputedStyle(_repetition, null).getPropertyValue("width"), 10);
	let b = _name.clientWidth;
	let d = b - a;
	if (d > 0) {
		_nameSize -= 0.1;
		scaleName();
	}
}

Math.clamp = function(a, b, c) {
	return Math.max(b, Math.min(c, a));
}

// automatically scale the name on load
module.exports = [scaleName(), updateState(), version(), hideStartupSplash(), ipcRenderer.send('get-debug-mode'), ipcRenderer.send('load-saved-file'), ipcRenderer.send('get-pending-excel-status')];
