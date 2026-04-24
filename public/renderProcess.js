const {ipcRenderer, clipboard} = require('electron');
let state = 0; // 0: select file, 1: select class, 2: click start, 3: set grade

// elements
const _minimize = document.getElementById('minimize');
const _quit = document.getElementById('quit');
const _drawerToggle = document.getElementById('drawer-toggle');
const _drawer = document.getElementById('drawer');
const _drawerScrim = document.getElementById('drawer-scrim');
const _drawerStatus = document.getElementById('drawer-status');
const _file = document.getElementById('drawer-file-btn');
const _filePath = document.getElementById('file-path');
const _backupBtn = document.getElementById('backup-btn');
const _logBtn = document.getElementById('log-btn');
const _reloadExcelBtn = document.getElementById('reload-excel-btn');
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
const _probabilitiesModal = document.getElementById('probabilities-modal');
const _probabilityList = document.getElementById('probability-list');
const _closeProbabilitiesModal = document.getElementById('close-probabilities-modal');
const _settingsModal = document.getElementById('settings-modal');
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
let _currentClass;
let _nameSize;

/*
 * events to main
 */

// minimize
_minimize.addEventListener('click', () => {
	ipcRenderer.send('minimize');
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

_closeProbabilitiesModal.addEventListener('click', () => {
	_probabilitiesModal.style.display = 'none';
});

_probabilitiesModal.addEventListener('click', (e) => {
	if (e.target === _probabilitiesModal) {
		_probabilitiesModal.style.display = 'none';
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
	_settingsModal.style.display = 'none';
});

_closeSettingsModal.addEventListener('click', () => {
	_settingsModal.style.display = 'none';
});

_settingsModal.addEventListener('click', (e) => {
	if (e.target === _settingsModal) {
		_settingsModal.style.display = 'none';
	}
});

_updatePrimaryBtn.addEventListener('click', () => {
	ipcRenderer.send('update-download-approved');
	showUpdatePanel({
		title: 'Update wird heruntergeladen',
		detail: 'Bitte warte kurz. Repetierer bereitet die neue Version vor.',
		progress: 0,
		primaryVisible: false,
		secondaryVisible: false
	});
});

_updateSecondaryBtn.addEventListener('click', () => {
	hideUpdatePanel();
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
	let v = parseFloat(_grade.value.replace(',', '.'));
	if (v && v >= 1 && v <= 6)
		ipcRenderer.send('ok', v);
	else
		error(_ok);

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

ipcRenderer.on('version', (event, appVersion) => {
	document.getElementById("version").innerText = "Repetierer v" + appVersion;
});

ipcRenderer.on('settings-data', (event, settings, paths) => {
	_extraJokerSetting.checked = !!settings.extraJokerAfterThreeGrades;
	_probabilityFactorSetting.value = settings.probabilityDecreaseFactor;
	_boostNeverSelectedSetting.checked = !!settings.boostNeverSelected;
	_neverSelectedBoostFactorSetting.value = settings.neverSelectedBoostFactor;
	_settingsLocation.innerText = paths ? paths.settingsPath : '';
	_settingsModal.style.display = 'block';
});

ipcRenderer.on('update-status', (event, payload) => {
	handleUpdateStatus(payload || {});
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
				meta.innerText = `${person.grades} Noten | Gewicht ${person.weight.toFixed(2)}`;

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
				secondaryVisible: true
			});
			break;
		case 'downloading':
			showUpdatePanel({
				title: 'Update wird heruntergeladen',
				detail: `Version ${payload.version} wird heruntergeladen.`,
				progress: payload.percent || 0,
				primaryVisible: false,
				secondaryVisible: false
			});
			break;
		case 'progress':
			showUpdatePanel({
				title: 'Update wird heruntergeladen',
				detail: `${Number(payload.percent || 0).toFixed(1)}% abgeschlossen.`,
				progress: payload.percent || 0,
				primaryVisible: false,
				secondaryVisible: false
			});
			break;
		case 'downloaded':
			showUpdatePanel({
				title: 'Update bereit',
				detail: 'Repetierer wird gleich geschlossen und aktualisiert.',
				progress: 100,
				primaryVisible: false,
				secondaryVisible: false
			});
			break;
		case 'installing':
			showUpdatePanel({
				title: 'Update wird installiert',
				detail: 'Die App startet nach der Installation automatisch neu.',
				progress: 100,
				primaryVisible: false,
				secondaryVisible: false
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

	_updatePanel.classList.remove('update-hidden');
}

function hideUpdatePanel() {
	_updatePanel.classList.add('update-hidden');
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
module.exports = [scaleName(), updateState(), version(), ipcRenderer.send('load-saved-file'), ipcRenderer.send('get-pending-excel-status')];
