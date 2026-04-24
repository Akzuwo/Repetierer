//const ipcRenderer = require('electron').ipcRenderer;
const {ipcRenderer, remote} = require('electron');
let state = 0; // 0: select file, 1: select class, 2: click start, 3: set grade

// elements
const _minimize = document.getElementById('minimize');
const _quit = document.getElementById('quit');
const _file = document.getElementById('file').children.item(1);
const _filePath = document.getElementById('file-path');
const _backupBtn = document.getElementById('backup-btn');
const _classes = document.getElementById('class-list').children.item(0);
const _start = document.getElementById('start-btn');
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

// backup
_backupBtn.addEventListener('click', () => {
	ipcRenderer.send('get-backup');
});

_closeBackupModal.addEventListener('click', () => {
	_backupModal.style.display = 'none';
});

_backupModal.addEventListener('click', (e) => {
	if (e.target === _backupModal) {
		_backupModal.style.display = 'none';
	}
});

// class
function classEvent(e) {
	e.addEventListener('click', () => {
		_currentClass = e;
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
		updateState();

		_classes.innerHTML = '';
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
		updateState();
		error(_file);
	}
})

ipcRenderer.on('saved-file-missing', (event, filePath) => {
	if (filePath) _filePath.innerText = `Gespeicherte Datei nicht gefunden: ${filePath}`;
});

// ready
ipcRenderer.on('ready', (event, args) => {
	if (!args) {
		state = 2;
		_className.innerText = `${_currentClass.innerText}`;
		updateState();
	} else {
		state = 1;
		_className.innerText = '';
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
			break;

	}

	function disable(e) {
		if (!e.classList.contains('disabled')) e.classList.add('disabled');
	}

	function enable(e) {
		if (e.classList.contains('disabled')) e.classList.remove('disabled');
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
	document.getElementById("version").innerText = "Repetierer v" + remote.app.getVersion();
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
module.exports = [scaleName(), updateState(), version(), ipcRenderer.send('load-saved-file')];
