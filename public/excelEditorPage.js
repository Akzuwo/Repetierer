const notify = require('./notifications');
const {ipcRenderer} = require('electron');
const motion = require('./uiMotion');

function initExcelEditor() {
	const content = document.getElementById('content'), page = document.getElementById('excel-editor-view');
	const select = document.getElementById('excel-editor-class-select'), list = document.getElementById('excel-editor-list');
	const name = document.getElementById('excel-editor-name'), joker = document.getElementById('excel-editor-joker');
	const add = document.getElementById('excel-editor-add-btn'), save = document.getElementById('excel-editor-save-btn'), back = document.getElementById('excel-editor-close-btn');
	const status = document.getElementById('excel-editor-status'), discard = document.getElementById('excel-editor-discard');
	let classes = [], persons = [], loadedClass = '', filePath = '', requestId = 0, dirty = false, saving = false, returnState, pendingAction;
	const visible = () => !page.classList.contains('update-hidden');
	const say = (text, error = false) => {
		if (error || text.startsWith('Gespeichert.')) {notify[error ? 'error' : 'success'](text,{title:'Excel bearbeiten'});status.textContent = '';return;}
		status.textContent = text;
	};
	function busy(value, writing = false) {
		saving = value && writing;
		select.disabled = saving; back.disabled = saving;
		list.inert = value;
		add.disabled = save.disabled = value || !loadedClass;
		name.disabled = joker.disabled = value || !loadedClass;
		page.setAttribute('aria-busy', String(value));
		updateCount();
	}
	function changed() {dirty = true; say('Ungespeicherte Änderungen'); updateCount();}
	function requestAction(action) {
		if (saving) return;
		if (!dirty) {action(); return;}
		pendingAction = action; discard.classList.remove('update-hidden');
		document.getElementById('excel-editor-keep-btn').focus();
	}
	function dismiss() {
		page.classList.add('update-hidden'); content.classList.remove('excel-editor-active');
		document.getElementById('excel-editor-btn').removeAttribute('aria-current');
		requestId++; returnState = null; loadedClass = ''; dirty = false; pendingAction = null;
	}
	function close() {
		requestAction(() => motion.pageChange(() => {
			const previous = returnState;
			dismiss();
			if (previous) {
				for (const [element, hidden] of previous.views) element.classList.toggle('update-hidden', hidden);
				content.classList.toggle('session-active', previous.sessionActive);
				content.classList.toggle('settings-active', previous.settingsActive);
				for (const [element, current] of previous.nav) current ? element.setAttribute('aria-current', current) : element.removeAttribute('aria-current');
			}
			window.restoreTeachingSessionView?.(); document.getElementById('drawer-toggle').focus();
		}));
	}
	function load(className) {
		select.value = className; dirty = false; loadedClass = ''; discard.classList.add('update-hidden'); pendingAction = null;
		document.getElementById('excel-editor-retry-btn').classList.add('update-hidden');
		list.replaceChildren(window.createRepetiererLoadingState('Klasse wird geladen …'));
		document.querySelector('.excel-page-scroll').scrollTop = 0;
		name.value = ''; joker.value = '1'; say(''); busy(true);
		window.repetiererSelects?.refresh();
		ipcRenderer.send('get-editor-persons', {className, requestId:++requestId});
	}
	function open() {
		if (visible()) return;
		classes = [...document.querySelectorAll('#class-list button')].map(button => button.textContent);
		if (!classes.length || document.getElementById('excel-editor-btn').disabled) return;
		motion.pageChange(() => {
			returnState = {views:[...content.children].filter(element => element !== page).map(element => [element,element.classList.contains('update-hidden')]), sessionActive:content.classList.contains('session-active'),settingsActive:content.classList.contains('settings-active'),nav:[...document.querySelectorAll('#drawer button[aria-current], #settings-btn')].map(element => [element,element.getAttribute('aria-current')])};
			for (const [element] of returnState.views) element.classList.add('update-hidden');
			for (const [element] of returnState.nav) element.removeAttribute('aria-current');
			content.classList.remove('session-active','settings-active'); content.classList.add('excel-editor-active');
			page.classList.remove('update-hidden');
			document.getElementById('excel-editor-btn').setAttribute('aria-current','page');
			motion.setDrawerOpen(false);
			select.replaceChildren(...classes.map(value => new Option(value,value)));
			const current = window.getRepetiererAiContext?.().classId;
			load(classes.includes(current) ? current : classes[0]);
			document.getElementById('excel-editor-heading').focus();
		});
	}
	function render(data) {
		persons = data.map(person => ({...person, joker:Number(person.joker || 0), deleted:false}));
		list.replaceChildren();
		persons.forEach((person, index) => {
			const row = document.createElement('div'); row.className = 'excel-editor-item';
			const number = document.createElement('span'); number.className = 'excel-row-number'; number.textContent = String(index+1).padStart(2,'0'); number.setAttribute('aria-hidden','true');
			const input = document.createElement('input'); input.type = 'text'; input.value = person.name;
			input.setAttribute('aria-label', `Name, Person ${index+1}`);
			input.addEventListener('input', () => {person.name = input.value; changed();});
			const count = document.createElement('input'); count.type = 'number'; count.min = '0'; count.step = '1'; count.value = person.joker;
			count.setAttribute('aria-label', `Joker, Person ${index+1}`);
			count.addEventListener('input', () => {person.joker = count.value; changed();});
			const grades = document.createElement('small'); grades.textContent = person.id == null ? 'Neu' : `${person.grades || 0} ${person.grades === 1 ? 'Note' : 'Noten'}`;
			const countLabel = document.createElement('label'); countLabel.className = 'excel-row-joker';
			const caption = document.createElement('span'); caption.textContent = 'Joker'; countLabel.append(caption,count);
			const remove = document.createElement('button'); remove.type = 'button'; remove.className = 'excel-row-remove';
			const paintAction = () => {
				remove.setAttribute('aria-label', person.deleted ? `${person.name} wiederherstellen` : `${person.name} entfernen`);
				remove.title = person.deleted ? 'Wiederherstellen' : 'Person entfernen';
				remove.innerHTML = `<svg class="nav-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="${person.deleted ? 'M4 10a8 8 0 1 1 1 8M4 4v6h6' : 'M4 7h16M9 7V4h6v3M6 7l1 14h10l1-14M10 11v6M14 11v6'}"/></svg>`;
			};
			paintAction();
			remove.addEventListener('click', () => {
				person.deleted = !person.deleted; input.disabled = count.disabled = person.deleted;
				row.classList.toggle('excel-editor-item-deleted',person.deleted); paintAction();
				changed(); updateCount();
			});
			row.append(number,input,countLabel,grades,remove); list.append(row);
		});
		if (!persons.length) {const empty = document.createElement('p'); empty.textContent = 'Noch keine Personen. Füge unten die erste Person hinzu.'; list.append(empty);}
		updateCount(); window.repetiererKristall?.refresh(); motion.refreshScrollEffects();
	}
	function updateCount() {
		const count = persons.filter(person => !person.deleted).length;
		document.getElementById('excel-editor-count').textContent = `${count} ${count === 1 ? 'Person' : 'Personen'}`;
		document.getElementById('excel-editor-count').title = 'Maximal 25 Personen pro Klasse';
		save.textContent = 'Speichern'; save.disabled = saving || !loadedClass || !dirty;
		add.disabled = page.getAttribute('aria-busy') === 'true' || !loadedClass || count >= 25;
	}
	select.addEventListener('change', () => {
		const next = select.value; select.value = loadedClass || next; window.repetiererSelects?.refresh();
		requestAction(() => load(next));
	});
	back.addEventListener('click',close);
	document.getElementById('excel-editor-keep-btn').addEventListener('click', () => {pendingAction = null; discard.classList.add('update-hidden');});
	document.getElementById('excel-editor-discard-btn').addEventListener('click', () => {const action = pendingAction; dirty = false; pendingAction = null; discard.classList.add('update-hidden'); action?.();});
	document.getElementById('excel-editor-retry-btn').addEventListener('click', () => load(select.value));
	add.addEventListener('click', () => {
		const value = Number(joker.value);
		if (!name.value.trim() || !Number.isInteger(value) || value < 0 || persons.filter(p => !p.deleted).length >= 25) {say('Bitte einen Namen und eine ganze Jokerzahl ab 0 angeben. Maximal 25 Personen.',true); return;}
		const next = persons.map(person => ({...person})); next.push({id:null,name:name.value.trim(),joker:value,grades:0});
		// Keep pending deletions when re-rendering after an addition.
		const deleted = next.map(person => person.deleted); render(next);
		deleted.forEach((value,index) => {if (value) list.children[index].querySelector('button').click();});
		name.value = ''; joker.value = '1'; changed(); list.lastElementChild?.scrollIntoView({block:'nearest',behavior:'smooth'});
	});
	save.addEventListener('click', () => {
		if (!loadedClass || saving) return;
		const data = persons.filter(person => !person.deleted).map(person => ({id:person.id,name:String(person.name || '').trim(),joker:Number(person.joker)}));
		if (data.length > 25 || data.some(person => !person.name || !Number.isInteger(person.joker) || person.joker < 0)) {say('Bitte prüfe Namen und Joker. Joker müssen ganze Zahlen ab 0 sein.',true); return;}
		discard.classList.add('update-hidden'); pendingAction = null; busy(true,true); say('Wird gespeichert …');
		ipcRenderer.send('save-editor-persons',data,{className:loadedClass,filePath,requestId});
	});
	const matches = meta => visible() && (!meta || meta.requestId === requestId && meta.className === select.value);
	ipcRenderer.on('editor-persons-data', (_event,data,meta) => {
		if (!matches(meta)) return;
		loadedClass = meta?.className || select.value; filePath = meta?.filePath || '';
		document.getElementById('excel-editor-workbook').textContent = filePath.split(/[\\/]/).pop() || 'Excel-Datei';
		dirty = false; render(data || []); busy(false); say('');
	});
	ipcRenderer.on('editor-persons-load-failed', (_event,result,meta) => {
		if (!matches(meta)) return;
		list.replaceChildren(); loadedClass = ''; busy(false); say('Diese Klasse konnte nicht geladen werden. Bitte erneut versuchen.',true);
		document.getElementById('excel-editor-retry-btn').classList.remove('update-hidden');
	});
	ipcRenderer.on('editor-persons-saved', (_event,result,meta) => {
		if (!matches(meta)) return;
		dirty = false; if (result.persons) render(result.persons); busy(false); say('Gespeichert. Namen und Joker wurden in Excel aktualisiert.');
	});
	ipcRenderer.on('editor-persons-save-failed', (_event,result,meta) => {
		if (!matches(meta)) return;
		busy(false); say(result?.reason === 'excel-locked' ? 'Excel ist geöffnet. Schliesse die Datei und speichere erneut. Deine Änderungen bleiben hier erhalten.' : 'Speichern fehlgeschlagen. Deine Änderungen bleiben hier erhalten. Bitte erneut versuchen.',true);
	});
	ipcRenderer.on('classes', () => {if (visible()) {dismiss(); window.openRepetiererRepetition?.();} requestId++;});
	window.openRepetiererExcelEditor = open;
	window.dismissRepetiererExcelEditorForNavigation = dismiss;
}
module.exports = {initExcelEditor};
