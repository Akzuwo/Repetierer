const { ipcRenderer: aiIpc } = require('electron');
const aiMotion = require('./uiMotion.js');
const { createThinkingOrb } = require('./visualFeedback.js');

const aiUi = {
	open: document.getElementById('ai-open-btn'), modal: document.getElementById('ai-modal'), close: document.getElementById('ai-close-btn'),
	status: document.getElementById('ai-chat-status'), studentCard: document.getElementById('ai-student-card'), messages: document.getElementById('ai-messages'), activity: document.getElementById('ai-activity'),
	activityLabel: document.getElementById('ai-activity-label'), configure: document.getElementById('ai-configure-btn'),
	form: document.getElementById('ai-form'), input: document.getElementById('ai-input'), send: document.getElementById('ai-send-btn'),
	setup: document.getElementById('ai-setup-btn'), setupStatus: document.getElementById('ai-setup-status'),
	progress: document.getElementById('ai-setup-progress'), progressBar: document.getElementById('ai-setup-progress-bar')
};
let aiReady = false;
let aiBusy = false;
let aiPending = false;
let aiOpenedStudent = null;
let aiReturnFocus = null;
let aiFlightActive = false, currentFlight = null;

const welcomeOrb = createThinkingOrb(document.getElementById('ai-welcome-orb'), { state: 'breathing' });
const typePreference = matchMedia('(prefers-reduced-motion: reduce)');
const writers = new Map();
function finishWriters() { for (const finish of [...writers.values()]) finish(); }
typePreference.addEventListener('change', () => { if(typePreference.matches) { finishWriters(); currentFlight?.finish(); } });
document.addEventListener('visibilitychange', () => { if(document.hidden) { finishWriters(); currentFlight?.finish(); } });
window.addEventListener('beforeunload', () => { welcomeOrb.destroy(); finishWriters(); }, {once:true});
function typeAnswer(text) {
 const item = aiMessage('assistant', '');
 if(typePreference.matches || document.hidden) { item.textContent=text; return; }
 const letters=Array.from(text); let frame=0; const start=performance.now();
 item.dataset.typing='true'; item.setAttribute('aria-busy','true');
 const finish=()=>{cancelAnimationFrame(frame);item.textContent=text;delete item.dataset.typing;item.removeAttribute('aria-busy');writers.delete(item);};
 writers.set(item,finish);
 const tick=time=>{if(!item.isConnected){finish();return;} const count=Math.min(letters.length,Math.max(1,Math.floor((time-start)/Math.min(22,8000/Math.max(1,letters.length))))); item.textContent=letters.slice(0,count).join('');aiUi.messages.scrollTop=aiUi.messages.scrollHeight;if(count<letters.length)frame=requestAnimationFrame(tick);else finish();};
 frame=requestAnimationFrame(tick);
}

function aiUpdateComposer() {
	aiUi.send.disabled = aiBusy || aiPending || !aiReady || !aiUi.input.value.trim();
	aiUi.input.disabled = !aiReady;
	aiUi.send.setAttribute('aria-label', aiBusy ? 'Anfrage wird verarbeitet' : 'Nachricht senden');
	aiUi.input.style.height = 'auto';
	aiUi.input.style.height = `${Math.max(48, Math.min(aiUi.input.scrollHeight + 2, 112))}px`;
}

function aiSetBusy(value) {
	aiBusy = value;
	aiUpdateComposer();
	aiSyncActivity();
	aiUi.activity.dataset.busy = String(value);

	aiUi.messages.setAttribute('aria-busy', String(value));
	if (value) aiUi.activityLabel.textContent = 'Antwort wird erstellt …';
	else if (aiUi.modal.classList.contains('modal-open') && aiReady && !aiPending) aiUi.input.focus();
}
function aiSyncActivity() { aiUi.activity.classList.toggle('update-hidden', !aiBusy || aiFlightActive); }

function aiMessage(role, content) {
	const empty = document.getElementById('ai-empty');
	if (empty) { welcomeOrb.destroy(); empty.remove(); }
	const item = document.createElement('div');
	item.className = `ai-message ai-${role}${role === 'error' ? ' ui-alert ui-alert-error' : ''}`;
	if (role === 'error') item.setAttribute('role', 'alert');
	item.textContent = content;
	aiUi.messages.appendChild(item);
	aiUi.messages.scrollTop = aiUi.messages.scrollHeight;
	return item;
}

function aiErrorText(error) {
	return String(error.message || 'Die Anfrage konnte nicht abgeschlossen werden.').replace(/^Error invoking remote method '[^']+': (?:Error: )?/, '');
}

function aiShowResult(result) {
	if (result.type === 'answer') { typeAnswer(result.text); return; }
	if (result.type !== 'confirmation') throw new Error('Ungültige Assistentenantwort.');
	aiPending = true;
	const card = document.createElement('div');
	card.className = 'ai-confirm-card';
	const heading = document.createElement('strong'); heading.textContent = result.title; card.appendChild(heading);
	for (const line of result.lines || []) { const row = document.createElement('span'); row.textContent = line; card.appendChild(row); }
	const buttons = document.createElement('div'); buttons.className = 'ai-confirm-actions';
	for (const [label, approved] of [['Abbrechen', false], ['Bestätigen', true]]) {
		const button = document.createElement('button'); button.className = approved ? 'btn-1' : 'btn-2'; button.textContent = label;
		button.addEventListener('click', async () => {
			buttons.querySelectorAll('button').forEach(item => item.disabled = true);
			aiSetBusy(true);
			try { const response = await aiIpc.invoke('ai:confirm', { token: result.token, approved }); card.remove(); aiPending = false; aiShowResult(response); }
			catch (error) {
				const message = aiErrorText(error);
				aiMessage('error', message);
				if (/Bestätigung ist ungültig oder abgelaufen/.test(message)) { card.remove(); aiPending = false; }
				else buttons.querySelectorAll('button').forEach(item => item.disabled = false);
			}
			finally { aiSetBusy(false); }
		});
		buttons.appendChild(button);
	}
	card.appendChild(buttons); aiUi.messages.appendChild(card); aiUi.messages.scrollTop = aiUi.messages.scrollHeight;
	if (aiUi.modal.classList.contains('modal-open')) buttons.firstElementChild.focus();
}

async function aiRefreshStatus() {
	try {
		const status = await aiIpc.invoke('ai:status');
		aiReady = status.state === 'ready';
		aiUi.setupStatus.textContent = status.message;
		aiUi.setup.textContent = aiReady ? 'Assistent öffnen' : 'Lokale KI einrichten';
		aiUi.status.textContent = aiReady ? 'Bereit · Deine Daten bleiben auf diesem Gerät' : status.message || 'Lokale KI in den Einstellungen einrichten.';
		aiUi.status.dataset.state = aiReady ? 'ready' : 'unavailable';
	} catch (error) {
		aiReady = false;
		aiUi.setupStatus.textContent = `Fehler: ${error.message}`;
		aiUi.status.textContent = 'Der lokale Assistent ist gerade nicht erreichbar. Prüfe die KI-Einstellungen.';
		aiUi.status.dataset.state = 'unavailable';
	} finally {
		aiUi.configure.classList.toggle('update-hidden', aiReady);
		aiUpdateComposer();
	}
}

async function aiOpen() {
	if (!aiUi.modal.classList.contains('modal-open') && !aiUi.modal.contains(document.activeElement)) aiReturnFocus = document.activeElement;
	aiMotion.openDialog(aiUi.modal);
	document.getElementById('ai-heading').tabIndex = -1;
	document.getElementById('ai-heading').focus();
	await aiRefreshStatus();
	if (aiUi.modal.classList.contains('modal-open') && aiReady && !aiPending) aiUi.input.focus();
}
aiUi.open.addEventListener('click', aiOpen);
aiUi.messages.addEventListener('click', event => {
	const suggestion = event.target.closest('[data-ai-prompt]');
	if (!suggestion) return;
	aiUi.input.value = suggestion.dataset.aiPrompt;
	aiUpdateComposer();
	aiUi.input.focus();
});
function aiClose(afterClose) {
	currentFlight?.finish();
	aiMotion.closeDialog(aiUi.modal, () => {
		if (typeof afterClose === 'function') afterClose();
		else if (aiReturnFocus && aiReturnFocus.isConnected) aiReturnFocus.focus();
	});
}
aiUi.close.addEventListener('click', aiClose);
aiUi.configure.addEventListener('click', () => aiClose(() => {
	document.getElementById('settings-btn').click();
}));
aiUi.modal.addEventListener('click', event => { if (event.target === aiUi.modal) aiClose(); });
document.addEventListener('keydown', event => {
	if (event.ctrlKey && event.key.toLowerCase() === 'k') { event.preventDefault(); aiOpen(); }
	if (!aiUi.modal.classList.contains('modal-open')) return;
	if (event.key === 'Escape') { event.preventDefault(); aiClose(); }
	if (event.key === 'Tab') {
		const controls = Array.from(aiUi.modal.querySelectorAll('button:not(:disabled), textarea:not(:disabled), [tabindex="0"]')).filter(item => item.getClientRects().length);
		const first = controls[0], last = controls[controls.length - 1];
		if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus(); }
		else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus(); }
	}
});
aiUi.input.addEventListener('input', aiUpdateComposer);
aiUi.input.addEventListener('keydown', event => {
	if (event.key === 'Enter' && !event.shiftKey && !event.isComposing) {
		event.preventDefault();
		if (!aiUi.send.disabled) aiUi.form.requestSubmit();
	}
});

aiUi.setup.addEventListener('click', async () => {
	if (aiReady) {
		if (typeof window.closeRepetiererSettings === 'function') window.closeRepetiererSettings(aiOpen);
		else aiOpen();
		return;
	}
	aiUi.setup.disabled = true;
	aiUi.setupStatus.textContent = 'Einrichtung wird gestartet …';
	try {
		const result = await aiIpc.invoke('ai:setup');
		aiUi.setupStatus.textContent = result.message;
		aiUi.progress.classList.add('update-hidden');
		await aiRefreshStatus();
	} catch (error) { aiUi.setupStatus.textContent = `Fehler: ${error.message}`; }
	finally { aiUi.setup.disabled = false; }
});
aiIpc.on('ai-setup-progress', (event, progress) => {
	aiUi.setupStatus.textContent = progress.message;
	aiUi.progress.classList.toggle('update-hidden', progress.state !== 'downloading');
	if (typeof progress.percent === 'number') aiUi.progressBar.style.width = `${progress.percent}%`;
});
aiIpc.on('ai-activity', (event, activity) => {
	if (!aiBusy) return;
	aiUi.activityLabel.textContent = activity.label || 'Klassendaten werden geprüft …';
	const state = activity.tool === 'searchStudents' ? 'searching' : ['createRepetitionGrade', 'addParticipationPoints'].includes(activity.tool) ? 'solving' : 'working';
	aiUi.activity.dataset.activity = state;
});

aiUi.form.addEventListener('submit', async event => {
	event.preventDefault();
	const text = aiUi.input.value.trim();
	if (!text || !aiReady || aiBusy || aiPending) return;
	const source = aiUi.input.getBoundingClientRect();
 aiUi.input.value = '';
 aiFlightActive = !typePreference.matches && !document.hidden;
 aiSetBusy(true);
 const message = aiMessage('user', text);
 let landed = Promise.resolve();
 if (aiFlightActive) {
   let resolveLanding; landed = new Promise(resolve => {resolveLanding=resolve;});
   const target = message.getBoundingClientRect();
   const ghost = document.createElement('div'); ghost.className = 'ai-message ai-user ai-send-flight'; ghost.textContent = text; ghost.setAttribute('aria-hidden','true');
   ghost.style.cssText = `position:fixed;pointer-events:none;z-index:2000;left:${source.left}px;top:${source.top}px;width:${target.width}px;height:${target.height}px;box-sizing:border-box;overflow:hidden`;
   document.body.append(ghost); message.dataset.inFlight='true';
   const flight=ghost.animate([{transform:'translate(0,0)',opacity:0},{transform:`translate(${target.left-source.left}px,${target.top-source.top}px)`,opacity:1}],{duration:480,easing:'cubic-bezier(.7,0,.84,0)'});
   currentFlight=flight;
   const finish=()=>{ghost.remove();delete message.dataset.inFlight;aiUi.messages.scrollTop=aiUi.messages.scrollHeight;aiFlightActive=false;currentFlight=null;aiSyncActivity();resolveLanding();};flight.onfinish=finish;flight.oncancel=finish;
 }

 aiUi.messages.scrollTop=aiUi.messages.scrollHeight;
 requestAnimationFrame(()=>{aiUi.messages.scrollTop=aiUi.messages.scrollHeight;});
	try {
		const context = typeof window.getRepetiererAiContext === 'function' ? window.getRepetiererAiContext() : {};
		if (aiOpenedStudent && context.classId === aiOpenedStudent.classId && context.studentId === null && context.page !== 'evaluation' && context.page !== 'statistics') { context.studentId = aiOpenedStudent.studentId; context.page = 'student'; }
		const result=await aiIpc.invoke('ai:ask', { text, context });
		await landed;
		aiShowResult(result);
	} catch (error) { await landed; aiMessage('error', aiErrorText(error)); }
	finally { aiSetBusy(false); }
});

aiIpc.on('ai-navigate', (event, command) => {
	if (command.page !== 'student') { aiOpenedStudent = null; aiUi.studentCard.classList.add('update-hidden'); }
	if (!window.navigateRepetiererAi(command)) { aiMessage('error', 'Die Klasse konnte nicht geöffnet werden.'); return; }
	if (command.page === 'evaluation') {
		aiIpc.once('class-statistics-data', () => setTimeout(() => {
			window.selectRepetiererStatsClass?.(command.classId);
			const card = Array.from(document.querySelectorAll('#stats-list .stats-class-card')).find(item => item.querySelector('h2, h3, h4')?.textContent === command.classId);
			if (card) card.scrollIntoView({ behavior: 'smooth', block: 'center' });
		}, 120));
	} else if (command.page === 'student') {
		aiOpenedStudent = { classId: command.classId, studentId: command.studentId };
		aiUi.status.textContent = `${command.name} · ${command.classId}`;
		aiUi.studentCard.textContent = `${command.name} · Klasse ${command.classId} · ${command.gradeCount} Repetitionsnoten · ${command.remainingJokers} Joker`;
		aiUi.studentCard.classList.remove('update-hidden');
	}
});

aiRefreshStatus();
