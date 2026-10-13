const { ipcRenderer: aiIpc } = require('electron');
const { createThinkingOrb } = require('./visualFeedback.js');
const aiMotion = require('./uiMotion.js');

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
const aiThinkingOrb = createThinkingOrb(document.getElementById('ai-thinking-orb'), { state: 'composing', paused: true });
const aiWelcomeOrb = createThinkingOrb(document.getElementById('ai-welcome-orb'), { state: 'breathing' });
window.addEventListener('beforeunload', () => { aiThinkingOrb.destroy(); aiWelcomeOrb.destroy(); }, { once: true });

function aiUpdateComposer() {
	aiUi.send.disabled = aiBusy || aiPending || !aiReady || !aiUi.input.value.trim();
	aiUi.input.disabled = !aiReady;
	aiUi.send.textContent = aiBusy ? 'Verarbeitet …' : 'Senden';
	aiUi.input.style.height = 'auto';
	aiUi.input.style.height = `${Math.max(48, Math.min(aiUi.input.scrollHeight + 2, 112))}px`;
}

function aiSetBusy(value) {
	aiBusy = value;
	aiUpdateComposer();
	aiUi.activity.classList.toggle('update-hidden', !value);
	aiThinkingOrb.setPaused(!value);
	if (value) aiThinkingOrb.setState('composing');
	aiUi.messages.setAttribute('aria-busy', String(value));
	if (value) aiUi.activityLabel.textContent = 'Antwort wird erstellt …';
	else if (aiUi.modal.classList.contains('modal-open') && aiReady && !aiPending) aiUi.input.focus();
}

function aiMessage(role, content) {
	const empty = document.getElementById('ai-empty');
	if (empty) { aiWelcomeOrb.destroy(); empty.remove(); }
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
	if (result.type === 'answer') { aiMessage('assistant', result.text); return; }
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
	aiUi.close.focus();
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
	aiThinkingOrb.setState(state);
});

aiUi.form.addEventListener('submit', async event => {
	event.preventDefault();
	const text = aiUi.input.value.trim();
	if (!text || !aiReady || aiBusy || aiPending) return;
	aiUi.input.value = '';
	aiMessage('user', text);
	aiSetBusy(true);
	try {
		const context = typeof window.getRepetiererAiContext === 'function' ? window.getRepetiererAiContext() : {};
		if (aiOpenedStudent && context.classId === aiOpenedStudent.classId && context.studentId === null && context.page !== 'evaluation' && context.page !== 'statistics') { context.studentId = aiOpenedStudent.studentId; context.page = 'student'; }
		aiShowResult(await aiIpc.invoke('ai:ask', { text, context }));
	} catch (error) { aiMessage('error', aiErrorText(error)); }
	finally { aiSetBusy(false); }
});

aiIpc.on('ai-navigate', (event, command) => {
	if (command.page !== 'student') { aiOpenedStudent = null; aiUi.studentCard.classList.add('update-hidden'); }
	if (!window.navigateRepetiererAi(command)) { aiMessage('error', 'Die Klasse konnte nicht geöffnet werden.'); return; }
	if (command.page === 'evaluation') {
		aiIpc.once('class-statistics-data', () => setTimeout(() => {
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
