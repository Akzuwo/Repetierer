const notify = require('./notifications');
const roundNoticeSignatures = new Set();
const { ipcRenderer } = require('electron');
const assessmentMotion = require('./uiMotion.js');
const {deadlineBounds,validateDeadline,zurichDate,DAY} = require('../src/assessment/deadlinePolicy');

const elements = {
	view: document.getElementById('assessment-view'),
	body: document.getElementById('assessment-body'),
	heading: document.getElementById('assessment-heading'),
	classSelect: document.getElementById('assessment-class-select'),
	back: document.getElementById('assessment-back-btn'),
	operationStatus: document.getElementById('assessment-operation-status'),
	classPanel: document.getElementById('class'),
	repetition: document.getElementById('repetition'),
	stats: document.getElementById('stats-view'),
	sessionHeader: document.getElementById('session-header'),
	participation: document.getElementById('participation-view'),
	sessionEmpty: document.getElementById('session-empty-view'),
	title: document.getElementById('title'),
	drawer: document.getElementById('drawer'),
	drawerScrim: document.getElementById('drawer-scrim'),
	drawerToggle: document.getElementById('drawer-toggle'),
	repetitionNav: document.getElementById('repetition-nav-btn'),
	assessmentNav: document.getElementById('assessment-nav-btn'),
	settings: document.getElementById('settings-btn'),
	assessmentConfigStatus: document.getElementById('assessment-config-status'),
	assessmentApiStatus: document.getElementById('assessment-api-status'),
	assessmentConfigImport: document.getElementById('assessment-config-import-btn'),
	assessmentApiTest: document.getElementById('assessment-api-test-btn'),
	assessmentConfigRemove: document.getElementById('assessment-config-remove-btn'),
	brevoStatus: document.getElementById('brevo-status'),
	brevoImport: document.getElementById('brevo-import-btn'),
	brevoTest: document.getElementById('brevo-test-btn'),
	brevoRemove: document.getElementById('brevo-remove-btn'),
	confirmModal: document.getElementById('assessment-confirm-modal'),
	confirmText: document.getElementById('assessment-confirm-text'),
	confirmSend: document.getElementById('assessment-confirm-send-btn'),
	confirmCancel: document.getElementById('assessment-confirm-cancel-btn'),
	deleteModal: document.getElementById('assessment-delete-modal'),
	deleteText: document.getElementById('assessment-delete-text'),
	deleteSentWarning: document.getElementById('assessment-delete-sent-warning'),
	deleteConfirm: document.getElementById('assessment-delete-confirm-btn'),
	deleteCancel: document.getElementById('assessment-delete-cancel-btn')
};

let loadedClasses = [];
let selectedClass = '';
let overview = null;
let roundView = null;
let teacherIndex = 0;
let pendingSendMode = 'pending';
let brevoState = { configured: false };
let busy = false,viewRequest=0,renderedTeacherKey=null;
const teacherDrafts=new Map();
function captureTeacherDraft(){
 if(!renderedTeacherKey || !elements.body.querySelector('.teacher-assessment'))return;
 const ratings={},comments={};
 elements.body.querySelectorAll('[data-teacher-criterion]').forEach(input=>{ratings[input.dataset.teacherCriterion]={value:Number(input.value),rated:input.dataset.rated==='true'};});
 elements.body.querySelectorAll('[data-teacher-comment]').forEach(input=>{comments[input.dataset.teacherComment]=input.value;});
 teacherDrafts.set(renderedTeacherKey,{ratings,comments,comment:document.getElementById('teacher-comment').value});
}
function teacherKey(student){return JSON.stringify([roundView.filePath||'',selectedClass,roundView.round.id,student.id]);}
const autoSyncErrors = new Set();
ipcRenderer.on('assessment:auto-synced', async (_event,payload) => {
	if(payload.className!==selectedClass || elements.view.classList.contains('update-hidden'))return;
	autoSyncErrors.delete(`${payload.filePath}:${payload.roundId}`);
	if(roundView?.round.id===payload.roundId && (!roundView.filePath || roundView.filePath===payload.filePath)) {
		if(busy)return;
		roundView=payload.view;
		// Background updates must never replace an unfinished teacher assessment,
		// selected response, date field or keyboard focus.
		if(!elements.body.querySelector('.teacher-assessment') && !elements.body.contains(document.activeElement)) renderRound(true);
	} else if(!roundView && !busy && !elements.body.contains(document.activeElement)) {
		const requestedClass=selectedClass;
		try {const next=await invoke('assessment:get-overview',{className:requestedClass});if(selectedClass===requestedClass && !roundView && !busy && !elements.body.contains(document.activeElement)) {overview=next;renderOverview();}}catch{}
	}
});
ipcRenderer.on('assessment:auto-sync-failed',(_event,payload)=>{
	if(payload.className!==selectedClass || elements.view.classList.contains('update-hidden'))return;
	const key=`${payload.filePath}:${payload.roundId}`;
	if(!autoSyncErrors.has(key)) {autoSyncErrors.add(key);notify.warning(`${payload.message} Automatischer Wiederholungsversuch in 30 Sekunden.`,{title:'Synchronisation',key:'assessment-auto-sync'});}
});

ipcRenderer.on('classes', (event, classes) => {
	loadedClasses = Array.isArray(classes) ? classes : [];
	renderClassOptions();
});

elements.assessmentNav.addEventListener('click', openAssessment);
elements.repetitionNav.addEventListener('click', openRepetition);
elements.classSelect.addEventListener('change', async () => {
 captureTeacherDraft();
	selectedClass = elements.classSelect.value;
	roundView = null;
 if(elements.view.classList.contains('update-hidden'))return;
	if (selectedClass) await loadOverview();
	else renderEmpty('Bitte wähle eine Klasse.');
});
elements.back.addEventListener('click', async () => {
	if (elements.body.querySelector('.teacher-assessment')) {
		renderRound();
		return;
	}
	if (roundView) {
		roundView = null;
		await loadOverview();
	}
});

elements.settings.addEventListener('click', loadIntegrationStatuses);
const packageImport = document.getElementById('integration-package-import-btn');
const packageExport = document.getElementById('integration-package-export-btn');
const packageStatus = document.getElementById('integration-package-status');
for (const [button, channel] of [[packageImport,'assessment:config-import'],[packageExport,'assessment:config-export']]) button.addEventListener('click', async () => {
	packageImport.disabled = packageExport.disabled = true;
	packageStatus.textContent = channel.endsWith('export') ? 'Paket wird erstellt …' : 'Paket wird geprüft …';
	try {
		const result = await invoke(channel);
		await loadIntegrationStatuses();
		packageStatus.textContent = '';
		if (!result.cancelled) notify.success(result.exported ? 'Das gemeinsame .repetierer-Paket wurde gespeichert.' : 'Konfiguration importiert. Die Verbindungsanzeigen wurden aktualisiert.',{title:'Integrationen'});
	} catch (error) { packageStatus.textContent = ''; notify.error(error.message,{title:'Integrationen'}); }
	finally {packageImport.disabled = packageExport.disabled = false;}
});
elements.assessmentConfigImport.addEventListener('click', () => integrationAction('assessment:config-import', elements.assessmentConfigStatus, result => {
		loadIntegrationStatuses();
		if (result && result.deleteError) showNotice(result.deleteError, 'warning');
		if (result && result.formWarning) showNotice(result.formWarning, 'warning');
}));
elements.assessmentApiTest.addEventListener('click', async () => {
	try {
		const status = await invoke('assessment:api-test');
		setIntegrationMessage(elements.assessmentApiStatus, status.fallback ? 'API: Erreichbar über akzuwo.ch (Fallback)' : 'API: Erreichbar', 'success');
	} catch (error) {
		setIntegrationMessage(elements.assessmentApiStatus, `API: ${error.message}`, 'error');
	}
});
elements.assessmentConfigRemove.addEventListener('click', () => integrationAction('assessment:config-remove', elements.assessmentConfigStatus, loadIntegrationStatuses));
elements.brevoImport.addEventListener('click', () => integrationAction('assessment:config-import', elements.brevoStatus, result => {
		loadIntegrationStatuses();
		if (result && result.deleteError) showNotice(result.deleteError, 'warning');
		if (result && result.formWarning) showNotice(result.formWarning, 'warning');
	}));
elements.brevoTest.addEventListener('click', async () => {
	const result = await invoke('assessment:brevo-test');
	setIntegrationMessage(elements.brevoStatus, result.ok ? 'Verbindung erfolgreich' : result.error, result.ok ? 'success' : 'error');
});
elements.brevoRemove.addEventListener('click', () => integrationAction('assessment:brevo-remove', elements.brevoStatus, loadIntegrationStatuses));

elements.confirmCancel.addEventListener('click', closeConfirm);
elements.confirmModal.addEventListener('click', event => { if (event.target === elements.confirmModal) closeConfirm(); });
elements.confirmSend.addEventListener('click', async () => {
	closeConfirm();
	await sendRound(pendingSendMode);
});
elements.deleteCancel.addEventListener('click', closeDeleteConfirm);
elements.deleteModal.addEventListener('click', event => { if (event.target === elements.deleteModal) closeDeleteConfirm(); });
elements.deleteConfirm.addEventListener('click', deleteRound);

async function openAssessment() {
 if(busy && !elements.view.classList.contains('update-hidden')){closeDrawer();return;}
 if(busy){viewRequest++;setBusy(false);}
	if (assessmentMotion.beginPageChange(() => { openAssessment(); })) return;
	window.dismissRepetiererSettingsForNavigation?.();
	window.dismissRepetiererExcelEditorForNavigation?.();
	closeDrawer();
	elements.repetitionNav.removeAttribute('aria-current');
	elements.assessmentNav.setAttribute('aria-current', 'page');
	document.getElementById('stats-view-btn').removeAttribute('aria-current');
	document.getElementById('content').classList.remove('session-active');
	elements.classPanel.classList.add('update-hidden');
	elements.repetition.classList.add('update-hidden');
	elements.stats.classList.add('update-hidden');
	elements.sessionHeader.classList.add('update-hidden');
	elements.participation.classList.add('update-hidden');
 document.getElementById('homework-view').classList.add('update-hidden');
	elements.sessionEmpty.classList.add('update-hidden');
 document.getElementById('session-attendance-view').classList.add('update-hidden');
	elements.view.classList.remove('update-hidden');
	assessmentMotion.revealView(elements.view);
	elements.title.innerText = 'Repetierer';
	elements.title.dataset.text = 'Repetierer';
	const activeClass = document.querySelector('#class-list .class-selected');
	if (!selectedClass && activeClass) selectedClass = activeClass.innerText;
	renderClassOptions();
	if (selectedClass) await loadOverview();
	else renderEmpty(loadedClasses.length ? 'Bitte wähle eine Klasse.' : 'Bitte zuerst eine Repetierer-Datei laden.');
}

function openRepetition() {
 captureTeacherDraft();viewRequest++;setBusy(false);
	if (assessmentMotion.beginPageChange(() => { openRepetition(); })) return;
	window.dismissRepetiererSettingsForNavigation?.();
	window.dismissRepetiererExcelEditorForNavigation?.();
	closeDrawer();
	elements.assessmentNav.removeAttribute('aria-current');
	elements.repetitionNav.setAttribute('aria-current', 'page');
	document.getElementById('stats-view-btn').removeAttribute('aria-current');
	elements.view.classList.add('update-hidden');
	elements.stats.classList.add('update-hidden');
	if (typeof window.restoreTeachingSessionView === 'function') window.restoreTeachingSessionView();
	else {
		elements.classPanel.classList.remove('update-hidden');
		elements.repetition.classList.remove('update-hidden');
	}
	elements.title.innerText = 'Repetierer';
	elements.title.dataset.text = 'Repetierer';
}

window.openRepetiererRepetition = openRepetition;

function closeDrawer() {
	assessmentMotion.setDrawerOpen(false);
}

function renderClassOptions() {
	const current = selectedClass;
	elements.classSelect.innerHTML = '<option value="">Klasse wählen</option>';
	for (const className of loadedClasses) {
		const option = document.createElement('option');
		option.value = className;
		option.innerText = className;
		elements.classSelect.appendChild(option);
	}
	if (loadedClasses.includes(current)) elements.classSelect.value = current;
}

async function loadOverview() {
 captureTeacherDraft();roundView=null;renderedTeacherKey=null;const request=++viewRequest,className=selectedClass;
	setBusy(true, 'Beurteilungsrunden werden geladen …', true, 'overview');
	try {
		const result=await invoke('assessment:get-overview', {className});
  if(request!==viewRequest || className!==selectedClass || elements.view.classList.contains('update-hidden'))return;
  overview=result;renderOverview();
	} catch (error) {
		if(request===viewRequest)renderError(error.message);
	} finally {
		if(request===viewRequest)setBusy(false);
	}
}

function renderOverview() {
	if (assessmentMotion.beginPageChange(() => { renderOverview(); })) return;
	elements.body.scrollTop = 0;
	roundView = null;
	elements.heading.innerText = `${overview.name} · Beurteilungsrunden`;
	elements.back.classList.add('update-hidden');
	const rounds = overview.rounds.map(round => `
		<button class="assessment-round-card" data-round-id="${escapeAttribute(round.id)}">
			<span><strong>${escapeHtml(round.title)}</strong><small>${round.deadline ? `Abgabe ${formatDate(round.deadline)}` : 'Keine Abgabefrist'}</small></span>
			<span class="assessment-round-counts"><b>${round.selfCount}/${round.studentCount}</b> Selbst · <b>${round.teacherCount}/${round.studentCount}</b> Lehrer · <b>${round.sentCount}/${round.studentCount}</b> Versand</span>
		</button>`).join('');
	elements.body.innerHTML = `
		<section class="assessment-create">
			<h3>Neue Beurteilungsrunde</h3>
			<div class="assessment-create-fields">
				<label>Titel / Semester<input id="assessment-round-title" maxlength="80" placeholder="HS 2026"></label>
				<label>Abgabefrist<input id="assessment-round-deadline" type="date" min="${deadlineBounds().min}" max="${deadlineBounds().max}" aria-describedby="assessment-deadline-hint"></label>
				<button class="btn-1" id="assessment-create-round-btn">Erstellen</button>
			</div>
			<p id="assessment-deadline-hint">Abgabefrist: heute bis ${formatDate(deadlineBounds().max)} (höchstens 90 Tage). Ohne Datum gilt eine Frist von 30 Tagen.</p>
		</section>
		<section class="assessment-round-list">
			<h3>Bestehende Runden</h3>
			${rounds || '<p class="assessment-empty">Für diese Klasse gibt es noch keine Beurteilungsrunde.</p>'}
		</section>`;
	document.getElementById('assessment-create-round-btn').addEventListener('click', createRound);
	elements.body.querySelectorAll('[data-round-id]').forEach(button => button.addEventListener('click', () => openRound(button.dataset.roundId)));
}

async function createRound() {
	if (busy) return;
 const request=++viewRequest,className=selectedClass;
	const title = document.getElementById('assessment-round-title').value;
	const deadline = document.getElementById('assessment-round-deadline').value;
	try {
		validateDeadline(deadline);
		setBusy(true, 'Runde wird erstellt …');
		const result=await invoke('assessment:create-round', {className,title,deadline});
  if(request!==viewRequest || className!==selectedClass || elements.view.classList.contains('update-hidden'))return;
  overview=result;
		renderOverview();
	} catch (error) {
		if(request===viewRequest)showNotice(error.message, 'error');
	} finally {
		if(request===viewRequest)setBusy(false);
	}
}

async function openRound(roundId) {
	if (busy) return;
 const request=++viewRequest,className=selectedClass;
	try {
		setBusy(true, 'Beurteilungsrunde wird geladen …', true);
		const result=await invoke('assessment:get-round', {className,roundId});
  if(request!==viewRequest || className!==selectedClass || elements.view.classList.contains('update-hidden'))return;
  roundView=result;
		await refreshBrevoState();
		if(request===viewRequest && !elements.view.classList.contains('update-hidden'))renderRound();
	} catch (error) {
		if(request===viewRequest)renderError(error.message);
	} finally {
		if(request===viewRequest)setBusy(false);
	}
}

function renderRound(background = false) {
 captureTeacherDraft();
	if (!background && assessmentMotion.beginPageChange(() => { renderRound(); })) return;
	const previousScroll=elements.body.scrollTop;
	elements.body.scrollTop = 0;
	const { round, summary, students, invitation, invitationError, deliveryValidation } = roundView;
	const assessmentConfigured = roundView.assessmentConfigured && !!round.cloudStatus;
	elements.heading.innerText = `${selectedClass} · ${round.title}`;
	elements.back.classList.remove('update-hidden');
	for (const message of [invitationError, summary.lastSyncResult?.invalid ? `${summary.lastSyncResult.invalid} ungültige Antworten wurden beim letzten Abruf ignoriert.` : '', summary.lastSyncResult?.ignoredRound ? `${summary.lastSyncResult.ignoredRound} Antworten aus einer anderen Runde wurden ignoriert.` : '']) {
		const signature = `${round.id}:${summary.lastSyncAt}:${message}`;
		if(!background && message && !roundNoticeSignatures.has(signature)) {roundNoticeSignatures.add(signature);showNotice(message,'warning');}
	}
	const syncLabel = summary.lastSyncAt ? `Zuletzt ${formatDateTime(summary.lastSyncAt)}` : 'Noch nicht synchronisiert';
	const studentRows = students.map(student => {
		const assessment = round.assessments[student.id] || {};
		return `<button type="button" class="assessment-student-row" data-teacher-index="${students.indexOf(student)}"><strong>${escapeHtml(student.name)}</strong><span>Selbst ${statusIcon(assessment.studentAnswers)}</span><span>Lehrer ${statusIcon(assessment.teacherAnswers)}</span><span>Versand ${mailIcon(assessment.mailStatus)}</span></button>`;
	}).join('');
	const unmatched = (round.unmatchedResponses || []).map(response => `
		<div class="assessment-match-row">
			<span><strong>${escapeHtml(response.name || 'Name fehlt')}</strong><small>${escapeHtml(response.email || 'E-Mail fehlt')}</small></span>
			<select data-match-response="${escapeAttribute(response.externalResponseId)}"><option value="">Schüler zuordnen …</option>${students.map(student => `<option value="${escapeAttribute(student.id)}">${escapeHtml(student.name)}</option>`).join('')}</select>
			<button class="btn-1" data-assign-response="${escapeAttribute(response.externalResponseId)}">Zuordnen</button>
		</div>`).join('');
	const duplicates = (round.duplicates || []).map(duplicate => `
		<div class="assessment-duplicate-row">
			<span><strong>${escapeHtml(duplicate.email)}</strong><small>${duplicate.count} Antworten</small></span>
			<select data-select-response>${duplicate.responses.map(response => `<option value="${escapeAttribute(response.externalResponseId)}"${response.externalResponseId === duplicate.selectedResponseId ? ' selected' : ''}>${escapeHtml(formatDateTime(response.timestamp))}</option>`).join('')}</select>
		</div>`).join('');
	const invalidHint = !deliveryValidation.valid ? `${deliveryValidation.incomplete.length} Schüler noch nicht versandbereit.` : !brevoState.configured ? 'Brevo SMTP ist noch nicht eingerichtet.' : '';
	elements.body.innerHTML = `
		<div class="assessment-metrics">
			${metric('Selbstbeurteilungen', summary.selfCount, summary.studentCount)}
			${metric('Lehrerbeurteilungen', summary.teacherCount, summary.studentCount)}
			${metric('Versendet', summary.sentCount, summary.studentCount)}
			${metric('Synchronisation', summary.lastSyncStatus === 'failed' ? 'Fehler' : syncLabel, '')}
		</div>
		<div class="assessment-toolbar">
			<button class="btn-1${assessmentConfigured ? '' : ' disabled'}" id="assessment-sync-btn" title="${assessmentConfigured ? '' : 'Mitmachnoten-Konfiguration fehlt'}">Antworten aktualisieren</button>
			<button class="btn-2${invitation?.url ? '' : ' disabled'}" id="assessment-invitation-btn" ${invitation?.url ? '' : 'disabled'}>Einladungslink kopieren</button>
			<button class="btn-1" id="assessment-teacher-btn">Beurteilungen öffnen</button>
			<button class="btn-2${deliveryValidation.valid && brevoState.configured ? '' : ' disabled'}" id="assessment-send-btn" title="${escapeAttribute(invalidHint)}">Klassensatz versenden</button>
			${summary.failedCount ? '<button class="btn-2" id="assessment-retry-btn">Fehlgeschlagene erneut senden</button>' : ''}
		</div>
		${round.cloudStatus === 'open' ? '<button class="btn-2" id="assessment-close-round-btn">Runde schliessen</button>' : ''}
		${round.cloudStatus === 'closed' ? '<p class="assessment-warning">Runde geschlossen. Gespeicherte Antworten können weiterhin aktualisiert werden.</p>' : ''}
		${assessmentConfigured && ['closed','pending'].includes(round.cloudStatus) ? `<div class="assessment-reopen"><label>Neue Abgabefrist<input id="assessment-reopen-deadline" type="date" min="${deadlineBounds().min}" max="${deadlineBounds().max}" value="${zurichDate(Date.now()+7*DAY)}"></label><button class="btn-2" id="assessment-reopen-round-btn">${round.cloudStatus==='pending' ? 'Abgabefrist korrigieren und veröffentlichen' : 'Runde wieder öffnen'}</button><p>Der bestehende Umfragelink und bereits gespeicherte Antworten bleiben erhalten.</p></div>` : ''}
		${assessmentConfigured ? '<p class="assessment-auto-sync-note">Antworten werden automatisch alle 30 Sekunden synchronisiert.</p>' : ''}
		<details class="assessment-danger-zone"><summary>Weitere Aktionen</summary><button class="btn-danger" id="assessment-delete-round-btn">Runde lokal und online löschen</button></details>
		${!assessmentConfigured ? `<div class="assessment-warning">${!round.cloudStatus ? 'Diese Runde ist lokal archiviert. Gespeicherte Beurteilungen bleiben bearbeitbar. Für neue Online-Abgaben erstelle eine neue Runde.' : 'Die Online-Umfrage ist noch nicht verbunden. Importiere die Mitmachnoten-Konfiguration in den Einstellungen.'}</div>` : ''}
		${summary.duplicateCount ? `<section class="assessment-duplicates"><h3>Mehrfachantworten</h3><p class="assessment-warning">Standardmäßig wird die neueste Antwort verwendet. Hier kann eine frühere Abgabe gewählt werden.</p>${duplicates}</section>` : ''}
		${unmatched ? `<section class="assessment-unmatched"><h3>Manuelle Zuordnung erforderlich</h3>${unmatched}</section>` : ''}
		<section class="assessment-roster"><h3>Klassenübersicht</h3>${studentRows}</section>
		`;
	const sync = document.getElementById('assessment-sync-btn');
	if (sync && assessmentConfigured) sync.addEventListener('click', () => runRoundAction('assessment:sync-responses'));
	document.getElementById('assessment-teacher-btn').addEventListener('click', () => openTeacher(0));
 elements.body.querySelectorAll('[data-teacher-index]').forEach(button=>button.addEventListener('click',()=>openTeacher(Number(button.dataset.teacherIndex))));
	const invite = document.getElementById('assessment-invitation-btn');
	if (invite && invitation?.url) invite.addEventListener('click', async () => {
  try {await invoke('assessment:copy-invitation-link',{className:selectedClass,roundId:round.id});showNotice('Einladungslink kopiert.', 'success');}
  catch {showNotice('Der Einladungslink konnte nicht kopiert werden. Bitte erneut versuchen.','error');}
	});
	const send = document.getElementById('assessment-send-btn');
	if (!send.classList.contains('disabled')) send.addEventListener('click', () => openSendConfirm('pending'));
	const retry = document.getElementById('assessment-retry-btn');
	if (retry) retry.addEventListener('click', () => openSendConfirm('failed_only'));
	document.getElementById('assessment-delete-round-btn').addEventListener('click', openDeleteConfirm);
	document.getElementById('assessment-close-round-btn')?.addEventListener('click', () => runRoundAction('assessment:close-round'));
	document.getElementById('assessment-reopen-round-btn')?.addEventListener('click',()=>{
		const deadline=document.getElementById('assessment-reopen-deadline').value;
		try {validateDeadline(deadline,Date.now(),true);runRoundAction('assessment:reopen-round',{deadline});} catch(error){showNotice(error.message,'error');}
	});
	elements.body.querySelectorAll('[data-assign-response]').forEach(button => button.addEventListener('click', () => assignResponse(button.dataset.assignResponse)));
	elements.body.querySelectorAll('[data-select-response]').forEach(select => select.addEventListener('change', () => selectResponse(select.value)));
	if(background)elements.body.scrollTop=previousScroll;
}

async function runRoundAction(channel, extra = {}) {
	if(busy)return;
 const request=++viewRequest,className=selectedClass,roundId=roundView.round.id;
	try {
		setBusy(true, channel==='assessment:reopen-round' ? 'Runde wird geöffnet …' : 'Antworten werden aktualisiert …');
		const updated=await invoke(channel, {className,roundId,...extra});
  if(request!==viewRequest || className!==selectedClass || elements.view.classList.contains('update-hidden'))return;
  roundView=updated;
		renderRound();
		const result = roundView.summary.lastSyncResult;
		if(channel==='assessment:reopen-round')showNotice('Die Runde ist wieder geöffnet. Der bisherige Umfragelink bleibt gültig.','success');
		else if (channel==='assessment:sync-responses' && result) showNotice(result.added ? `${result.added} neue Antwort${result.added === 1 ? '' : 'en'} gespeichert.` : 'Keine neuen Antworten.', result.invalid ? 'warning' : 'success');
	} catch (error) {
		if(request===viewRequest)showNotice(error.message, 'error');
	} finally {
		if(request===viewRequest)setBusy(false);
	}
}

async function selectResponse(responseId) {
	try {
		roundView = await invoke('assessment:select-response', { className: selectedClass, roundId: roundView.round.id, responseId });
		renderRound();
		showNotice('Ausgewählte Antwort übernommen.', 'success');
	} catch (error) {
		showNotice(error.message, 'error');
	}
}

async function assignResponse(responseId) {
	const select = elements.body.querySelector(`[data-match-response="${cssEscape(responseId)}"]`);
	if (!select || !select.value) return showNotice('Bitte einen Schüler auswählen.', 'warning');
	try {
		roundView = await invoke('assessment:assign-response', { className: selectedClass, roundId: roundView.round.id, responseId, studentId: select.value });
		renderRound();
	} catch (error) {
		showNotice(error.message, 'error');
	}
}

function openTeacher(index) {
 captureTeacherDraft();
 if(!roundView?.students.length){renderEmpty('In dieser Klasse sind keine Schüler vorhanden.');return;}
	if (assessmentMotion.beginPageChange(() => { openTeacher(index); })) return;
	elements.body.scrollTop = 0;
	teacherIndex = Math.max(0, Math.min(index, roundView.students.length - 1));
	const student = roundView.students[teacherIndex];
	const assessment = roundView.round.assessments[student.id] || {};
 const key=teacherKey(student),draft=teacherDrafts.get(key);
	elements.heading.innerText = `${teacherIndex + 1} / ${roundView.students.length} · ${student.name}`;
	const criteriaRows = roundView.round.criteria.map(criterion => {
		const self = assessment.studentAnswers && assessment.studentAnswers[criterion.id];
		const teacher = draft ? (draft.ratings[criterion.id]?.rated ? draft.ratings[criterion.id].value : undefined) : assessment.teacherAnswers?.[criterion.id];
		const rated = Number.isInteger(teacher) && teacher >= 1 && teacher <= 5;
		const inputId = `teacher-rating-${criterion.id}`;
		const comment = draft?.comments[criterion.id] ?? assessment.teacherCriterionComments?.[criterion.id] ?? '';
		return `<div class="teacher-criterion-block">
			<div class="teacher-criterion">
				<strong>${escapeHtml(criterion.label)}</strong>
				<span class="teacher-self">${escapeHtml(formatCriterionValue(criterion, self))}</span>
				<div class="teacher-rating${rated ? ' is-rated' : ''}">
					<output for="${escapeAttribute(inputId)}">${rated ? escapeHtml(formatCriterionValue(criterion, teacher)) : 'Noch nicht bewertet'}</output>
					<input id="${escapeAttribute(inputId)}" type="range" min="1" max="5" step="1" value="${rated ? teacher : 3}" data-teacher-criterion="${escapeAttribute(criterion.id)}" data-rated="${rated}" aria-label="Lehrerbewertung: ${escapeAttribute(criterion.label)}" aria-valuetext="${rated ? escapeAttribute(formatCriterionValue(criterion, teacher)) : 'Noch nicht bewertet'}">
					<div class="teacher-rating-ticks" aria-hidden="true"><span>1</span><span>2</span><span>3</span><span>4</span><span>5</span></div>
				</div>
				<b class="teacher-difference" data-difference="${escapeAttribute(criterion.id)}">${self && rated ? signed(teacher - self) : '–'}</b>
			</div>
			<details class="teacher-criterion-comment"><summary>${comment ? 'Kommentar bearbeiten' : 'Kommentar hinzufügen'} <span>(optional)</span></summary><label>Kommentar zu diesem Kriterium<textarea data-teacher-comment="${escapeAttribute(criterion.id)}" aria-label="Kommentar: ${escapeAttribute(criterion.label)}" rows="3" maxlength="2000">${escapeHtml(comment)}</textarea></label></details>
		</div>`;
	}).join('');
	elements.body.innerHTML = `
		<section class="teacher-assessment">
			<p class="teacher-rating-hint">Bewertung von 1 bis 5: Regler anklicken oder ziehen. Kommentare kannst du bei Bedarf aufklappen.</p>
			<div class="teacher-grid-heading"><span>Kriterium</span><span>Schüler</span><span>Lehrer</span><span>Differenz</span></div>
			${criteriaRows}
			<div class="teacher-comments"><div><h3>Zusätzliche Qualität</h3><p>${escapeHtml(assessment.studentAdditionalQuality || 'Keine Angabe.')}</p></div><div><h3>Sonstige Bemerkungen</h3><p>${escapeHtml(assessment.studentComment || 'Keine Angabe.')}</p></div><label><h3>Kommentar Lehrer</h3><textarea id="teacher-comment" rows="5" maxlength="4000">${escapeHtml(draft?.comment ?? assessment.teacherComment ?? '')}</textarea></label></div>
			<div class="teacher-actions"><button class="btn-2${teacherIndex === 0 ? ' disabled' : ''}" id="teacher-previous-btn" ${teacherIndex===0 ? 'disabled' : ''}>Vorheriger</button><button class="btn-1" id="teacher-save-btn">Speichern</button><button class="btn-1" id="teacher-next-btn">Speichern & nächster</button></div>
		</section>`;
 renderedTeacherKey=key;
	elements.body.querySelectorAll('[data-teacher-criterion]').forEach(input => {
		const rate = () => {input.dataset.rated = 'true'; updateDifferences();};
		input.addEventListener('input', rate);
		input.addEventListener('change', rate);
		// A click on the middle value must also count as an intentional rating.
		input.addEventListener('pointerup', rate);
		input.addEventListener('keydown', event => {
			if (event.key === 'Enter' || event.key === ' ') {event.preventDefault(); rate();}
		});
	});
	window.repetiererKristall?.refresh();
	document.getElementById('teacher-previous-btn').addEventListener('click', () => openTeacher(teacherIndex - 1));
	document.getElementById('teacher-save-btn').addEventListener('click', () => saveTeacher(false));
	document.getElementById('teacher-next-btn').addEventListener('click', () => saveTeacher(true));
}

function formatCriterionValue(criterion, value) {
	const numeric = Number(value);
	if (!Number.isInteger(numeric) || numeric < 1 || numeric > 5) return '–';
	const label = Array.isArray(criterion.scaleLabels) ? criterion.scaleLabels[numeric - 1] : '';
	return label ? `${numeric} · ${label}` : String(numeric);
}

function updateDifferences() {
	const assessment = roundView.round.assessments[roundView.students[teacherIndex].id] || {};
	elements.body.querySelectorAll('[data-teacher-criterion]').forEach(input => {
		const id = input.dataset.teacherCriterion;
		const self = assessment.studentAnswers && assessment.studentAnswers[id];
		const criterion = roundView.round.criteria.find(item => item.id === id);
		const rated = input.dataset.rated === 'true';
		const rating = input.closest('.teacher-rating');
		rating.classList.toggle('is-rated', rated);
		rating.querySelector('output').textContent = rated ? formatCriterionValue(criterion, input.value) : 'Noch nicht bewertet';
		input.setAttribute('aria-valuetext', rated ? formatCriterionValue(criterion, input.value) : 'Noch nicht bewertet');
		elements.body.querySelector(`[data-difference="${cssEscape(id)}"]`).innerText = self && rated ? signed(Number(input.value) - self) : '–';
	});
}

async function saveTeacher(moveNext) {
 if(busy)return;
 const request=++viewRequest,className=selectedClass,roundId=roundView.round.id;
	const student = roundView.students[teacherIndex];
	const answers = {};
	const criterionComments = {};
	elements.body.querySelectorAll('[data-teacher-criterion]').forEach(input => { answers[input.dataset.teacherCriterion] = input.dataset.rated === 'true' ? Number(input.value) : 0; });
	elements.body.querySelectorAll('[data-teacher-comment]').forEach(field => { criterionComments[field.dataset.teacherComment] = field.value; });
	try {
  setBusy(true,'Lehrerbeurteilung wird gespeichert …');
		const updated = await invoke('assessment:save-teacher', { className: selectedClass, roundId: roundView.round.id, studentId: student.id, answers, criterionComments, comment: document.getElementById('teacher-comment').value });
  if(request!==viewRequest || className!==selectedClass || roundView?.round.id!==roundId || elements.view.classList.contains('update-hidden'))return;
  teacherDrafts.delete(teacherKey(student));renderedTeacherKey=null;
		roundView = updated;
  showNotice('Lehrerbeurteilung gespeichert.','success');
		if (moveNext && teacherIndex < roundView.students.length - 1) openTeacher(teacherIndex + 1);
		else if(moveNext)renderRound();
  else openTeacher(teacherIndex);
	} catch (error) {
		if(request===viewRequest)showNotice(error.message, 'error');
	} finally {if(request===viewRequest)setBusy(false);}
}

function openSendConfirm(mode) {
	pendingSendMode = mode;
	const count = mode === 'failed_only' ? roundView.summary.failedCount : roundView.summary.studentCount - roundView.summary.sentCount;
	elements.confirmText.innerText = `${count} E-Mail${count === 1 ? '' : 's'} werden über Brevo versendet. Bereits erfolgreich versendete Resultate werden nicht erneut verschickt.`;
	elements.confirmSend.innerText = `${count} E-Mail${count === 1 ? '' : 's'} senden`;
	assessmentMotion.openDialog(elements.confirmModal);
}

function closeConfirm() {
	assessmentMotion.closeDialog(elements.confirmModal);
}

function openDeleteConfirm() {
	const sentCount = roundView.summary.sentCount;
	elements.deleteText.innerText = `Die Beurteilungsrunde „${roundView.round.title}“ und alle lokal gespeicherten Antworten und Bewertungen dieser Runde werden gelöscht.`;
	elements.deleteSentWarning.innerText = sentCount ? `Achtung: Für ${sentCount} bereits versendete${sentCount === 1 ? 's Resultat' : ' Resultate'} gehen die lokalen Versanddaten verloren.` : '';
	elements.deleteSentWarning.classList.toggle('update-hidden', sentCount === 0);
	assessmentMotion.openDialog(elements.deleteModal);
}

function closeDeleteConfirm() {
	assessmentMotion.closeDialog(elements.deleteModal);
}

async function deleteRound() {
	if (elements.deleteConfirm.classList.contains('disabled')) return;
	try {
		elements.deleteConfirm.classList.add('disabled');
		overview = await invoke('assessment:delete-round', { className: selectedClass, roundId: roundView.round.id });
		closeDeleteConfirm();
		roundView = null;
		renderOverview();
		showNotice('Beurteilungsrunde gelöscht.', 'success');
	} catch (error) {
		showNotice(error.message, 'error');
	} finally {
		elements.deleteConfirm.classList.remove('disabled');
	}
}

async function sendRound(mode) {
	try {
		setBusy(true, 'Persönliche Auswertungen werden versendet …');
		const payload = await invoke('assessment:send-round', { className: selectedClass, roundId: roundView.round.id, mode });
		roundView = payload.round;
		renderRound();
		showNotice(`${payload.result.sent} erfolgreich versendet, ${payload.result.failed} fehlgeschlagen.`, payload.result.failed ? 'warning' : 'success');
	} catch (error) {
		showNotice(error.message, 'error');
	} finally {
		setBusy(false);
	}
}

async function loadIntegrationStatuses() {
	try {
		const [assessment, brevo] = await Promise.all([invoke('assessment:config-status'), invoke('assessment:brevo-status')]);
		brevoState = brevo;
		elements.assessmentConfigStatus.innerText = assessment.configured ? 'Umfrage: Eingerichtet' : assessment.invalid ? assessment.error : 'Umfrage: Nicht eingerichtet';
		elements.assessmentApiStatus.innerText = assessment.configured ? 'API: Noch nicht geprüft' : 'API: Nicht eingerichtet';
		toggle(elements.assessmentApiTest, assessment.configured);
		toggle(elements.assessmentConfigRemove, assessment.configured);
		elements.assessmentConfigImport.innerText = assessment.configured ? 'Konfiguration ersetzen' : 'Konfiguration importieren';
		elements.brevoStatus.innerText = brevo.configured ? `Eingerichtet · ${brevo.fromName} <${brevo.fromEmail}>` : brevo.invalid ? brevo.error : 'Nicht eingerichtet';
		toggle(elements.brevoTest, brevo.configured);
		toggle(elements.brevoRemove, brevo.configured);
		elements.brevoImport.innerText = brevo.configured ? 'Konfiguration ersetzen' : 'Konfiguration importieren';
	} catch (error) {
		setIntegrationMessage(elements.brevoStatus, error.message, 'error');
	}
}

async function refreshBrevoState() {
	try { brevoState = await invoke('assessment:brevo-status'); } catch (error) { brevoState = { configured: false }; }
}

async function integrationAction(channel, target, done) {
	try {
		const result = await invoke(channel);
		if (result && result.cancelled) return;
		await done(result);
	} catch (error) {
		setIntegrationMessage(target, error.message, 'error');
	}
}

async function invoke(channel, args) {
	const response = await ipcRenderer.invoke(channel, args || {});
	if (!response || !response.ok) {
		const error = new Error(response && response.error && response.error.message || 'Die Aktion konnte nicht abgeschlossen werden.');
		error.code = response && response.error && response.error.code;
		error.incomplete = response && response.error && response.error.incomplete;
		throw error;
	}
	return response.data;
}

function setBusy(active, message, replaceContent = false, layout = 'assessment') {
	busy = active;
	elements.body.classList.toggle('assessment-busy', active);
	elements.body.classList.toggle('assessment-loading-content', active && replaceContent);
	elements.body.setAttribute('aria-busy', String(active));
	elements.body.inert = active && !replaceContent;
	elements.classSelect.disabled = active;
	elements.back.disabled = active;
	elements.operationStatus.textContent = active && !replaceContent ? message || 'Wird geladen …' : '';
	elements.operationStatus.classList.toggle('update-hidden', !active || replaceContent);
	if (active && replaceContent) elements.body.replaceChildren(window.createRepetiererLoadingState(message, layout));
}

function renderEmpty(message) {
	elements.heading.innerText = 'Beurteilungsrunden';
	elements.back.classList.add('update-hidden');
	elements.body.innerHTML = `<p class="assessment-empty">${escapeHtml(message)}</p>`;
}

function renderError(message) {
	showNotice(message, 'error');
	elements.body.innerHTML = `<p class="assessment-empty">Die Ansicht konnte nicht geladen werden.</p><div><button class="btn-2" id="assessment-retry-btn">Erneut laden</button></div>`;
	document.getElementById('assessment-retry-btn').addEventListener('click', loadOverview);
}

function showNotice(message, type = 'info') {
	return notify[type](message,{title:'Mitmachnote'});
}

function metric(label, value, total) {
	return `<div class="assessment-metric"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}${total !== '' ? ` <small>/ ${escapeHtml(total)}</small>` : ''}</strong></div>`;
}

function setIntegrationMessage(element, message, type) {
	notify[type](message,{title:'Integrationen'});
}

function toggle(element, visible) {
	element.classList.toggle('update-hidden', !visible);
}

function statusIcon(value) { return value ? '<b class="status-done">✓</b>' : '<b class="status-open">○</b>'; }
function mailIcon(status) { return status === 'sent' ? '<b class="status-done">✓</b>' : status === 'failed' ? '<b class="status-error">!</b>' : '<b class="status-open">○</b>'; }
function signed(value) { return value > 0 ? `+${value}` : String(value); }
function formatDate(value) { const [year, month, day] = String(value).split('-'); return `${day}.${month}.${year}`; }
function formatDateTime(value) { const date = new Date(value); return Number.isNaN(date.getTime()) ? 'unbekannt' : date.toLocaleString('de-CH', { dateStyle: 'short', timeStyle: 'short' }); }
function escapeHtml(value) { return String(value === undefined || value === null ? '' : value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;'); }
function escapeAttribute(value) { return escapeHtml(value).replace(/`/g, '&#96;'); }
function cssEscape(value) { return window.CSS && CSS.escape ? CSS.escape(value) : String(value).replace(/["\\]/g, '\\$&'); }

module.exports = { invoke };
