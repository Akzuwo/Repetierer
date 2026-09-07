const { ipcRenderer, clipboard } = require('electron');

const elements = {
	view: document.getElementById('assessment-view'),
	body: document.getElementById('assessment-body'),
	heading: document.getElementById('assessment-heading'),
	classSelect: document.getElementById('assessment-class-select'),
	back: document.getElementById('assessment-back-btn'),
	classPanel: document.getElementById('class'),
	repetition: document.getElementById('repetition'),
	stats: document.getElementById('stats-view'),
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
let busy = false;

ipcRenderer.on('classes', (event, classes) => {
	loadedClasses = Array.isArray(classes) ? classes : [];
	renderClassOptions();
});

elements.assessmentNav.addEventListener('click', openAssessment);
elements.repetitionNav.addEventListener('click', openRepetition);
elements.classSelect.addEventListener('change', async () => {
	selectedClass = elements.classSelect.value;
	roundView = null;
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
elements.assessmentConfigImport.addEventListener('click', () => integrationAction('assessment:config-import', elements.assessmentConfigStatus, result => {
		loadIntegrationStatuses();
		if (result && result.deleteError) showNotice(result.deleteError, 'warning');
		if (result && result.formWarning) showNotice(result.formWarning, 'warning');
}));
elements.assessmentApiTest.addEventListener('click', async () => {
	try {
		await invoke('assessment:api-test');
		setIntegrationMessage(elements.assessmentApiStatus, 'API: Erreichbar', 'success');
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
	closeDrawer();
	elements.classPanel.classList.add('update-hidden');
	elements.repetition.classList.add('update-hidden');
	elements.stats.classList.add('update-hidden');
	elements.view.classList.remove('update-hidden');
	elements.title.innerText = 'Mitmachnote';
	const activeClass = document.querySelector('#class-list .class-selected');
	if (!selectedClass && activeClass) selectedClass = activeClass.innerText;
	renderClassOptions();
	if (selectedClass) await loadOverview();
	else renderEmpty(loadedClasses.length ? 'Bitte wähle eine Klasse.' : 'Bitte zuerst eine Repetierer-Datei laden.');
}

function openRepetition() {
	closeDrawer();
	elements.view.classList.add('update-hidden');
	elements.stats.classList.add('update-hidden');
	elements.classPanel.classList.remove('update-hidden');
	elements.repetition.classList.remove('update-hidden');
	elements.title.innerText = 'Repetierer';
}

function closeDrawer() {
	elements.drawer.classList.remove('drawer-open');
	elements.drawerScrim.classList.remove('scrim-visible');
	elements.drawerToggle.classList.remove('drawer-toggle-open');
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
	setBusy(true, 'Beurteilungsrunden werden geladen …');
	try {
		overview = await invoke('assessment:get-overview', { className: selectedClass });
		renderOverview();
	} catch (error) {
		renderError(error.message);
	} finally {
		setBusy(false);
	}
}

function renderOverview() {
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
				<label>Abgabefrist<input id="assessment-round-deadline" type="date"></label>
				<button class="btn-1" id="assessment-create-round-btn">Erstellen</button>
			</div>
		</section>
		<section class="assessment-round-list">
			<h3>Bestehende Runden</h3>
			${rounds || '<p class="assessment-empty">Für diese Klasse gibt es noch keine Beurteilungsrunde.</p>'}
		</section>`;
	document.getElementById('assessment-create-round-btn').addEventListener('click', createRound);
	elements.body.querySelectorAll('[data-round-id]').forEach(button => button.addEventListener('click', () => openRound(button.dataset.roundId)));
}

async function createRound() {
	const title = document.getElementById('assessment-round-title').value;
	const deadline = document.getElementById('assessment-round-deadline').value;
	try {
		setBusy(true, 'Runde wird erstellt …');
		overview = await invoke('assessment:create-round', { className: selectedClass, title, deadline });
		renderOverview();
	} catch (error) {
		showNotice(error.message, 'error');
	} finally {
		setBusy(false);
	}
}

async function openRound(roundId) {
	try {
		setBusy(true, 'Runde wird geladen …');
		roundView = await invoke('assessment:get-round', { className: selectedClass, roundId });
		await refreshBrevoState();
		renderRound();
	} catch (error) {
		renderError(error.message);
	} finally {
		setBusy(false);
	}
}

function renderRound() {
	const { round, summary, students, invitation, invitationError, assessmentConfigured, deliveryValidation } = roundView;
	elements.heading.innerText = `${selectedClass} · ${round.title}`;
	elements.back.classList.remove('update-hidden');
	const syncLabel = summary.lastSyncAt ? `Zuletzt ${formatDateTime(summary.lastSyncAt)}` : 'Noch nicht synchronisiert';
	const studentRows = students.map(student => {
		const assessment = round.assessments[student.id] || {};
		return `<div class="assessment-student-row"><strong>${escapeHtml(student.name)}</strong><span>Selbst ${statusIcon(assessment.studentAnswers)}</span><span>Lehrer ${statusIcon(assessment.teacherAnswers)}</span><span>Versand ${mailIcon(assessment.mailStatus)}</span></div>`;
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
			<button class="btn-2${invitation ? '' : ' disabled'}" id="assessment-invitation-btn">Einladung kopieren</button>
			<button class="btn-2" id="assessment-teacher-btn">Beurteilungen öffnen</button>
			<button class="btn-1${deliveryValidation.valid && brevoState.configured ? '' : ' disabled'}" id="assessment-send-btn" title="${escapeAttribute(invalidHint)}">Klassensatz versenden</button>
			${summary.failedCount ? '<button class="btn-2" id="assessment-retry-btn">Fehlgeschlagene erneut senden</button>' : ''}
			<button class="btn-2" id="assessment-delete-round-btn">Runde löschen</button>
		</div>
		${!assessmentConfigured ? '<div class="assessment-warning">Die Umfrage ist noch nicht eingerichtet. Importiere die Mitmachnoten-Konfiguration in den Einstellungen.</div>' : ''}
		${invitationError ? `<div class="assessment-warning">${escapeHtml(invitationError)}</div>` : ''}
		${summary.lastSyncResult && summary.lastSyncResult.invalid ? `<div class="assessment-warning">${summary.lastSyncResult.invalid} ungültige Antwort${summary.lastSyncResult.invalid === 1 ? '' : 'en'} wurde beim letzten Abruf ignoriert.</div>` : ''}
		${summary.lastSyncResult && summary.lastSyncResult.ignoredRound ? `<div class="assessment-warning">${summary.lastSyncResult.ignoredRound} Antwort${summary.lastSyncResult.ignoredRound === 1 ? '' : 'en'} aus einer anderen Runde wurde ignoriert.</div>` : ''}
		${summary.duplicateCount ? `<section class="assessment-duplicates"><h3>Mehrfachantworten</h3><p class="assessment-warning">Standardmäßig wird die neueste Antwort verwendet. Hier kann eine frühere Abgabe gewählt werden.</p>${duplicates}</section>` : ''}
		${unmatched ? `<section class="assessment-unmatched"><h3>Manuelle Zuordnung erforderlich</h3>${unmatched}</section>` : ''}
		<section class="assessment-roster"><h3>Klassenübersicht</h3>${studentRows}</section>
		<div id="assessment-invitation-area">${invitation ? renderInvitation(invitation, true) : ''}</div>`;
	const sync = document.getElementById('assessment-sync-btn');
	if (sync && assessmentConfigured) sync.addEventListener('click', () => runRoundAction('assessment:sync-responses'));
	document.getElementById('assessment-teacher-btn').addEventListener('click', () => openTeacher(0));
	const invite = document.getElementById('assessment-invitation-btn');
	if (invite && invitation) invite.addEventListener('click', () => {
		clipboard.writeText(invitation.complete);
		document.getElementById('assessment-invitation-area').classList.add('invitation-expanded');
		showNotice('Einladung kopiert.', 'success');
	});
	const send = document.getElementById('assessment-send-btn');
	if (!send.classList.contains('disabled')) send.addEventListener('click', () => openSendConfirm('pending'));
	const retry = document.getElementById('assessment-retry-btn');
	if (retry) retry.addEventListener('click', () => openSendConfirm('failed_only'));
	document.getElementById('assessment-delete-round-btn').addEventListener('click', openDeleteConfirm);
	elements.body.querySelectorAll('[data-assign-response]').forEach(button => button.addEventListener('click', () => assignResponse(button.dataset.assignResponse)));
	elements.body.querySelectorAll('[data-select-response]').forEach(select => select.addEventListener('change', () => selectResponse(select.value)));
	bindCopyButtons();
}

function renderInvitation(invitation, collapsed) {
	return `<section class="assessment-invitation${collapsed ? ' invitation-collapsed' : ''}"><h3>Einladung</h3><label>Betreff<textarea readonly rows="2">${escapeHtml(invitation.subject)}</textarea></label><label>Text<textarea readonly rows="8">${escapeHtml(invitation.text)}</textarea></label><div class="assessment-copy-actions"><button class="btn-2" data-copy="subject">Betreff kopieren</button><button class="btn-2" data-copy="text">Text kopieren</button><button class="btn-1" data-copy="complete">Gesamte Mail kopieren</button></div></section>`;
}

function bindCopyButtons() {
	if (!roundView.invitation) return;
	elements.body.querySelectorAll('[data-copy]').forEach(button => button.addEventListener('click', () => {
		clipboard.writeText(roundView.invitation[button.dataset.copy]);
		button.innerText = 'Kopiert';
		setTimeout(() => { button.innerText = button.dataset.copy === 'subject' ? 'Betreff kopieren' : button.dataset.copy === 'text' ? 'Text kopieren' : 'Gesamte Mail kopieren'; }, 1200);
	}));
}

async function runRoundAction(channel) {
	try {
		setBusy(true, 'Antworten werden aktualisiert …');
		roundView = await invoke(channel, { className: selectedClass, roundId: roundView.round.id });
		renderRound();
		const result = roundView.summary.lastSyncResult;
		if (result) showNotice(result.added ? `${result.added} neue Antwort${result.added === 1 ? '' : 'en'} gespeichert.` : 'Keine neuen Antworten.', result.invalid ? 'warning' : 'success');
	} catch (error) {
		showNotice(error.message, 'error');
	} finally {
		setBusy(false);
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
	teacherIndex = Math.max(0, Math.min(index, roundView.students.length - 1));
	const student = roundView.students[teacherIndex];
	const assessment = roundView.round.assessments[student.id] || {};
	elements.heading.innerText = `${teacherIndex + 1} / ${roundView.students.length} · ${student.name}`;
	const criteriaRows = roundView.round.criteria.map(criterion => {
		const self = assessment.studentAnswers && assessment.studentAnswers[criterion.id];
		const teacher = assessment.teacherAnswers && assessment.teacherAnswers[criterion.id];
		return `<div class="teacher-criterion"><strong>${escapeHtml(criterion.label)}</strong><span>${escapeHtml(formatCriterionValue(criterion, self))}</span><select data-teacher-criterion="${escapeAttribute(criterion.id)}"><option value="">–</option>${[1,2,3,4,5].map(value => `<option value="${value}"${teacher === value ? ' selected' : ''}>${escapeHtml(formatCriterionValue(criterion, value))}</option>`).join('')}</select><b data-difference="${escapeAttribute(criterion.id)}">${self && teacher ? signed(teacher - self) : '–'}</b></div>`;
	}).join('');
	elements.body.innerHTML = `
		<section class="teacher-assessment">
			<div class="teacher-grid-heading"><span>Kriterium</span><span>Schüler</span><span>Lehrer</span><span>Differenz</span></div>
			${criteriaRows}
			<div class="teacher-comments"><div><h3>Zusätzliche Qualität</h3><p>${escapeHtml(assessment.studentAdditionalQuality || 'Keine Angabe.')}</p></div><div><h3>Sonstige Bemerkungen</h3><p>${escapeHtml(assessment.studentComment || 'Keine Angabe.')}</p></div><label><h3>Kommentar Lehrer</h3><textarea id="teacher-comment" rows="5" maxlength="4000">${escapeHtml(assessment.teacherComment || '')}</textarea></label></div>
			<div class="teacher-actions"><button class="btn-2${teacherIndex === 0 ? ' disabled' : ''}" id="teacher-previous-btn">Vorheriger</button><button class="btn-1" id="teacher-save-btn">Speichern</button><button class="btn-1" id="teacher-next-btn">Speichern & nächster</button></div>
		</section>`;
	elements.body.querySelectorAll('[data-teacher-criterion]').forEach(select => select.addEventListener('change', updateDifferences));
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
	elements.body.querySelectorAll('[data-teacher-criterion]').forEach(select => {
		const self = assessment.studentAnswers && assessment.studentAnswers[select.dataset.teacherCriterion];
		const difference = elements.body.querySelector(`[data-difference="${cssEscape(select.dataset.teacherCriterion)}"]`);
		difference.innerText = self && select.value ? signed(Number(select.value) - self) : '–';
	});
}

async function saveTeacher(moveNext) {
	const student = roundView.students[teacherIndex];
	const answers = {};
	elements.body.querySelectorAll('[data-teacher-criterion]').forEach(select => { answers[select.dataset.teacherCriterion] = Number(select.value); });
	try {
		const updated = await invoke('assessment:save-teacher', { className: selectedClass, roundId: roundView.round.id, studentId: student.id, answers, comment: document.getElementById('teacher-comment').value });
		roundView = updated;
		if (moveNext && teacherIndex < roundView.students.length - 1) openTeacher(teacherIndex + 1);
		else openTeacher(teacherIndex);
	} catch (error) {
		showNotice(error.message, 'error');
	}
}

function openSendConfirm(mode) {
	pendingSendMode = mode;
	const count = mode === 'failed_only' ? roundView.summary.failedCount : roundView.summary.studentCount - roundView.summary.sentCount;
	elements.confirmText.innerText = `${count} E-Mail${count === 1 ? '' : 's'} werden über Brevo versendet. Bereits erfolgreich versendete Resultate werden nicht erneut verschickt.`;
	elements.confirmSend.innerText = `${count} E-Mail${count === 1 ? '' : 's'} senden`;
	elements.confirmModal.style.display = 'block';
	elements.confirmModal.classList.add('modal-open');
}

function closeConfirm() {
	elements.confirmModal.classList.remove('modal-open');
	elements.confirmModal.style.display = 'none';
}

function openDeleteConfirm() {
	const sentCount = roundView.summary.sentCount;
	elements.deleteText.innerText = `Die Beurteilungsrunde „${roundView.round.title}“ und alle lokal gespeicherten Antworten und Bewertungen dieser Runde werden gelöscht.`;
	elements.deleteSentWarning.innerText = sentCount ? `Achtung: Für ${sentCount} bereits versendete${sentCount === 1 ? 's Resultat' : ' Resultate'} gehen die lokalen Versanddaten verloren.` : '';
	elements.deleteSentWarning.classList.toggle('update-hidden', sentCount === 0);
	elements.deleteModal.style.display = 'block';
	elements.deleteModal.classList.add('modal-open');
}

function closeDeleteConfirm() {
	elements.deleteModal.classList.remove('modal-open');
	elements.deleteModal.style.display = 'none';
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

function setBusy(active, message) {
	busy = active;
	elements.body.classList.toggle('assessment-busy', active);
	elements.body.dataset.busyLabel = message || 'Wird geladen …';
	if (active && !elements.body.innerHTML.trim()) elements.body.innerHTML = `<div class="assessment-loading">${escapeHtml(message || 'Wird geladen …')}</div>`;
}

function renderEmpty(message) {
	elements.heading.innerText = 'Beurteilungsrunden';
	elements.back.classList.add('update-hidden');
	elements.body.innerHTML = `<p class="assessment-empty">${escapeHtml(message)}</p>`;
}

function renderError(message) {
	elements.body.innerHTML = `<div class="assessment-notice assessment-notice-error">${escapeHtml(message)}</div>`;
}

function showNotice(message, type) {
	let notice = elements.body.querySelector('.assessment-floating-notice');
	if (!notice) {
		notice = document.createElement('div');
		notice.className = 'assessment-floating-notice';
		elements.body.prepend(notice);
	}
	notice.className = `assessment-floating-notice assessment-notice-${type}`;
	notice.innerText = message;
	setTimeout(() => notice.remove(), 5500);
}

function metric(label, value, total) {
	return `<div class="assessment-metric"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}${total !== '' ? ` <small>/ ${escapeHtml(total)}</small>` : ''}</strong></div>`;
}

function setIntegrationMessage(element, message, type) {
	element.innerText = message;
	element.dataset.status = type;
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
