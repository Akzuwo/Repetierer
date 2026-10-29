const {ipcRenderer, clipboard} = require('electron');
const uiMotion = require('./uiMotion.js');
let currentRepetitionToken = null;
function startRepetitionTiming() { if(currentRepetitionToken && (state === 3 || state === 4)) ipcRenderer.send('repetition-timing-start', currentRepetitionToken); }
let state = 0; // 0: select file, 1: select class, 2: click start, 3: set grade

// elements
const _container = document.getElementById('container');
const _minimize = document.getElementById('minimize');
const _quit = document.getElementById('quit');
const _drawerToggle = document.getElementById('drawer-toggle');
const _drawer = document.getElementById('drawer');
const _drawerScrim = document.getElementById('drawer-scrim');
const _drawerStatus = document.getElementById('drawer-status');
const _drawerTitle = document.getElementById('drawer-title');
const _drawerBackBtn = document.getElementById('drawer-back-btn');
const _drawerHelpBtn = document.getElementById('drawer-help-btn');
const _drawerMainView = document.getElementById('drawer-main-view');
const _drawerDataView = document.getElementById('drawer-data-view');
const _drawerProtocolView = document.getElementById('drawer-protocol-view');
const _drawerDataViewBtn = document.getElementById('drawer-data-view-btn');
const _drawerProtocolViewBtn = document.getElementById('drawer-protocol-view-btn');
const _file = document.getElementById('drawer-file-btn');
const _filePath = document.getElementById('file-path');
const _backupBtn = document.getElementById('backup-btn');
const _logBtn = document.getElementById('log-btn');
const _exportProtocolBtn = document.getElementById('export-protocol-btn');
const _problemReportBtn = document.getElementById('problem-report-btn');
const {problemReportsEnabled} = require('../src/featureFlags');
_problemReportBtn.hidden = !problemReportsEnabled;
_problemReportBtn.disabled = !problemReportsEnabled;
_problemReportBtn.classList.toggle('update-hidden', !problemReportsEnabled);
const _statsViewBtn = document.getElementById('stats-view-btn');
const _importClassListBtn = document.getElementById('import-class-list-btn');
const _undoLastBtn = document.getElementById('undo-last-btn');
const _redoLastBtn = document.getElementById('redo-last-btn');
const _reloadExcelBtn = document.getElementById('reload-excel-btn');
const _excelEditorBtn = document.getElementById('excel-editor-btn');
const _debugActions = document.getElementById('debug-actions');
const _debugTestAlertBtn = document.getElementById('debug-test-alert-btn');
const _debugFakeUpdateBtn = document.getElementById('debug-fake-update-btn');
const _settingsBtn = document.getElementById('settings-btn');
const _classHelpBtn = document.getElementById('class-help-btn');
const _repetitionHelpBtn = document.getElementById('repetition-help-btn');
const _classes = document.getElementById('class-list').children.item(0);
const _start = document.getElementById('start-btn');
const _probabilitiesBtn = document.getElementById('probabilities-btn');
const _flushExcelBtn = document.getElementById('flush-excel-btn');
const _repetition = document.getElementById('repetition');
const _classPanel = document.getElementById('class');
const _statsView = document.getElementById('stats-view');
const _statsBackBtn = document.getElementById('stats-back-btn');
const _statsEmpty = document.getElementById('stats-empty');
const _statsList = document.getElementById('stats-list');
const _title = document.getElementById('title');
const _name = document.getElementById('name');
const _selectionEffect = document.getElementById('selection-effect');
const _spinwheelTrack = document.getElementById('spinwheel-track');
const _winnerName = document.getElementById('winner-name');
const _confettiStage = document.getElementById('confetti-stage');
const _completionEffect = document.getElementById('completion-effect');
const saveReceipt = require('./visualFeedback').createSaveReceipt(_completionEffect);
const _label = document.getElementById('grade').children.item(0);
const _grade = document.getElementById('grade').children.item(1);
const _cancel = document.getElementById('res').children.item(1);
const _ok = document.getElementById('res').children.item(2);
const _joker = document.getElementById('res').children.item(3);
const _manualSelectBtn = document.getElementById('manual-select-btn');
const _absencesBtn = document.getElementById('absences-btn');
const _personModal = document.getElementById('person-modal');
const _personList = document.getElementById('person-list');
const _closeModal = document.getElementById('close-modal');
const _randomSelectBtn = document.getElementById('random-select-btn');
const _absenceModal = document.getElementById('absence-modal');
const _absenceList = document.getElementById('absence-list');
const _saveAbsencesBtn = document.getElementById('save-absences-btn');
const _closeAbsenceModal = document.getElementById('close-absence-modal');
const _className = document.getElementById('class-name');
const _sessionStartHelpBtn = document.getElementById('session-start-help-btn');
const _sessionEmptyView = document.getElementById('session-empty-view');
const _sessionHeader = document.getElementById('session-header');
const _sessionClassName = document.getElementById('session-class-name');
const _sessionParticipationTab = document.getElementById('session-participation-tab');
const _sessionRepetitionTab = document.getElementById('session-repetition-tab');
const _endSessionBtn = document.getElementById('end-session-btn');
const _participationView = document.getElementById('participation-view');
const _participationStudentList = document.getElementById('participation-student-list');
const _participationSummaryBtn = document.getElementById('participation-summary-btn');
const _participationSummary = document.getElementById('participation-summary');
const _participationSummaryList = document.getElementById('participation-summary-list');
const _participationSummaryCloseBtn = document.getElementById('participation-summary-close-btn');
const _sessionAttendanceView = document.getElementById('session-attendance-view');
const _sessionStartClass = document.getElementById('session-start-class');
const _sessionAttendanceList = document.getElementById('session-attendance-list');
const _sessionPresentPreview = document.getElementById('session-present-preview');
const _sessionAbsentPreview = document.getElementById('session-absent-preview');
const _confirmSessionStartBtn = document.getElementById('confirm-session-start-btn');
const _cancelSessionStartBtn = document.getElementById('cancel-session-start-btn');
const _sessionEndModal = document.getElementById('session-end-modal');
const _confirmSessionEndBtn = document.getElementById('confirm-session-end-btn');
const _cancelSessionEndBtn = document.getElementById('cancel-session-end-btn');
const _participationModal = document.getElementById('participation-modal');
const _participationStudentName = document.getElementById('participation-student-name');
const _participationScoreList = document.getElementById('participation-score-list');
const _participationCancelBtn = document.getElementById('participation-cancel-btn');
const _backupModal = document.getElementById('backup-modal');
const _backupList = document.getElementById('backup-list');
const _backupLocation = document.getElementById('backup-location');
const _closeBackupModal = document.getElementById('close-backup-modal');
const _logModal = document.getElementById('log-modal');
const _logList = document.getElementById('log-list');
const _logLocation = document.getElementById('log-location');
const _copyLogBtn = document.getElementById('copy-log-btn');
const _closeLogModal = document.getElementById('close-log-modal');
const _problemReportModal = document.getElementById('problem-report-modal');
const _problemReportComment = document.getElementById('problem-report-comment');
const _problemReportSummary = document.getElementById('problem-report-summary');
const _problemReportDetails = document.getElementById('problem-report-details');
const _problemReportStatus = document.getElementById('problem-report-status');
const _sendProblemReportBtn = document.getElementById('send-problem-report-btn');
const _closeProblemReportBtn = document.getElementById('close-problem-report-btn');
const _toggleProblemReportDetailsBtn = document.getElementById('toggle-problem-report-details-btn');
const _reloadConflictModal = document.getElementById('reload-conflict-modal');
const _conflictFlushBtn = document.getElementById('conflict-flush-btn');
const _conflictForceReloadBtn = document.getElementById('conflict-force-reload-btn');
const _conflictCancelBtn = document.getElementById('conflict-cancel-btn');
const _classImportModal = document.getElementById('class-import-modal');
const _classImportDetail = document.getElementById('class-import-detail');
const _classImportName = document.getElementById('class-import-name');
const _classImportPreview = document.getElementById('class-import-preview');
const _confirmClassImportBtn = document.getElementById('confirm-class-import-btn');
const _cancelClassImportBtn = document.getElementById('cancel-class-import-btn');
const _teacherHelpModal = document.getElementById('teacher-help-modal');
const _teacherHelpTitle = document.getElementById('teacher-help-title');
const _teacherHelpBody = document.getElementById('teacher-help-body');
const _closeTeacherHelpBtn = document.getElementById('close-teacher-help-btn');
const _closeTeacherHelpBottomBtn = document.getElementById('close-teacher-help-bottom-btn');
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
const _settingsView = document.getElementById('settings-view');
const _settingsHelpBtn = document.getElementById('settings-help-btn');
const _settingsHelpModal = document.getElementById('settings-help-modal');
const _closeSettingsHelpBtn = document.getElementById('close-settings-help-btn');
const _closeSettingsHelpBottomBtn = document.getElementById('close-settings-help-bottom-btn');
const _settingsLocation = document.getElementById('settings-location');

const _updateNewsModal = document.getElementById('update-news-modal');
const _updateNewsTitle = document.getElementById('update-news-title');
const _updateNewsBody = document.getElementById('update-news-body');
const _closeUpdateNewsBtn = document.getElementById('close-update-news-btn');
const _extraJokerSetting = document.getElementById('extra-joker-setting');
const _extraJokerMigrationBtn = document.getElementById('extra-joker-migration-btn');
const _extraJokerMigrationStatus = document.getElementById('extra-joker-migration-status');
const _probabilityFactorSetting = document.getElementById('probability-factor-setting');
const _boostNeverSelectedSetting = document.getElementById('boost-never-selected-setting');
const _neverSelectedBoostFactorSetting = document.getElementById('never-selected-boost-factor-setting');
const _visualEffectsSetting = document.getElementById('visual-effects-setting');
const _wiggersRuleSetting = document.getElementById('wiggers-rule-setting');
const _wiggersRuleDurationSetting = document.getElementById('wiggers-rule-duration-setting');
const settingsRanges = [_probabilityFactorSetting, _neverSelectedBoostFactorSetting, _wiggersRuleDurationSetting];
const factorFormat = new Intl.NumberFormat('de-CH', {minimumFractionDigits:1, maximumFractionDigits:1});
function updateSettingsRangeValue(input) {
	const value = input.valueAsNumber;
	const minutes = input === _wiggersRuleDurationSetting;
	const formatted = minutes ? `${value} Min.` : `${factorFormat.format(value)}×`;
	document.getElementById(`${input.id}-value`).value = formatted;
	input.setAttribute('aria-valuetext', minutes ? `${value} ${value === 1 ? 'Minute' : 'Minuten'}` : `${factorFormat.format(value)}-fach`);
}
for (const input of settingsRanges) {
	input.addEventListener('input', () => updateSettingsRangeValue(input));
	input.addEventListener('change', () => updateSettingsRangeValue(input));
	updateSettingsRangeValue(input);
}
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
let absentPersonIds = [];
let activeTeachingSession = null;
let aiSelectedStudentId = null;
let pendingSessionSetupClass = '';
let sessionSetup = null;
let sessionSetupAbsentIds = [];
let selectedParticipationStudent = null;
let activeSessionArea = 'participation';
let excelEditorPersons = [];
let pendingExcelCount = 0;
let pendingClassImport = null;
let extraJokerMigrationStatus = { hasFile: false, completed: false };
let _currentClass;
let _nameSize;
let updatePrimaryAction = requestUpdateDownload;
let drawerView = 'main';
let updateNewsVersion = '';
let problemReportPreview = null;
let problemReportSending = false;
let statsResultsTimer = null;
let statsScrollFrame = null;
let statsScrollIdleTimer = null;
let statsLastScrollTop = 0;
let statsCardObserver = null;
let visualEffectsEnabled = false;
let selectionEffectRun = 0;
let completionEffectRun = 0;
const selectionEffectTimers = [];
const completionEffectTimers = [];
const prefersReducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)');
prefersReducedMotion.addEventListener('change', () => {
	if (prefersReducedMotion.matches) { clearSelectionEffect(); startRepetitionTiming(); }
});

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
document.getElementById('session-open-file-btn').addEventListener('click', () => _file.click());
document.getElementById('session-import-btn').addEventListener('click', () => _importClassListBtn.click());
document.getElementById('session-data-btn').addEventListener('click', () => { if (!_drawer.classList.contains('drawer-open')) toggleDrawer(); setDrawerView('data'); });

_drawerToggle.addEventListener('click', () => {
	toggleDrawer();
});

_drawerScrim.addEventListener('click', () => {
	closeDrawer();
});

_drawerBackBtn.addEventListener('click', () => {
	setDrawerView('main');
});

_drawerDataViewBtn.addEventListener('click', () => {
	setDrawerView('data');
});

_drawerProtocolViewBtn.addEventListener('click', () => {
	setDrawerView('protocol');
});

_drawerHelpBtn.addEventListener('click', () => {
	openTeacherHelp(`drawer-${drawerView}`);
});

// backup
_backupBtn.addEventListener('click', () => {
	ipcRenderer.send('get-backup');
});

_logBtn.addEventListener('click', () => {
	ipcRenderer.send('get-logs');
});

_exportProtocolBtn.addEventListener('click', () => {
	ipcRenderer.send('export-session-protocol', 'both');
});

_problemReportBtn.addEventListener('click', () => {
	openProblemReportModal();
});

_statsViewBtn.addEventListener('click', () => {
	openStatsView();
});

_statsBackBtn.addEventListener('click', () => {
	closeStatsView();
});

_statsList.addEventListener('scroll', handleStatsScroll, { passive: true });

_importClassListBtn.addEventListener('click', () => {
	ipcRenderer.send('import-class-list');
});

_undoLastBtn.addEventListener('click', () => {
	ipcRenderer.send('undo-last-action');
});

_redoLastBtn.addEventListener('click', () => {
	ipcRenderer.send('redo-last-action');
});

_closeBackupModal.addEventListener('click', () => {
	closeModal(_backupModal);
});

_backupModal.addEventListener('click', (e) => {
	if (e.target === _backupModal) {
		closeModal(_backupModal);
	}
});

_closeLogModal.addEventListener('click', () => {
	closeModal(_logModal);
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
		closeModal(_logModal);
	}
});

_closeProblemReportBtn.addEventListener('click', () => {
	closeProblemReportModal();
});

_problemReportModal.addEventListener('click', (e) => {
	if (e.target === _problemReportModal) {
		closeProblemReportModal();
	}
});

_problemReportComment.addEventListener('input', () => {
	renderProblemReportPreview();
});

_toggleProblemReportDetailsBtn.addEventListener('click', () => {
	const isHidden = _problemReportDetails.classList.contains('update-hidden');
	_problemReportDetails.classList.toggle('update-hidden', !isHidden);
	_toggleProblemReportDetailsBtn.innerText = isHidden ? 'Details ausblenden' : 'Details anzeigen';
});

_sendProblemReportBtn.addEventListener('click', () => {
	sendProblemReport();
});

// probabilities
_probabilitiesBtn.addEventListener('click', () => {
	ipcRenderer.send('get-probabilities');
});

_flushExcelBtn.addEventListener('click', () => {
	ipcRenderer.send('flush-pending-excel');
});

_reloadExcelBtn.addEventListener('click', () => {
	if (pendingExcelCount > 0) return;
	ipcRenderer.send('reload-excel');
});

_excelEditorBtn.addEventListener('click', () => {
	ipcRenderer.send('get-editor-persons');
});

_conflictFlushBtn.addEventListener('click', () => {
	closeModal(_reloadConflictModal);
	ipcRenderer.send('flush-pending-excel');
});

_conflictForceReloadBtn.addEventListener('click', () => {
	closeModal(_reloadConflictModal);
	ipcRenderer.send('reload-excel-force');
});

_conflictCancelBtn.addEventListener('click', () => {
	closeModal(_reloadConflictModal);
});

_confirmClassImportBtn.addEventListener('click', () => {
	if (!pendingClassImport) return;
	const className = _classImportName.value.trim();
	if (!className) {
		error(_confirmClassImportBtn);
		return;
	}
	closeModal(_classImportModal);
	ipcRenderer.send('confirm-class-list-import', {
		sourcePath: pendingClassImport.sourcePath,
		className: className
	});
});

_cancelClassImportBtn.addEventListener('click', () => {
	pendingClassImport = null;
	closeModal(_classImportModal);
});

_classImportModal.addEventListener('click', (e) => {
	if (e.target === _classImportModal) {
		pendingClassImport = null;
		closeModal(_classImportModal);
	}
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
	closeModal(_probabilitiesModal);
});

_probabilitiesModal.addEventListener('click', (e) => {
	if (e.target === _probabilitiesModal) {
		closeModal(_probabilitiesModal);
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
	if (!_settingsView.classList.contains('update-hidden')) { document.getElementById('settings-heading').focus(); return; }
	ipcRenderer.send('get-settings');
});

_classHelpBtn.addEventListener('click', () => {
	openTeacherHelp('class');
});

_sessionStartHelpBtn.addEventListener('click', () => {
	openTeacherHelp('class');
});

_repetitionHelpBtn.addEventListener('click', () => {
	openTeacherHelp('repetition');
});

_closeTeacherHelpBtn.addEventListener('click', () => {
	closeModal(_teacherHelpModal);
});

_closeTeacherHelpBottomBtn.addEventListener('click', () => {
	closeModal(_teacherHelpModal);
});

_teacherHelpModal.addEventListener('click', (e) => {
	if (e.target === _teacherHelpModal) {
		closeModal(_teacherHelpModal);
	}
});

_saveSettingsBtn.addEventListener('click', () => {
	const probabilityDecreaseFactor = parseFloat(_probabilityFactorSetting.value.replace(',', '.'));
	const neverSelectedBoostFactor = parseFloat(_neverSelectedBoostFactorSetting.value.replace(',', '.'));
	const wiggersRuleDurationMinutes = parseInt(_wiggersRuleDurationSetting.value, 10);
	if (!probabilityDecreaseFactor || probabilityDecreaseFactor < 1.1 || probabilityDecreaseFactor > 10 ||
		!neverSelectedBoostFactor || neverSelectedBoostFactor < 1.1 || neverSelectedBoostFactor > 10 ||
		!wiggersRuleDurationMinutes || wiggersRuleDurationMinutes < 1 || wiggersRuleDurationMinutes > 480) {
		error(_saveSettingsBtn);
		return;
	}

	ipcRenderer.send('save-settings', {
		theme: 'midnight',
		extraJokerAfterThreeGrades: _extraJokerSetting.checked,
		probabilityDecreaseFactor: probabilityDecreaseFactor,
		boostNeverSelected: _boostNeverSelectedSetting.checked,
		neverSelectedBoostFactor: neverSelectedBoostFactor,
		visualEffectsEnabled: _visualEffectsSetting.checked,
		wiggersRuleEnabled: _wiggersRuleSetting.checked,
		wiggersRuleDurationMinutes: wiggersRuleDurationMinutes
	});
	closeModal(_settingsView);
});

_extraJokerMigrationBtn.addEventListener('click', () => {
	if (!extraJokerMigrationStatus || !extraJokerMigrationStatus.hasFile) {
		showExtraJokerMigrationMessage('Extra-Joker-Migration', 'Bitte zuerst eine Datei laden.');
		return;
	}
	if (extraJokerMigrationStatus.completed) {
		showExtraJokerMigrationMessage('Extra-Joker-Migration', 'Diese Migration wurde für diese Datei bereits durchgeführt.');
		return;
	}

	disableElement(_extraJokerMigrationBtn);
	_extraJokerMigrationBtn.disabled = true;
	ipcRenderer.send('run-extra-joker-migration');
});

_closeSettingsModal.addEventListener('click', () => {
	closeModal(_settingsView);
});
document.getElementById('settings-back-btn').addEventListener('click', () => closeSettingsPage());

_closeUpdateNewsBtn.addEventListener('click', () => {
	closeUpdateNewsModal();
});

_updateNewsModal.addEventListener('click', (e) => {
	if (e.target === _updateNewsModal) {
		closeUpdateNewsModal();
	}
});

_settingsHelpBtn.addEventListener('click', () => {
	openModal(_settingsHelpModal);
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

_wiggersRuleSetting.addEventListener('change', () => {
	updateWiggersRuleFields();
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
		selectClassForNavigation(e.innerText);
	});
}

function selectClassForNavigation(className) {
	const button = Array.from(_classes.querySelectorAll('button')).find(item => item.innerText === className);
	if (!button) return false;
	setActiveClassButton(button);
	if (!activeTeachingSession) pendingSessionSetupClass = className;
	ipcRenderer.send('class', className);
	return true;
}

// start
_start.addEventListener('click', () => {
	ipcRenderer.send('start');
});

// cancel
_cancel.addEventListener('click', () => {
	ipcRenderer.send('repetition-timing-cancel'); currentRepetitionToken = null;
	state = 2;
	updateState();
});

// ok
_ok.addEventListener('click', () => {
	let v = parseGradeInput(_grade.value);
	if (isValidGrade(v)) {
		runCompletionEffect();
		ipcRenderer.send('ok', v);
	} else {
		error(_ok);
		showInvalidGradeAlert();
	}

});
_joker.addEventListener('click', () => {
	runCompletionEffect();
	ipcRenderer.send('joker');

});

// manual select button
_manualSelectBtn.addEventListener('click', () => {
	selectedPersonIds = [];
	ipcRenderer.send('get-persons');
});

_absencesBtn.addEventListener('click', () => {
	ipcRenderer.send('get-absence-persons');
});

_cancelSessionStartBtn.addEventListener('click', uiMotion.pageHandler(() => {
 sessionSetup=null;sessionSetupAbsentIds=[];
 _sessionAttendanceView.classList.add('update-hidden');
 renderTeachingSessionView();
 const selected=_classes.querySelector('.class-selected');if(selected)selected.focus();
}));
_confirmSessionStartBtn.addEventListener('click', () => {
	if (!sessionSetup) return;
	disableElement(_confirmSessionStartBtn);
 _confirmSessionStartBtn.disabled=true;
 _cancelSessionStartBtn.disabled=true;
 _sessionAttendanceList.inert=true;
	ipcRenderer.send('start-teaching-session', { className: sessionSetup.className, absentIds: sessionSetupAbsentIds });
});

_sessionParticipationTab.addEventListener('click', uiMotion.pageHandler(() => showSessionArea('participation')));
_sessionRepetitionTab.addEventListener('click', uiMotion.pageHandler(() => showSessionArea('repetition')));
_endSessionBtn.addEventListener('click', () => {
	if (activeTeachingSession) openModal(_sessionEndModal);
});
_cancelSessionEndBtn.addEventListener('click', () => closeModal(_sessionEndModal));
_sessionEndModal.addEventListener('click', event => {
	if (event.target === _sessionEndModal) closeModal(_sessionEndModal);
});
_confirmSessionEndBtn.addEventListener('click', () => {
	if (!activeTeachingSession) return;
	disableElement(_confirmSessionEndBtn);
	ipcRenderer.send('end-teaching-session', activeTeachingSession.id);
});

_participationSummaryBtn.addEventListener('click', () => {
	if (!activeTeachingSession) return;
	ipcRenderer.send('get-participation-summary', activeTeachingSession.className);
});
_participationSummaryCloseBtn.addEventListener('click', closeParticipationSummary);
_participationCancelBtn.addEventListener('click', () => closeModal(_participationModal));
_participationModal.addEventListener('click', event => {
	if (event.target === _participationModal) closeModal(_participationModal);
});

// close modal
_closeModal.addEventListener('click', () => {
	closeModal(_personModal);
	selectedPersonIds = [];
});

// close modal when clicking outside
_personModal.addEventListener('click', (e) => {
	if (e.target === _personModal) {
		closeModal(_personModal);
		selectedPersonIds = [];
	}
});

_saveAbsencesBtn.addEventListener('click', () => {
	ipcRenderer.send('save-absences', absentPersonIds);
});

_closeAbsenceModal.addEventListener('click', () => {
	closeModal(_absenceModal);
});

_absenceModal.addEventListener('click', (e) => {
	if (e.target === _absenceModal) {
		closeModal(_absenceModal);
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
	closeModal(_personModal);
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
 sessionSetup=null;sessionSetupAbsentIds=[];_sessionAttendanceView.classList.add('update-hidden');
	if (args) {
		state = 1;
		if (filePath) _filePath.innerText = filePath;
		_drawerStatus.innerText = 'Datei geladen. Wähle jetzt eine Klasse.';
		updateState();

		_classes.innerHTML = '';
		_currentClass = null;
		enableElement(_statsViewBtn);
		for (let i = 0; i < args.length; i++) {
			let x = document.createElement('button')
			x.className = 'btn-2';
			x.innerText = args[i];
			classEvent(x);
			_classes.appendChild(x);
		}
		if (activeTeachingSession) {
			const activeButton = Array.from(_classes.querySelectorAll('button')).find(button => button.innerText === activeTeachingSession.className);
			if (activeButton) {
				setActiveClassButton(activeButton);
				ipcRenderer.send('class', activeTeachingSession.className);
			}
		}
		renderSessionStartState(args.length ? '' : 'In dieser Datei wurden keine Repetierer-Klassen gefunden.');
	} else {
		state = 0;
		_filePath.innerText = '';
		_drawerStatus.innerText = 'Die Datei konnte nicht gelesen werden.';
		closeStatsView();
		disableElement(_statsViewBtn);
		updateState();
		renderSessionStartState('Die Datei konnte nicht gelesen werden. Öffne eine gültige Repetierer-Datei.');
		error(_file);
	}
})

ipcRenderer.on('saved-file-missing', (event, filePath) => {
	if (filePath) _filePath.innerText = `Gespeicherte Datei nicht gefunden: ${filePath}`;
	_drawerStatus.innerText = 'Die zuletzt gespeicherte Datei wurde nicht gefunden.';
	closeStatsView();
	disableElement(_statsViewBtn);
	renderSessionStartState(filePath ? 'Die zuletzt verwendete Datei wurde nicht gefunden. Öffne sie erneut oder wähle eine andere Datei.' : '');
});

ipcRenderer.on('joker-migration-required', (event, status, filePath) => {
	state = 0;
	_classes.innerHTML = '';
	_className.innerText = '';
	closeStatsView();
	disableElement(_statsViewBtn);
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
	document.getElementById("version").innerText = "v" + appVersion;
});

ipcRenderer.on('update-news-data', (event, payload) => {
	showUpdateNews(payload || {});
});

ipcRenderer.on('class-statistics-data', (event, stats) => {
	renderClassStatistics(stats || []);
});

ipcRenderer.on('debug-mode', (event, isDebugMode) => {
	if (isDebugMode) {
		_debugActions.classList.remove('update-hidden');
	}
});

ipcRenderer.on('settings-data', (event, settings, paths, migrationStatus) => {
	_extraJokerSetting.checked = !!settings.extraJokerAfterThreeGrades;
	_probabilityFactorSetting.value = settings.probabilityDecreaseFactor;
	_boostNeverSelectedSetting.checked = !!settings.boostNeverSelected;
	_neverSelectedBoostFactorSetting.value = settings.neverSelectedBoostFactor;
	_visualEffectsSetting.checked = !!settings.visualEffectsEnabled;
	updateVisualEffects(settings.visualEffectsEnabled);
	_wiggersRuleSetting.checked = settings.wiggersRuleEnabled !== false;
	_wiggersRuleDurationSetting.value = settings.wiggersRuleDurationMinutes;
	settingsRanges.forEach(updateSettingsRangeValue);
	updateNeverSelectedBoostField();
	updateWiggersRuleFields();
	updateExtraJokerMigrationStatus(migrationStatus);
	_settingsLocation.innerText = paths ? paths.settingsPath : '';
	openModal(_settingsView);
});

ipcRenderer.on('extra-joker-migration-done', (event, result) => {
	updateExtraJokerMigrationStatus(result && result.status);
	const updatedPersons = result && result.updatedPersons ? result.updatedPersons : 0;
	showExtraJokerMigrationMessage(
		'Extra-Joker-Migration',
		updatedPersons > 0
			? `Extra-Joker-Migration abgeschlossen: ${updatedPersons} Personen aktualisiert.`
			: 'Keine berechtigten Personen gefunden.'
	);
});

ipcRenderer.on('extra-joker-migration-failed', (event, result) => {
	updateExtraJokerMigrationStatus(result && result.status ? result.status : extraJokerMigrationStatus);
	const reason = result && result.reason;
	let detail = 'Die Extra-Joker-Migration konnte nicht durchgeführt werden.';
	if (reason === 'no-file-selected') detail = 'Bitte zuerst eine Datei laden.';
	if (reason === 'already-completed') detail = 'Diese Migration wurde für diese Datei bereits durchgeführt.';
	if (reason === 'excel-locked') detail = 'Die Excel-Datei scheint geöffnet zu sein. Bitte schließe sie und versuche es erneut.';
	showExtraJokerMigrationMessage('Extra-Joker-Migration', detail);
});

ipcRenderer.on('update-status', (event, payload) => {
	handleUpdateStatus(payload || {});
});

ipcRenderer.on('window-restored', () => {
	animateRestore();
});

ipcRenderer.on('settings-saved', (event, settings) => {
	applyTheme(settings && settings.theme);
	updateVisualEffects(settings && settings.visualEffectsEnabled);
});

ipcRenderer.on('ui-settings', (event, settings) => {
	applyTheme(settings && settings.theme);
	updateVisualEffects(settings && settings.visualEffectsEnabled);
});

ipcRenderer.on('pending-excel-status', (event, count) => {
	pendingExcelCount = Number(count || 0);
	if (count > 0) {
		enableElement(_flushExcelBtn);
		_flushExcelBtn.innerText = `Excel aktualisieren (${count})`;
		disableElement(_reloadExcelBtn);
		_reloadExcelBtn.title = `${count} lokale Änderung${count === 1 ? '' : 'en'} müssen zuerst mit "Excel aktualisieren" geschrieben werden.`;
		_drawerStatus.innerText = `${count} Eintrag${count === 1 ? '' : 'e'} warten auf Excel.`;
	} else {
		disableElement(_flushExcelBtn);
		_flushExcelBtn.innerText = 'Excel aktualisieren';
		enableElement(_reloadExcelBtn);
		_reloadExcelBtn.title = 'Liest die ausgewählte Excel-Datei erneut ein.';
		if (_filePath.innerText) _drawerStatus.innerText = 'Excel und Backup sind synchron.';
	}
});

ipcRenderer.on('undo-status', (event, status) => {
	if (status && status.available) {
		enableElement(_undoLastBtn);
		const entry = status.entry || {};
		const undoTarget = entry.type === 'joker'
			? `${entry.personName || 'Joker'}: Joker rückgängig machen`
			: `${entry.personName || 'Note'}: Note rückgängig machen`;
		_undoLastBtn.innerText = '↶';
		_undoLastBtn.title = undoTarget;
		_undoLastBtn.setAttribute('aria-label', undoTarget);
	} else {
		disableElement(_undoLastBtn);
		_undoLastBtn.innerText = '↶';
		_undoLastBtn.title = 'Letzte Aktion rückgängig';
		_undoLastBtn.setAttribute('aria-label', 'Letzte Aktion rückgängig');
	}
});

ipcRenderer.on('redo-status', (event, status) => {
	if (status && status.available) {
		enableElement(_redoLastBtn);
		const entry = status.entry || {};
		const redoTarget = entry.type === 'joker'
			? `${entry.personName || 'Joker'}: Joker wiederholen`
			: `${entry.personName || 'Note'}: Note wiederholen`;
		_redoLastBtn.innerText = '↷';
		_redoLastBtn.title = redoTarget;
		_redoLastBtn.setAttribute('aria-label', redoTarget);
	} else {
		disableElement(_redoLastBtn);
		_redoLastBtn.innerText = '↷';
		_redoLastBtn.title = 'Wiederholen';
		_redoLastBtn.setAttribute('aria-label', 'Wiederholen');
	}
});

ipcRenderer.on('undo-completed', (event, result) => {
	const entry = result && result.entry ? result.entry : {};
	state = 2;
	updateState();
	showUpdatePanel({
		title: 'Aktion rückgängig gemacht',
		detail: entry.type === 'joker'
			? `${entry.personName || 'Die Person'} hat den Joker wieder zurückerhalten.`
			: `Die letzte Note von ${entry.personName || 'der Person'} wurde entfernt.`,
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true
	});
});

ipcRenderer.on('redo-completed', (event, result) => {
	const entry = result && result.entry ? result.entry : {};
	state = 2;
	updateState();
	showUpdatePanel({
		title: 'Aktion wiederhergestellt',
		detail: entry.type === 'joker'
			? `${entry.personName || 'Die Person'} hat den Joker wieder gesetzt.`
			: `Die Note von ${entry.personName || 'der Person'} wurde wieder eingetragen.`,
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true
	});
});

ipcRenderer.on('redo-failed', (event, result) => {
	const reason = result && result.reason;
	showUpdatePanel({
		title: 'Wiederholen nicht möglich',
		detail: reason === 'excel-locked'
			? 'Die Excel-Datei scheint geöffnet zu sein. Bitte schließe sie und versuche es erneut.'
			: 'Die rückgängig gemachte Aktion konnte nicht wiederhergestellt werden.',
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true
	});
});

ipcRenderer.on('undo-failed', (event, result) => {
	const reason = result && result.reason;
	showUpdatePanel({
		title: 'Rückgängig nicht möglich',
		detail: reason === 'excel-locked'
			? 'Die Excel-Datei scheint geöffnet zu sein. Bitte schließe sie und versuche es erneut.'
			: reason === 'grade-mismatch'
				? 'Die letzte Note passt nicht mehr zum Excel-Inhalt. Es wurde nichts geändert.'
				: 'Die letzte Aktion konnte nicht rückgängig gemacht werden.',
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true
	});
});

ipcRenderer.on('session-protocol-empty', () => {
	showUpdatePanel({
		title: 'Kein Sitzungsprotokoll',
		detail: 'In dieser Sitzung wurden noch keine Noten oder Joker erfasst.',
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true,
		placement: 'toast'
	});
});

ipcRenderer.on('session-protocol-exported', (event, result) => {
	const detail = result && result.format === 'both' && result.filePaths
		? `PDF und CSV wurden gespeichert: ${result.filePaths.join(' | ')}`
		: result && result.filePath ? result.filePath : 'Das Sitzungsprotokoll wurde gespeichert.';
	showUpdatePanel({
		title: 'Protokoll exportiert',
		detail: detail,
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true
	});
});

ipcRenderer.on('session-protocol-export-failed', (event, result) => {
	showUpdatePanel({
		title: 'Export nicht möglich',
		detail: 'Das Sitzungsprotokoll konnte nicht gespeichert werden.',
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true
	});
});

ipcRenderer.on('class-list-import-preview', (event, result) => {
	pendingClassImport = result;
	_classImportName.value = result.detectedClassName || '';
	_classImportDetail.innerText = `${result.count || 0} Name${result.count === 1 ? '' : 'n'} gefunden. Die Klasse wird zur aktuell geladenen Repetierer-Datei hinzugefügt.`;
	_classImportPreview.innerHTML = '';

	const title = document.createElement('strong');
	title.innerText = result.truncated ? 'Vorschau (erste 25 Namen)' : 'Vorschau';
	_classImportPreview.appendChild(title);

	(result.names || []).slice(0, 25).forEach(name => {
		const row = document.createElement('span');
		row.innerText = name;
		_classImportPreview.appendChild(row);
	});

	openModal(_classImportModal);
	_classImportName.focus();
	_classImportName.select();
});

ipcRenderer.on('class-list-imported', (event, result) => {
	state = 1;
	updateState();
	showUpdatePanel({
		title: 'Klassenliste importiert',
		detail: `${result.count || 0} Namen wurden als Klasse ${result.className || ''} zur aktuellen Repetierer-Datei hinzugefügt.${result.truncated ? ' Es wurden maximal 25 Personen importiert.' : ''}`,
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true
	});
	if (result.className) setActiveClassByName(result.className);
	pendingClassImport = null;
});

ipcRenderer.on('class-list-import-failed', (event, result) => {
	const reason = result && result.reason;
	showUpdatePanel({
		title: 'Import nicht möglich',
		detail: reason === 'no-target-file'
			? 'Lade zuerst eine bestehende Repetierer-Datei. Danach kann die Klasse dort hinzugefügt werden.'
			: reason === 'class-exists'
				? 'Diese Klasse existiert in der aktuellen Datei bereits. Wähle beim Import einen anderen Klassennamen.'
				: reason === 'missing-class-name'
					? 'Bitte gib einen Klassennamen ein.'
					: reason === 'pending-changes'
						? 'Aktualisiere zuerst die Excel-Datei, damit alle lokalen Änderungen gespeichert sind.'
						: reason === 'no-names-found'
							? 'In der ausgewählten Datei wurden keine Namen gefunden.'
							: 'Die Klassenliste konnte nicht importiert werden.',
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true
	});
});

// ready
ipcRenderer.on('ready', (event, args) => {
	if (!args) {
		state = 2;
		_className.innerText = `${_currentClass.innerText}`;
		setActiveClassButton(_currentClass);
		updateState();
		if (!activeTeachingSession && pendingSessionSetupClass === _currentClass.innerText) {
			pendingSessionSetupClass = '';
			ipcRenderer.send('get-session-setup');
		}
		if (activeTeachingSession && _currentClass && activeTeachingSession.className === _currentClass.innerText) {
			ipcRenderer.send('resume-teaching-session');
		}
	} else {
		pendingSessionSetupClass = '';
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
	openModal(_reloadConflictModal);
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
ipcRenderer.on('name', (event, args, selectionNames, selectedPerson) => {
	currentRepetitionToken = selectedPerson?.selectionToken || null;
	aiSelectedStudentId = selectedPerson && selectedPerson.id;
	
	const [name, joker] = args
    
    // check if joker is available
    if (joker) {
        state = 3
    } else {
        state = 4
    }
    
    updateState();
	_grade.value = '';
	const displayName = getFirstName(name);
	_name.innerText = displayName;
	_nameSize = 5.5;
	scaleName();
	runSelectionEffect(name, selectionNames || [], () => {uiMotion.revealStudent(_name);startRepetitionTiming();});

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
		openModal(_personModal);
	} else {
		error(_manualSelectBtn);
	}
});

ipcRenderer.on('absence-persons-data', (event, persons) => {
	renderAbsencePersons(persons || []);
	openModal(_absenceModal);
});

ipcRenderer.on('absences-saved', (event, result) => {
	closeModal(_absenceModal);
	showUpdatePanel({
		title: 'Abwesenheiten gespeichert',
		detail: `${result.count || 0} Person${result.count === 1 ? '' : 'en'} werden heute übersprungen.`,
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true,
		placement: 'toast'
	});
});

ipcRenderer.on('absences-save-failed', (event, result) => {
	showUpdatePanel({
		title: 'Abwesenheiten nicht gespeichert',
		detail: 'Wähle zuerst eine Klasse aus und versuche es erneut.',
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true
	});
});

ipcRenderer.on('backup-data', (event, backup, paths) => {
	const entries = backup && backup.entries ? backup.entries.slice().reverse() : [];
	const totalEntries = backup && Number.isFinite(backup.totalEntries) ? backup.totalEntries : entries.length;
	_backupList.innerHTML = '';
	_backupLocation.innerText = paths ? paths.backupPath : '';

	if (entries.length === 0) {
		const empty = document.createElement('p');
		empty.className = 'backup-empty';
		empty.innerText = 'Noch keine Backup-Einträge vorhanden.';
		_backupList.appendChild(empty);
	} else {
		if (totalEntries > entries.length) {
			const previewInfo = document.createElement('p');
			previewInfo.className = 'backup-empty';
			previewInfo.innerText = `Zeige die letzten ${entries.length} von ${totalEntries} Backup-Eintraegen.`;
			_backupList.appendChild(previewInfo);
		}

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

	openModal(_backupModal);
});

ipcRenderer.on('log-data', (event, logs, paths) => {
	_logLocation.innerText = paths ? paths.logPath : '';
	_logList.innerText = logs && logs.trim() ? logs : 'Noch keine Logs vorhanden.';
	openModal(_logModal);
});

ipcRenderer.on('report-preview-data', (event, payload) => {
	problemReportPreview = payload || {};
	renderProblemReportPreview();
});

ipcRenderer.on('problem-report-result', (event, result) => {
	problemReportSending = false;
	_sendProblemReportBtn.innerText = 'Senden';
	enableElement(_sendProblemReportBtn);

	if (result && result.ok) {
		const reportIdText = result.id ? ` ID: ${result.id}` : '';
		setProblemReportStatus(`Fehlerbericht wurde gesendet.${reportIdText}`, 'success');
		showUpdatePanel({
			title: 'Fehlerbericht wurde gesendet',
			detail: result.id ? `Report-ID: ${result.id}` : 'Danke. Der Bericht wurde gespeichert.',
			primaryVisible: false,
			secondaryText: 'OK',
			secondaryVisible: true
		});
		setTimeout(() => closeProblemReportModal(), 650);
		return;
	}

	setProblemReportStatus('Fehlerbericht konnte nicht gesendet werden. Bitte prüfe deine Internetverbindung.', 'error');
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
			.forEach((person, index) => {
				const item = document.createElement('div');
				item.className = 'probability-item';
				item.style.setProperty('--item-index', Math.min(index, 12));

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
				const wiggersText = person.wiggersRuleActive ? ` | Wiggersche Regel ${person.wiggersRuleRemainingMinutes} min` : '';
				meta.innerText = `${person.grades} Noten${jokerText} | Gewicht ${person.weight.toFixed(2)}${boostText}${wiggersText}`;

				header.appendChild(name);
				header.appendChild(value);
				item.appendChild(header);
				item.appendChild(bar);
				item.appendChild(meta);
				_probabilityList.appendChild(item);
			});
	}

	openModal(_probabilitiesModal);
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

    
ipcRenderer.on('session-setup-data', (event, setup) => {
	if (uiMotion.beginPageChange(() => { ipcRenderer.emit('session-setup-data', event, setup); })) return;
	if (!setup || !setup.className || !Array.isArray(setup.persons) || setup.persons.length === 0) {
		showUpdatePanel({ title: 'Session nicht möglich', detail: 'Wähle zuerst eine Klasse mit Schülern aus.', primaryVisible: false, secondaryText: 'OK', secondaryVisible: true });
		return;
	}
	sessionSetup = setup;
	sessionSetupAbsentIds = [];
	_sessionStartClass.innerText = setup.className;
	renderSessionAttendance();
	enableElement(_confirmSessionStartBtn);
 _confirmSessionStartBtn.disabled=false;
 _cancelSessionStartBtn.disabled=false;
 _sessionAttendanceList.inert=false;
	_sessionEmptyView.classList.add('update-hidden');
 _sessionAttendanceView.classList.remove('update-hidden');
 uiMotion.revealView(_sessionAttendanceView);
 document.getElementById('session-attendance-heading').focus();
});

ipcRenderer.on('teaching-session-started', uiMotion.pageHandler((event, session) => {
	_sessionAttendanceView.classList.add('update-hidden');
 sessionSetup=null;sessionSetupAbsentIds=[];
	setActiveTeachingSession(session);
}));

ipcRenderer.on('teaching-session-start-failed', (event, result) => {
	enableElement(_confirmSessionStartBtn);
 _confirmSessionStartBtn.disabled=false;
 _cancelSessionStartBtn.disabled=false;
 _sessionAttendanceList.inert=false;
	showUpdatePanel({ title: 'Session nicht gestartet', detail: result && result.reason === 'session-already-active' ? 'Es läuft bereits eine Session.' : 'Die Session konnte nicht gestartet werden.', primaryVisible: false, secondaryText: 'OK', secondaryVisible: true });
});

ipcRenderer.on('active-teaching-session-data', (event, session) => {
	activeTeachingSession = session || null;
	if (!session) {
		renderTeachingSessionView();
		return;
	}
	const classButton = Array.from(_classes.querySelectorAll('button')).find(button => button.innerText === session.className);
	if (classButton && (!_currentClass || _currentClass.innerText !== session.className)) {
		setActiveClassButton(classButton);
		ipcRenderer.send('class', session.className);
		return;
	}
	setActiveTeachingSession(session);
});

ipcRenderer.on('teaching-session-ended', () => {
	currentRepetitionToken = null;
	closeModal(_sessionEndModal);
	enableElement(_confirmSessionEndBtn);
	activeTeachingSession = null;
	activeSessionArea = 'participation';
	closeParticipationSummary();
	renderTeachingSessionView();
	state = _currentClass ? 2 : 1;
	updateState();
	showUpdatePanel({ title: 'Session beendet', detail: 'Alle Einträge wurden gespeichert.', primaryVisible: false, secondaryText: 'OK', secondaryVisible: true, placement: 'toast' });
});

ipcRenderer.on('teaching-session-end-failed', () => {
	enableElement(_confirmSessionEndBtn);
	closeModal(_sessionEndModal);
	showUpdatePanel({ title: 'Session nicht beendet', detail: 'Die Session konnte nicht beendet werden. Bitte versuchen Sie es erneut.', primaryVisible: false, secondaryText: 'OK', secondaryVisible: true });
});

ipcRenderer.on('participation-entry-saved', (event, session) => {
	setActiveTeachingSession(session);
});

ipcRenderer.on('participation-entry-failed', () => {
	showUpdatePanel({ title: 'Nicht gespeichert', detail: 'Der Mitarbeitseintrag konnte nicht gespeichert werden.', primaryVisible: false, secondaryText: 'OK', secondaryVisible: true });
});

ipcRenderer.on('participation-summary-data', (event, rows) => renderParticipationSummary(rows || []));
ipcRenderer.on('teaching-session-required', () => {
	showUpdatePanel({ title: 'Session erforderlich', detail: 'Wähle eine Klasse und starte zuerst eine Unterrichts-Session.', primaryVisible: false, secondaryText: 'OK', secondaryVisible: true });
});

function renderSessionAttendance() {
 _sessionAttendanceList.replaceChildren();
 (sessionSetup.persons || []).forEach(person => {
  const button=document.createElement('button');button.type='button';button.className='attendance-person';
  button.textContent=person.name;
  const update=()=>{const absent=sessionSetupAbsentIds.some(id=>String(id)===String(person.id));button.classList.toggle('is-absent',absent);button.setAttribute('aria-pressed',String(absent));button.setAttribute('aria-label',`${person.name}, ${absent?'abwesend':'anwesend'}`);};
  update();button.addEventListener('click',()=>{
   if(sessionSetupAbsentIds.some(id=>String(id)===String(person.id))) sessionSetupAbsentIds=sessionSetupAbsentIds.filter(id=>String(id)!==String(person.id));
   else sessionSetupAbsentIds.push(person.id);
   update();updateAttendancePreview();
  });
  _sessionAttendanceList.append(button);
 });
 updateAttendancePreview();
}

function updateAttendancePreview() {
	const total = sessionSetup && sessionSetup.persons ? sessionSetup.persons.length : 0;
	_sessionPresentPreview.innerText = `${total - sessionSetupAbsentIds.length} anwesend`;
	_sessionAbsentPreview.innerText = `${sessionSetupAbsentIds.length} abwesend`;
}

function setActiveTeachingSession(session) {
	activeTeachingSession = session;
	_sessionClassName.innerText = session.className;
	renderParticipationStudents();
	renderTeachingSessionView();
	updateState();
}

function renderTeachingSessionView() {
	if (!_settingsView.classList.contains('update-hidden')) return;
	const assessmentView = document.getElementById('assessment-view');
	if (assessmentView && !assessmentView.classList.contains('update-hidden')) return;
	if (_statsView && !_statsView.classList.contains('update-hidden')) return;
	const isActive = !!activeTeachingSession;
	const destination = isActive ? (activeSessionArea === 'repetition' ? _repetition : _participationView) : (sessionSetup ? _sessionAttendanceView : _sessionEmptyView);
	if (destination.classList.contains('update-hidden') && uiMotion.beginPageChange(() => { renderTeachingSessionView(); })) return;
	const homeWasHidden = _sessionEmptyView.classList.contains('update-hidden');
	_contentToggleSessionClass(isActive);
	_sessionHeader.classList.toggle('update-hidden', !isActive);
	_classPanel.classList.toggle('update-hidden', !isActive);
	const classList = document.getElementById('class-list');
	if (isActive) document.getElementById('session-class-slot').appendChild(classList);
	else _sessionEmptyView.insertBefore(classList, document.getElementById('session-class-home-slot'));
	_sessionEmptyView.classList.toggle('update-hidden', isActive || !!sessionSetup);
 _sessionAttendanceView.classList.toggle('update-hidden', isActive || !sessionSetup);
	if (!isActive) {
		_participationView.classList.add('update-hidden');
		_repetition.classList.add('update-hidden');
		renderSessionStartState();
		if (homeWasHidden) uiMotion.revealView(_sessionEmptyView);
		return;
	}
	showSessionArea(activeSessionArea);
}

function renderSessionStartState(message = '') {
	const count = _classes.querySelectorAll('button').length;
	const hasClasses = count > 0;
	document.getElementById('session-file-empty').classList.toggle('update-hidden', hasClasses);
	document.getElementById('class-list').classList.toggle('update-hidden', !hasClasses);
	document.getElementById('session-source').classList.toggle('update-hidden', !hasClasses);
	document.getElementById('session-start-description').textContent = hasClasses
		? 'Wähle eine Klasse und erfasse die heutigen Abwesenheiten.'
		: 'Lade deine Klassen, um eine Unterrichtssession zu beginnen.';
	document.getElementById('session-start-heading').textContent = hasClasses ? 'Klasse wählen' : 'Unterricht starten';
	document.getElementById('session-file-heading').textContent = 'Excel-Datei auswählen';
	document.getElementById('session-file-description').textContent = message || 'Öffne eine Repetierer-Datei oder importiere eine Klassenliste.';
	const fileName = _filePath.textContent.split(/[\\/]/).pop();
	document.getElementById('session-source-label').textContent = `${count} ${count === 1 ? 'Klasse' : 'Klassen'}${fileName ? ` · ${fileName}` : ''}`;
	document.getElementById('session-workbook-name').textContent = fileName || 'Repetierer-Datei';
	document.getElementById('session-workbook-name').title = _filePath.textContent;
	_classes.querySelectorAll('button').forEach(button => { button.disabled = !!activeTeachingSession && button.innerText !== activeTeachingSession.className; });
}

// Shared placeholders for actual IPC reads; decoration is hidden from assistive technology.
window.createRepetiererLoadingState = (label, layout = 'list') => {
	const state = document.createElement('div');
	state.className = `ui-skeleton-state ui-skeleton-${layout}`;
	state.setAttribute('role', 'status');
	const status = document.createElement('p');
	status.className = 'ui-loading-status';
	status.textContent = label;
	const shapes = document.createElement('div');
	shapes.className = 'ui-skeleton-shapes';
	shapes.setAttribute('aria-hidden', 'true');
	shapes.innerHTML = '<div class="skeleton-block skeleton-title"></div><div class="skeleton-metrics"><div class="skeleton-block"></div><div class="skeleton-block"></div><div class="skeleton-block"></div><div class="skeleton-block"></div></div><div class="skeleton-block skeleton-row"></div><div class="skeleton-block skeleton-row"></div><div class="skeleton-block skeleton-row"></div>';
	state.append(status, shapes);
	return state;
};

function _contentToggleSessionClass(isActive) {
	const content = document.getElementById('content');
	if (content) content.classList.toggle('session-active', isActive);
}

function showSessionArea(area) {
	if (!activeTeachingSession) return;
	activeSessionArea = area === 'repetition' ? 'repetition' : 'participation';
	const participationActive = activeSessionArea === 'participation';
	const nextView = participationActive ? _participationView : _repetition;
	const viewWasHidden = nextView.classList.contains('update-hidden');
	_participationView.classList.toggle('update-hidden', !participationActive);
	_repetition.classList.toggle('update-hidden', participationActive);
	_sessionParticipationTab.className = participationActive ? 'btn-1' : 'btn-2';
	_sessionRepetitionTab.className = participationActive ? 'btn-2' : 'btn-1';
	_sessionParticipationTab.setAttribute('aria-pressed', String(participationActive));
	_sessionRepetitionTab.setAttribute('aria-pressed', String(!participationActive));
	if (viewWasHidden) uiMotion.revealView(nextView);
}

function renderParticipationStudents() {
	_participationStudentList.innerHTML = '';
	const entryCounts = new Map();
	(activeTeachingSession.participationEntries || []).forEach(entry => {
		const key = String(entry.studentId);
		entryCounts.set(key, (entryCounts.get(key) || 0) + 1);
	});
	(activeTeachingSession.presentStudents || []).forEach(person => {
		const button = document.createElement('button');
		button.className = 'participation-student';
		const name = document.createElement('strong');
		name.innerText = person.name;
		const count = document.createElement('span');
		const value = entryCounts.get(String(person.id)) || 0;
		count.innerText = `${value} ${value === 1 ? 'Eintrag' : 'Einträge'}`;
		button.append(name, count);
		button.addEventListener('click', () => openParticipationScore(person));
		_participationStudentList.appendChild(button);
	});
}

function openParticipationScore(person) {
	delete _participationScoreList.dataset.submitting;
	selectedParticipationStudent = person;
	_participationStudentName.innerText = person.name;
	_participationScoreList.innerHTML = '';
	const descriptions = [
		'falsche Antwort',
		'originell, aber unpassend',
		'sachlich korrekt, aber einfach / stichwortartig',
		'passender Beitrag mit erkennbarem Mitdenken',
		'weiterführender Beitrag, der auf den Kern hinsteuert',
		'diskussionsprägender Beitrag / neue Perspektive'
	];
	descriptions.forEach((description, index) => {
		const button = document.createElement('button');
		button.className = 'btn-2 participation-score';
		const score = document.createElement('strong');
		score.innerText = String(index + 1);
		const text = document.createElement('span');
		text.innerText = description;
		button.append(score, text);
		button.addEventListener('click', () => {
			if (_participationScoreList.dataset.submitting === 'true') return;
			_participationScoreList.dataset.submitting = 'true';
			_participationScoreList.querySelectorAll('button').forEach(item => {item.disabled = true;});
			ipcRenderer.send('add-participation-entry', { sessionId: activeTeachingSession.id, studentId: person.id, points: index + 1 });
			uiMotion.confirmScore(button, () => closeModal(_participationModal));
		});
		_participationScoreList.appendChild(button);
	});
	openModal(_participationModal);
}

function renderParticipationSummary(rows) {
	if (_participationSummary.classList.contains('update-hidden') && uiMotion.beginPageChange(() => { renderParticipationSummary(rows); })) return;
	_participationStudentList.classList.add('update-hidden');
	_participationSummary.classList.remove('update-hidden');
	_participationSummaryList.innerHTML = '';
	if (!rows.length) {
		const empty = document.createElement('p');
		empty.className = 'stats-empty';
		empty.innerText = 'Noch keine Sessiondaten vorhanden.';
		_participationSummaryList.appendChild(empty);
		return;
	}
	rows.forEach(row => {
		const card = document.createElement('article');
		card.className = 'participation-summary-card';
		const heading = document.createElement('h4');
		heading.innerText = row.studentName;
		const details = document.createElement('dl');
		const metrics = [
			['Besuchte Sessions', row.visitedSessions],
			['Sessions mit Mitarbeit', row.sessionsWithParticipation],
			['Bewertungen', row.ratingCount],
			['Durchschnitt', row.averagePoints == null ? '—' : row.averagePoints.toFixed(2)],
			['Beteiligungsquote', row.participationRate == null ? '—' : `${Math.round(row.participationRate * 100)} %`]
		];
		metrics.forEach(([label, value]) => {
			const dt = document.createElement('dt'); dt.innerText = label;
			const dd = document.createElement('dd'); dd.innerText = String(value);
			details.append(dt, dd);
		});
		card.append(heading, details);
		if (!row.dataSufficient) {
			const note = document.createElement('span');
			note.className = 'data-insufficient';
			note.innerText = 'Datengrundlage unzureichend';
			card.appendChild(note);
		}
		_participationSummaryList.appendChild(card);
	});
}

function closeParticipationSummary() {
	if (!_participationSummary.classList.contains('update-hidden') && uiMotion.beginPageChange(() => { closeParticipationSummary(); })) return;
	_participationSummary.classList.add('update-hidden');
	_participationStudentList.classList.remove('update-hidden');
}

window.restoreTeachingSessionView = renderTeachingSessionView;
window.closeRepetiererSettings = afterClose => closeSettingsPage(afterClose);
window.dismissRepetiererSettingsForNavigation = dismissSettingsForNavigation;
window.navigateRepetiererAi = command => {
	if (!selectClassForNavigation(command.classId)) return false;
	if (typeof window.openRepetiererRepetition === 'function') window.openRepetiererRepetition();
	else renderTeachingSessionView();
	if (command.page === 'evaluation') openStatsView();
	return true;
};
window.getRepetiererAiContext = () => ({
	page: document.getElementById('assessment-view').classList.contains('update-hidden') ?
		(document.getElementById('stats-view').classList.contains('update-hidden') ? activeSessionArea : 'statistics') : 'evaluation',
	classId: _currentClass ? _currentClass.innerText : null,
	studentId: (state === 3 || state === 4) && Number.isInteger(aiSelectedStudentId) ? aiSelectedStudentId : null,
	sessionId: activeTeachingSession ? activeTeachingSession.id : null
});

/*
 * other
 */

// update the buttons based on the state
function updateState() {
	if (state !== 3 && state !== 4) clearSelectionEffect();
	switch (state) {
		case 0:
			_classes.innerHTML = '';
			_className.innerText = '';
			disableElement(_statsViewBtn);
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
			disable(_absencesBtn);
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
			enable(_absencesBtn);
			enable(_excelEditorBtn);
			ipcRenderer.send('get-undo-status');
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
			disable(_absencesBtn);
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
			disable(_absencesBtn);
			disable(_excelEditorBtn);
			break;

	}

	if (state === 2 && !activeTeachingSession) {
		disableElement(_start);
		disableElement(_manualSelectBtn);
		disableElement(_probabilitiesBtn);
	}
	disableElement(_absencesBtn);
	renderSessionStartState();
	uiMotion.setGradingVisible(document.getElementById('res'), state === 3 || state === 4);
	document.getElementById('repetition-idle-hint').classList.toggle('update-hidden', state === 3 || state === 4);

	function disable(e) {
		disableElement(e);
	}

	function enable(e) {
		enableElement(e);
	}

	function text() {
		_grade.value = '';
		_name.innerText = 'Bereit für die nächste Auswahl';
		_nameSize = 5.5;
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

function renderAbsencePersons(persons) {
	_absenceList.innerHTML = '';
	absentPersonIds = [];

	if (!persons || persons.length === 0) {
		const empty = document.createElement('p');
		empty.className = 'backup-empty';
		empty.innerText = 'Keine Personen in dieser Klasse.';
		_absenceList.appendChild(empty);
		return;
	}

	persons.forEach(person => {
		const personDiv = document.createElement('div');
		personDiv.className = 'person-item';

		const checkbox = document.createElement('input');
		checkbox.type = 'checkbox';
		checkbox.id = `absence-${person.id}`;
		checkbox.checked = !!person.absent;
		if (checkbox.checked) absentPersonIds.push(person.id);
		checkbox.addEventListener('change', (e) => {
			if (e.target.checked) {
				absentPersonIds.push(person.id);
			} else {
				absentPersonIds = absentPersonIds.filter(id => id !== person.id);
			}
		});

		const label = document.createElement('label');
		label.htmlFor = `absence-${person.id}`;
		label.textContent = `${person.name}`;
		label.style.marginLeft = '10px';
		label.style.cursor = 'pointer';
		label.style.flexGrow = '1';

		const meta = document.createElement('small');
		meta.innerText = `${person.grades} Noten | ${person.joker} Joker`;

		personDiv.appendChild(checkbox);
		personDiv.appendChild(label);
		personDiv.appendChild(meta);
		_absenceList.appendChild(personDiv);
	});
}

function openStatsView() {
	if (_statsViewBtn.classList.contains('disabled')) return;
	if (uiMotion.beginPageChange(() => { openStatsView(); })) return;
	dismissSettingsForNavigation();
	uiMotion.cancel(_statsView);
	resetStatsScrollMotion();
	disconnectStatsCardObserver();
	closeDrawer();
	document.getElementById('content').classList.remove('session-active');
	_statsViewBtn.setAttribute('aria-current', 'page');
	document.getElementById('repetition-nav-btn').removeAttribute('aria-current');
	document.getElementById('assessment-nav-btn').removeAttribute('aria-current');
	document.getElementById('assessment-view').classList.add('update-hidden');
	_sessionHeader.classList.add('update-hidden');
	_participationView.classList.add('update-hidden');
	_sessionEmptyView.classList.add('update-hidden');
 _sessionAttendanceView.classList.add('update-hidden');
	[_classPanel, _repetition].forEach(panel => {
		if (!panel) return;
		panel.classList.remove('stats-panel-return');
		panel.classList.add('update-hidden');
	});
	_statsView.classList.remove('stats-closing', 'stats-opening', 'stats-results-entering');
	_statsView.classList.remove('update-hidden');
	uiMotion.revealView(_statsView);
	_statsEmpty.classList.add('update-hidden');
	document.getElementById('stats-class-picker').classList.add('update-hidden');
	_statsList.replaceChildren(window.createRepetiererLoadingState('Statistiken werden geladen …', 'stats'));
	_statsList.setAttribute('aria-busy', 'true');
	_statsList.scrollTop = 0;
	statsLastScrollTop = 0;
	ipcRenderer.send('get-class-statistics');
}

function closeStatsView() {
	if (!_statsView) return;
	if (!_statsView.classList.contains('update-hidden') && uiMotion.beginPageChange(() => { closeStatsView(); })) return;
	_statsViewBtn.removeAttribute('aria-current');
	document.getElementById('repetition-nav-btn').setAttribute('aria-current', 'page');
	if (activeTeachingSession && !_statsView.classList.contains('update-hidden')) {
		_statsView.classList.add('update-hidden');
		_statsView.classList.remove('stats-closing', 'stats-opening', 'stats-results-entering');
		renderTeachingSessionView();
		return;
	}
	if (_statsView.classList.contains('update-hidden')) {
		renderTeachingSessionView();
		return;
	}
	if (_statsView.classList.contains('stats-closing')) return;

	resetStatsScrollMotion(false);
	disconnectStatsCardObserver();
	clearTimeout(statsResultsTimer);
	_statsView.classList.remove('stats-opening', 'stats-results-entering');
	_statsView.classList.add('stats-closing');
	_statsView.classList.add('update-hidden');
	_statsView.classList.remove('stats-closing', 'stats-scrolling-up', 'stats-scrolling-down', 'stats-is-scrolling');
	renderTeachingSessionView();
}

function renderClassStatistics(stats) {
	const select = document.getElementById('stats-class-select');
	const previous = select.value;
	select.replaceChildren();
	_statsList.innerHTML = '';
	_statsList.setAttribute('aria-busy', 'false');
	document.getElementById('stats-class-picker').classList.toggle('update-hidden', !stats?.length);

	if (!stats || stats.length === 0) {
		_statsEmpty.classList.remove('update-hidden');
		return;
	}

	_statsEmpty.classList.add('update-hidden');
	stats.forEach(classStats => {
		select.add(new Option(classStats.className, classStats.className));
		const card = createStatsClassCard(classStats);
		_statsList.appendChild(card);
	});
	const selected = stats.some(item => item.className === previous) ? previous : stats.some(item => item.className === _currentClass?.innerText) ? _currentClass.innerText : stats[0].className;
	selectStatsClass(selected);
}

function selectStatsClass(className) {
	const select = document.getElementById('stats-class-select');
	if (![...select.options].some(option => option.value === className)) return false;
	select.value = className;
	for (const card of _statsList.children) card.classList.toggle('update-hidden', card.dataset.className !== className);
	_statsList.scrollTop = 0;
	setupStatsCardObserver();
	window.repetiererKristall?.refresh();
	return true;
}
document.getElementById('stats-class-select').addEventListener('change', event => {
	const className = event.target.value;
	uiMotion.pageChange(() => selectStatsClass(className));
});
window.selectRepetiererStatsClass = selectStatsClass;

function handleStatsScroll() {
	if (_statsView.classList.contains('update-hidden') || statsScrollFrame) return;
	statsScrollFrame = requestAnimationFrame(() => {
		const currentScrollTop = _statsList.scrollTop;
		const difference = currentScrollTop - statsLastScrollTop;
		if (Math.abs(difference) > 1) {
			const scrollingDown = difference > 0;
			_statsView.classList.toggle('stats-scrolling-down', scrollingDown);
			_statsView.classList.toggle('stats-scrolling-up', !scrollingDown);
			_statsView.classList.add('stats-is-scrolling');
			statsLastScrollTop = currentScrollTop;
			clearTimeout(statsScrollIdleTimer);
			statsScrollIdleTimer = setTimeout(() => {
				_statsView.classList.remove('stats-is-scrolling');
			}, motionDelay(180));
		}
		statsScrollFrame = null;
	});
}

function setupStatsCardObserver() {
 disconnectStatsCardObserver();
 uiMotion.refreshScrollEffects();
}

function disconnectStatsCardObserver() {
	if (statsCardObserver) {
		statsCardObserver.disconnect();
		statsCardObserver = null;
	}
	Array.from(_statsList.children).forEach(card => {
		card.classList.remove('stats-card-outside', 'stats-from-top', 'stats-from-bottom');
	});
}

function resetStatsScrollMotion(removeDirection = true) {
	if (statsScrollFrame) {
		cancelAnimationFrame(statsScrollFrame);
		statsScrollFrame = null;
	}
	clearTimeout(statsScrollIdleTimer);
	_statsView.classList.remove('stats-is-scrolling');
	if (removeDirection) {
		_statsView.classList.remove('stats-scrolling-up', 'stats-scrolling-down');
	}
}

function createStatsClassCard(classStats) {
	const card = document.createElement('section');
	card.className = 'stats-class-card stats-class-detail';
	card.dataset.className = classStats.className;

	const heading = document.createElement('div');
	heading.className = 'stats-class-heading';
	const title = document.createElement('h3');
	title.innerText = classStats.className || 'Klasse';
	const count = document.createElement('span');
	count.textContent = `${formatInteger(classStats.personCount)} Personen`;
	heading.append(title, count);
	card.appendChild(heading);

	const rows = [
		['Ø Note', formatDecimal(classStats.averageGrade), 'Notenskala 1–6'],
		['Repetitionen', formatInteger(classStats.repetitionsTotal), 'In der Excel-Datei erfasst'],
		['Ø Repetitionsdauer', formatRepetitionDuration(classStats.averageRepetitionDurationMs), 'Minuten : Sekunden'],
		['Ø pro Person', formatDecimal(classStats.averageRepetitionsPerPerson), 'Repetitionen']
	];
	const metrics = document.createElement('div');
	metrics.className = 'stats-metrics';
	rows.forEach(([label, value, caption]) => {
		const row = document.createElement('div');
		row.className = 'stats-metric';
		const labelElement = document.createElement('span');
		labelElement.innerText = label;
		const valueElement = document.createElement('strong');
		valueElement.innerText = value;
		const hint = document.createElement('small');
		hint.textContent = caption;
		row.append(labelElement, valueElement, hint);
		metrics.appendChild(row);
	});
	card.appendChild(metrics);
	const detail = document.createElement('div');
	detail.className = 'stats-detail';
	const topLabel = document.createElement('span'); topLabel.textContent = 'Am häufigsten repetiert';
	const top = document.createElement('strong'); top.textContent = formatTopRepetitionPerson(classStats);
	detail.append(topLabel, top);
	card.appendChild(detail);
	const note = document.createElement('p');
	note.className = 'stats-timing-note';
	const measured = Number(classStats.timedRepetitionsCount || 0);
	note.textContent = measured > 0 ? `Zeitmessung: ${measured} ${measured === 1 ? 'Repetition' : 'Repetitionen'} auf diesem Gerät. Von der Namensauswahl bis zum Speichern. Ältere Einträge ohne Messung zählen nicht mit.` : 'Noch keine Zeitmessung. Neue Repetitionen werden von der Namensauswahl bis zum Speichern gemessen. Joker und abgebrochene Versuche zählen nicht mit.';
	card.appendChild(note);
	return card;
}

function formatRepetitionDuration(value) {
	if (typeof value !== 'number' || !Number.isFinite(value) || value < 0) return '–';
	const seconds = Math.round(value / 1000);
	return `${Math.floor(seconds / 60)}:${String(seconds % 60).padStart(2, '0')}`;
}

function formatTopRepetitionPerson(classStats) {
	const person = classStats && classStats.topRepetitionPerson;
	if (!person || !person.name) return '–';
	const additionalCount = Number(classStats.topRepetitionAdditionalCount || 0);
	return `${person.name} (${formatInteger(person.repetitions)})${additionalCount > 0 ? ` + ${additionalCount} weitere` : ''}`;
}

function formatInteger(value) {
	const number = Number(value);
	return Number.isFinite(number) ? String(number) : '–';
}

function formatDecimal(value) {
	if (value === null || value === undefined || value === '') return '–';
	const number = Number(value);
	return Number.isFinite(number) ? number.toFixed(2) : '–';
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

function openTeacherHelp(topic) {
	const help = {
		class: {
			title: 'Klasse',
			paragraphs: [
				'Nach dem Laden der Excel-Datei erscheinen hier alle verfügbaren Klassenblätter.',
				'Die ausgewählte Klasse bestimmt, aus welcher Namensliste Repetierer Personen auswählt.',
				'Die Klasse kann jederzeit gewechselt werden, solange gerade keine Note eingetragen wird.'
			]
		},
		repetition: {
			title: 'Repetition',
			paragraphs: [
				'Start wählt eine Person zufällig aus. Personen mit weniger Noten werden dabei stärker berücksichtigt.',
				'Freiwillig erlaubt eine manuelle Auswahl, Chancen zeigt die aktuelle Gewichtung und Abwesend überspringt Personen für den Tag.',
				'Nach einer Auswahl kann eine Note gespeichert oder ein Joker gesetzt werden. Die runden Pfeile machen die letzte Aktion rückgängig oder stellen sie wieder her.'
			]
		},
		'drawer-main': {
			title: 'Dateien',
			paragraphs: [
				'Dateiverwaltung bündelt das Auswählen der Excel-Datei, den Import von Klassenlisten und das Synchronisieren der Excel-Daten.',
				'Protokoll öffnet Backups, technische Logs und den Export des Sitzungsprotokolls.',
				'Statistiken zeigt Auswertungen der geladenen Klassen und bleibt direkt aus diesem Menü erreichbar.'
			]
		},
		'drawer-data': {
			title: 'Dateiverwaltung',
			paragraphs: [
				'Datei auswählen lädt die Excel-Datei, aus der Klassen, Namen, Noten und Joker gelesen werden.',
				'Klassenliste importieren übernimmt eine Namensliste aus Excel oder CSV und erstellt daraus eine neue Repetierer-Klasse.',
				'Excel aktualisieren schreibt wartende lokale Änderungen in die Excel-Datei, sobald Änderungen vorhanden sind.',
				'Excel neu laden liest die ausgewählte Excel-Datei erneut ein. Bei wartenden lokalen Änderungen bleibt diese Funktion gesperrt.',
				'Excel bearbeiten öffnet die Bearbeitung von Personen und Jokern der ausgewählten Klasse.'
			]
		},
		'drawer-protocol': {
			title: 'Protokoll',
			paragraphs: [
				'Backup anzeigen zeigt gespeicherte Backup-Einträge dieser Sitzung und Datei.',
				'Log anzeigen öffnet technische Meldungen, falls etwas geprüft werden muss.',
				'Protokoll exportieren speichert das Sitzungsprotokoll als PDF und CSV.',
				'Problem melden sendet einen bereinigten Fehlerbericht mit Kommentar und den letzten Logeinträgen.'
			]
		}
	}[topic];

	if (!help) return;

	_teacherHelpTitle.innerText = help.title;
	_teacherHelpBody.innerHTML = '';
	help.paragraphs.forEach(text => {
		const paragraph = document.createElement('p');
		paragraph.innerText = text;
		_teacherHelpBody.appendChild(paragraph);
	});
	openModal(_teacherHelpModal);
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
		setDrawerView('main');
		uiMotion.setDrawerOpen(true);
	}
}

function closeDrawer() {
	uiMotion.setDrawerOpen(false);
}

function setDrawerView(view) {
	if (drawerView === view && !document.getElementById('drawer-content').dataset.viewMotion) return;
	uiMotion.changeDrawerView(() => {
	drawerView = view;
	_drawer.classList.toggle('drawer-submenu', view !== 'main');
	_drawerTitle.innerText = view === 'data' ? 'Dateiverwaltung' : view === 'protocol' ? 'Protokoll' : 'Navigation';

	[_drawerMainView, _drawerDataView, _drawerProtocolView].forEach(drawerViewElement => {
		drawerViewElement.classList.add('update-hidden');
	});

	if (view === 'data') {
		_drawerDataView.classList.remove('update-hidden');
		_drawerBackBtn.classList.remove('update-hidden');
	} else if (view === 'protocol') {
		_drawerProtocolView.classList.remove('update-hidden');
		_drawerBackBtn.classList.remove('update-hidden');
	} else {
		_drawerMainView.classList.remove('update-hidden');
		_drawerBackBtn.classList.add('update-hidden');
	}
	});
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
	const alertTone = options.tone || (/(nicht|Fehler|fehlgeschlagen|ungültig)/i.test(options.title || '') ? 'error' : /(gespeichert|erfolgreich|exportiert|gesendet|aktualisiert|beendet)/i.test(options.title || '') ? 'success' : 'info');
	_updatePanel.classList.remove('ui-alert-info', 'ui-alert-success', 'ui-alert-warning', 'ui-alert-error');
	_updatePanel.classList.add(`ui-alert-${alertTone}`);
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

function showUpdateNews(payload) {
	if (!payload.version || !payload.notes) return;

	updateNewsVersion = payload.version;
	_updateNewsTitle.innerText = `Update ${payload.version}`;
	renderUpdateNewsNotes(payload.notes);
	openModal(_updateNewsModal);
}

function renderUpdateNewsNotes(notes) {
	_updateNewsBody.innerHTML = '';
	let currentList = null;

	String(notes).split(/\r?\n/).forEach(line => {
		const text = line.trim();
		if (!text) {
			currentList = null;
			return;
		}

		if (text.startsWith('- ')) {
			if (!currentList) {
				currentList = document.createElement('ul');
				_updateNewsBody.appendChild(currentList);
			}
			const item = document.createElement('li');
			item.innerText = text.slice(2).trim();
			currentList.appendChild(item);
			return;
		}

		currentList = null;
		const paragraph = document.createElement('p');
		paragraph.innerText = text;
		_updateNewsBody.appendChild(paragraph);
	});
}

function closeUpdateNewsModal() {
	closeModal(_updateNewsModal);
	ipcRenderer.send('update-news-dismissed', updateNewsVersion);
	updateNewsVersion = '';
}

function openProblemReportModal() {
	if (!problemReportsEnabled) return;
	problemReportPreview = null;
	problemReportSending = false;
	_problemReportComment.value = '';
	_problemReportSummary.innerHTML = '';
	_problemReportDetails.innerText = 'Vorschau wird geladen...';
	_problemReportDetails.classList.add('update-hidden');
	_toggleProblemReportDetailsBtn.innerText = 'Details anzeigen';
	setProblemReportStatus('', '');
	_sendProblemReportBtn.innerText = 'Senden';
	enableElement(_sendProblemReportBtn);
	openModal(_problemReportModal);
	ipcRenderer.send('get-report-preview');
}

function closeProblemReportModal() {
	if (problemReportSending) return;
	closeModal(_problemReportModal);
}

function renderProblemReportPreview() {
	if (!problemReportPreview) return;

	const comment = _problemReportComment.value || '';
	const payload = Object.assign({}, problemReportPreview, {
		message: comment || problemReportPreview.message || ''
	});
	const metadata = payload.metadata || {};
	const summaryRows = [
		['App', payload.appId || metadata.appName || ''],
		['Typ', payload.type || ''],
		['Version', payload.version || ''],
		['Kommentar', comment ? `${comment.length} Zeichen` : 'leer'],
		['Log', payload.consoleOutput ? `${payload.consoleOutput.length} Zeichen, bereinigt und gekürzt` : 'leer'],
		['Zeitpunkt', metadata.clientTimestamp || ''],
		['System', [metadata.os, metadata.platform].filter(Boolean).join(' / ')],
		['Umgebung', payload.environment || '']
	];

	_problemReportSummary.innerHTML = '';
	summaryRows.forEach(([label, value]) => {
		const row = document.createElement('div');
		row.className = 'report-summary-row';
		const labelElement = document.createElement('span');
		labelElement.innerText = label;
		const valueElement = document.createElement('strong');
		valueElement.innerText = value || '-';
		row.appendChild(labelElement);
		row.appendChild(valueElement);
		_problemReportSummary.appendChild(row);
	});

	_problemReportDetails.innerText = JSON.stringify(payload, null, 2);
}

function sendProblemReport() {
	if (!problemReportsEnabled) return;
	if (problemReportSending) return;
	if ((_problemReportComment.value || '').length > 4000) {
		setProblemReportStatus('Der Kommentar ist zu lang.', 'error');
		return;
	}

	problemReportSending = true;
	setProblemReportStatus('Fehlerbericht wird gesendet...', '');
	_sendProblemReportBtn.innerText = 'Senden...';
	disableElement(_sendProblemReportBtn);
	ipcRenderer.send('send-problem-report', {
		comment: _problemReportComment.value || ''
	});
}

function setProblemReportStatus(message, type) {
	_problemReportStatus.innerText = message || '';
	_problemReportStatus.classList.remove('report-status-error', 'report-status-success');
	if (type === 'error') _problemReportStatus.classList.add('report-status-error');
	if (type === 'success') _problemReportStatus.classList.add('report-status-success');
}

let settingsReturnState = null;
function openSettingsPage() {
	if (!_settingsView.classList.contains('update-hidden')) return;
	uiMotion.pageChange(() => {
		if (!_settingsView.classList.contains('update-hidden')) return;
		const content = document.getElementById('content');
		settingsReturnState = {
			views: [...content.children].filter(element => element !== _settingsView).map(element => [element, element.classList.contains('update-hidden')]),
			sessionActive: content.classList.contains('session-active')
		};
		for (const [element] of settingsReturnState.views) element.classList.add('update-hidden');
		content.classList.remove('session-active');
		content.classList.add('settings-active');
		_settingsView.classList.remove('update-hidden');
		_settingsBtn.setAttribute('aria-current', 'page');
		closeDrawer();
		document.getElementById('settings-back-btn').focus();
	});
}
function dismissSettingsForNavigation() {
	_settingsView.classList.add('update-hidden');
	document.getElementById('content').classList.remove('settings-active');
	_settingsBtn.removeAttribute('aria-current');
	settingsReturnState = null;
}
function closeSettingsPage(afterClose) {
	if (!settingsReturnState) { afterClose?.(); return; }
	uiMotion.pageChange(() => {
		const previous = settingsReturnState;
		if (!previous) return;
		dismissSettingsForNavigation();
		for (const [element, hidden] of previous.views) element.classList.toggle('update-hidden', hidden);
		document.getElementById('content').classList.toggle('session-active', previous.sessionActive);
		renderTeachingSessionView();
		_settingsBtn.focus();
	}, afterClose);
}
document.addEventListener('keydown', event => {
	if (event.key === 'Escape' && !_settingsView.classList.contains('update-hidden') && !document.querySelector('.modal-open, .modal-closing')) {
		event.preventDefault(); closeSettingsPage();
	}
});
function openModal(modal) {
	if (modal === _settingsView) return openSettingsPage();
	uiMotion.openDialog(modal);
}

function closeModal(modal, afterClose) {
	if (modal === _settingsView) return closeSettingsPage(afterClose);
	uiMotion.closeDialog(modal, afterClose);
}

function motionDelay(duration) {
	return prefersReducedMotion.matches ? 0 : duration;
}

function closeSettingsHelp() {
	closeModal(_settingsHelpModal, () => _settingsHelpBtn.focus());
}

function closeJokerMigrationHelp() {
	closeModal(_jokerMigrationHelpModal, () => openModal(_jokerMigrationModal));
}

function updateExtraJokerMigrationStatus(status) {
	extraJokerMigrationStatus = Object.assign({ hasFile: false, completed: false }, status || {});

	if (!extraJokerMigrationStatus.hasFile) {
		_extraJokerMigrationStatus.innerText = 'Bitte zuerst eine Datei laden.';
		disableElement(_extraJokerMigrationBtn);
		_extraJokerMigrationBtn.disabled = true;
		return;
	}

	if (extraJokerMigrationStatus.completed) {
		_extraJokerMigrationStatus.innerText = 'Diese Migration wurde für diese Datei bereits durchgeführt.';
		disableElement(_extraJokerMigrationBtn);
		_extraJokerMigrationBtn.disabled = true;
		return;
	}

	_extraJokerMigrationStatus.innerText = 'Bereit: prüft alle Personen mit mindestens 3 Noten in der aktuellen Datei.';
	enableElement(_extraJokerMigrationBtn);
	_extraJokerMigrationBtn.disabled = false;
}

function showExtraJokerMigrationMessage(title, detail) {
	showUpdatePanel({
		title: title,
		detail: detail,
		primaryVisible: false,
		secondaryText: 'OK',
		secondaryVisible: true,
		placement: 'toast'
	});
}

function updateNeverSelectedBoostField() {
	const isEnabled = _boostNeverSelectedSetting.checked;
	_neverSelectedBoostFactorSetting.disabled = !isEnabled;
	_neverSelectedBoostFactorSetting.closest('.setting-field').classList.toggle('setting-disabled', !isEnabled);
}

function updateWiggersRuleFields() {
	const isEnabled = _wiggersRuleSetting.checked;
	_wiggersRuleDurationSetting.disabled = !isEnabled;
	_wiggersRuleDurationSetting.closest('.setting-field').classList.toggle('setting-disabled', !isEnabled);
}

function runSelectionEffect(selectedName, names, onComplete) {
	clearSelectionEffect();
	if (!visualEffectsEnabled || prefersReducedMotion.matches || !_selectionEffect || !_spinwheelTrack) {
		onComplete();
		return;
	}

	const runId = selectionEffectRun;
	const candidates = Array.from(new Set((names || []).map(name => String(name || '').trim()).filter(Boolean)));
	if (!candidates.includes(selectedName)) candidates.push(selectedName);
	const reelNames = [];
	// Every pass includes every candidate; the actual selection remains authoritative.
	const passes = Math.max(2, Math.ceil(24 / candidates.length));
	for (let pass = 0; pass < passes; pass += 1) {
		const shuffled = [...candidates];
		for (let index = shuffled.length - 1; index > 0; index -= 1) {
			const other = Math.floor(Math.random() * (index + 1));
			[shuffled[index], shuffled[other]] = [shuffled[other], shuffled[index]];
		}
		if (shuffled.length > 1 && shuffled[0] === reelNames[reelNames.length - 1]) {
			[shuffled[0], shuffled[1]] = [shuffled[1], shuffled[0]];
		}
		reelNames.push(...shuffled);
	}
	if (reelNames[reelNames.length - 1] !== selectedName) reelNames.push(selectedName);

	_spinwheelTrack.replaceChildren(...reelNames.map(name => {
		const item = document.createElement('div');
		item.className = 'spinwheel-item';
		item.innerText = getFirstName(name);
		return item;
	}));
	_selectionEffect.closest('.repetition-stage').classList.add('selection-running');
	document.getElementById('res').inert = true;
	_selectionEffect.setAttribute('aria-hidden', 'false');
	_selectionEffect.classList.add('is-active', 'is-spinning');
	const rowHeight = _spinwheelTrack.firstElementChild.offsetHeight;
	let delay = 0;
	for (let index = 1; index < reelNames.length; index += 1) {
		delay += 40 + 150 * Math.pow(index / reelNames.length, 5);
		queueSelectionEffect(() => {
			_spinwheelTrack.style.transform = `translateY(-${index * rowHeight}px)`;
		}, delay, runId);
	}
	queueSelectionEffect(() => {
		clearSelectionEffect();
		onComplete();
	}, delay + 450, runId);
}

function getFirstName(name) {
	return String(name || '').trim().split(/\s+/)[0] || '';
}

function createConfetti() {
	if (!_confettiStage) return;
	_confettiStage.innerHTML = '';
	const colors = ['#ffffff', '#93c5fd', '#818cf8', '#c4b5fd', '#67e8f9', '#dbeafe'];
	for (let index = 0; index < 48; index += 1) {
		const piece = document.createElement('i');
		const fromLeft = index % 2 === 0;
		const direction = fromLeft ? 1 : -1;
		piece.className = `confetti-piece ${fromLeft ? 'confetti-left' : 'confetti-right'}`;
		piece.style.setProperty('--confetti-color', colors[index % colors.length]);
		piece.style.setProperty('--burst-x', `${direction * (4 + Math.random() * 27)}rem`);
		piece.style.setProperty('--burst-y', `${-(9 + Math.random() * 19)}rem`);
		piece.style.setProperty('--fall-x', `${direction * (8 + Math.random() * 34)}rem`);
		piece.style.setProperty('--fall-y', `${8 + Math.random() * 17}rem`);
		piece.style.setProperty('--confetti-rotation', `${direction * (360 + Math.random() * 900)}deg`);
		piece.style.setProperty('--confetti-mid-rotation', `${direction * (180 + Math.random() * 360)}deg`);
		piece.style.setProperty('--confetti-delay', `${Math.random() * 180}ms`);
		piece.style.setProperty('--confetti-width', `${0.34 + Math.random() * 0.34}rem`);
		piece.style.setProperty('--confetti-height', `${0.65 + Math.random() * 0.55}rem`);
		_confettiStage.appendChild(piece);
	}
}

function queueSelectionEffect(callback, delay, runId) {
	const timer = setTimeout(() => {
		if (runId === selectionEffectRun) callback();
	}, motionDelay(delay));
	selectionEffectTimers.push(timer);
}

function clearSelectionEffect() {
	selectionEffectRun += 1;
	selectionEffectTimers.splice(0).forEach(timer => clearTimeout(timer));
	document.getElementById('res').inert = false;
	if (!_selectionEffect) return;
	_selectionEffect.closest('.repetition-stage').classList.remove('selection-running');
	_selectionEffect.classList.remove('is-active', 'is-spinning', 'is-rolling', 'is-celebrating', 'spotlights-on');
	_selectionEffect.setAttribute('aria-hidden', 'true');
	if (_spinwheelTrack) {
		_spinwheelTrack.style.transform = 'translateY(0)';
	}
}

function runCompletionEffect() {
	clearSelectionEffect();
	clearCompletionEffect();
	if (!visualEffectsEnabled || prefersReducedMotion.matches || !_completionEffect) return;
	const runId = completionEffectRun;
	_completionEffect.setAttribute('aria-hidden', 'false');
	_completionEffect.classList.add('is-active');
	saveReceipt.start();
	const timer = setTimeout(() => {
		if (runId === completionEffectRun) clearCompletionEffect();
	}, motionDelay(1350));
	completionEffectTimers.push(timer);
}

function applyTheme(theme) {
	document.documentElement.dataset.theme = theme === 'classic' ? 'classic' : 'midnight';
}

function clearCompletionEffect() {
	saveReceipt.stop();
	completionEffectRun += 1;
	completionEffectTimers.splice(0).forEach(timer => clearTimeout(timer));
	if (!_completionEffect) return;
	_completionEffect.classList.remove('is-active');
	_completionEffect.setAttribute('aria-hidden', 'true');
}

function updateVisualEffects(isEnabled) {
	const enabled = !!isEnabled;
	visualEffectsEnabled = enabled;
	if (_container) _container.classList.toggle('logo-animation-enabled', enabled);
	if (_title) {
		_title.classList.toggle('logo-animation-enabled', enabled);
		if (enabled) {
			_title.classList.remove('logo-animation-kick');
			void _title.offsetWidth;
			_title.classList.add('logo-animation-kick');
		} else {
			_title.classList.remove('logo-animation-kick');
		}
	}
	if (!enabled) {
		clearSelectionEffect();
		startRepetitionTiming();
		clearCompletionEffect();
	}
}

function animateMinimize() {
	if (!_container) {
		ipcRenderer.send('minimize');
		return;
	}

	uiMotion.minimizeWindow(_container, () => ipcRenderer.send('minimize'));
}

function animateRestore() {
	if (!_container) return;
	uiMotion.restoreWindow(_container);
}

// scale the name to fit the screen
function scaleName() {
	const availableWidth = _name.parentElement.clientWidth;
	if (!availableWidth) return;
	_nameSize = 5.5;
	_name.style.fontSize = _nameSize + 'rem';
	while (_name.scrollWidth > availableWidth && _nameSize > 2.4) {
		_nameSize = Math.max(2.4, _nameSize - 0.2);
		_name.style.fontSize = _nameSize + 'rem';
	}
}

const nameStageObserver = new ResizeObserver(entries => {
	if (entries[0].contentRect.width > 0) scaleName();
});
nameStageObserver.observe(_name.parentElement);

Math.clamp = function(a, b, c) {
	return Math.max(b, Math.min(c, a));
}

document.querySelectorAll('#title-bar .ico').forEach(control => {
	control.addEventListener('keydown', event => {
		if (event.key !== 'Enter' && event.key !== ' ') return;
		event.preventDefault();
		control.click();
	});
});

// automatically scale the name on load
module.exports = [scaleName(), updateState(), version(), ipcRenderer.send('get-update-news'), ipcRenderer.send('get-debug-mode'), ipcRenderer.send('get-ui-settings'), ipcRenderer.send('load-saved-file'), ipcRenderer.send('get-pending-excel-status'), ipcRenderer.send('get-undo-status'), ipcRenderer.send('get-active-teaching-session')];
