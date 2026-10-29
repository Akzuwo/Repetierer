const { ipcRenderer } = require('electron');
const motion = require('./uiMotion');
const KEY = 'repetierer-debug-scenario';
let installed = false, enabled = false, simulation = false, panel, launcher, current = '';
const people = Array.from({length: 25}, (_, i) => ({id: i + 6, name: ['Anna Müller', 'Leon Weber', 'Mia Keller', 'Noah Fischer', 'Emma Meier'][i % 5] + (i > 4 ? ` ${Math.floor(i / 5) + 1}` : ''), grades: 2, joker: 3, absent: false}));
const session = () => ({id: 'debug-session', className: 'Demo A', presentStudents: people, absentStudents: [], participationEntries: [], repetitionEntries: []});
const settings = {theme: 'midnight', probabilityDecreaseFactor: .8, neverSelectedBoostFactor: 1.5, visualEffectsEnabled: true, wiggersRuleDurationMinutes: 30, wiggersRuleEnabled: true};
const emit = (channel, ...args) => ipcRenderer.emit(channel, {}, ...args);
const later = callback => setTimeout(callback, 30);
const click = id => document.getElementById(id)?.click();

// Install before the real renderer starts. In a simulation every outgoing IPC
// call is handled locally, including unknown future write actions. Reload is
// the only way back to live IPC, so a forgotten preview cannot write to Excel.
function isolate() {
  if (simulation) return;
  simulation = true;
  document.documentElement.dataset.debugSimulation = 'true';
  const nativeSend = ipcRenderer.send.bind(ipcRenderer);
  ipcRenderer.send = (channel, ...args) => {
    if (['quit', 'minimize'].includes(channel)) return nativeSend(channel, ...args);
    later(() => {
      if (channel === 'get-debug-mode') emit('debug-mode', true);
      else if (channel === 'version') emit('version', '7.0.0 · Debug');
      else if (channel === 'get-settings') emit('settings-data', settings, {settingsPath: 'Demo / settings.json'}, {hasFile: true});
      else if (channel === 'get-ui-settings') emit('ui-settings', settings);
      else if (channel === 'class') emit('ready', false);
      else if (channel === 'get-session-setup') emit('session-setup-data', {className: 'Demo A', persons: people});
      else if (channel === 'start-teaching-session') emit('teaching-session-started', session());
      else if (channel === 'end-teaching-session') emit('teaching-session-ended');
      else if (channel === 'add-participation-entry') emit('participation-entry-saved', {...session(), participationEntries: [args[0]]});
      else if (channel === 'start') emit('name', [people[0].name, 3], people.map(p => p.name), people[0]);
      else if (channel === 'run-joker-date-migration') emit('joker-date-migration-result', current === 'migration-locked' ? {success: false, reason: 'excel-locked'} : {success: true, dates: 18, backupPath: 'Demo / Klasse.backup.xlsx'});
      else if (channel === 'update-download-approved') emit('update-status', {status: 'progress', version: '99.9.9', percent: 62});
      else if (channel === 'get-class-statistics') emit('class-statistics-data', []);
      else if (channel === 'get-pending-excel-status') emit('pending-excel-status', 3);
      // All other sends intentionally remain inside the preview.
    });
  };
  ipcRenderer.invoke = async (channel, args) => {
    if (channel === 'ai:status') return {state: 'ready', message: 'Debug · lokale Simulation'};
    if (channel === 'ai:ask') {
      if (current === 'ai-loading') return new Promise(() => {});
      if (current === 'ai-error') throw new Error('Simulation: Die lokale KI ist gerade nicht erreichbar.');
      if (current === 'ai-confirmation') return {type: 'confirmation', token: 'debug-only', title: 'Mitarbeit erfassen?', lines: ['Anna Müller · Demo A', '4 Punkte für den heutigen Beitrag']};
      return {type: 'answer', text: 'Anna hat heute einen passenden Beitrag mit erkennbarem Mitdenken geleistet. Die Änderung wurde nur in der Vorschau simuliert.'};
    }
    if (channel === 'ai:confirm') return {type: 'answer', text: args.approved ? 'Vorschau bestätigt. Es wurden keine Daten verändert.' : 'Vorschau abgebrochen.'};
    if (channel === 'assessment:get-overview') return {ok: true, data: {name: 'Demo A', rounds: []}};
    if (channel === 'assessment:config-status' || channel === 'assessment:brevo-status') return {ok: true, data: {configured: false}};
    return {ok: false, error: {message: 'Debug-Vorschau: Diese Aktion wird nicht ausgeführt.'}};
  };
}

function modalLabel(modal) { return modal.querySelector('h1,h2,h3')?.textContent.trim() || modal.id; }
function seedClasses() { emit('classes', ['Demo A', 'Demo B', 'Demo C', 'Demo D', 'Demo E'], 'Demo / Excel Template.xlsx'); }
function seedMigration() {
  emit('joker-date-migration-required', {fingerprint: 'debug-only', sheets: [{name: 'Demo A', needsMigration: true, columns: 6, dates: 18}, {name: 'Demo B', needsMigration: true, columns: 6, dates: 9, keepsLegacyColumn: true}]}, 'Demo / Klasse.xlsx');
}
function scenarioList() {
  return [
    ['Pages', 'classes', 'Klassenauswahl'], ['Pages', 'attendance', 'Abwesenheiten'], ['Pages', 'participation', 'Session · Mitarbeit'], ['Pages', 'repetition', 'Session · Repetition'], ['Pages', 'settings', 'Einstellungen'], ['Pages', 'assessment', 'Beurteilungsrunden'], ['Pages', 'statistics', 'Statistiken · leer'],
    ['Updates & Meldungen', 'update-available', 'Update verfügbar'], ['Updates & Meldungen', 'update-progress', 'Update · Download'], ['Updates & Meldungen', 'update-error', 'Update · Fehler'], ['Updates & Meldungen', 'update-installed', 'Update installiert'], ['Updates & Meldungen', 'notification', 'Benachrichtigung'], ['Updates & Meldungen', 'pending', 'Excel · Speichern ausstehend'],
    ['KI-Assistent', 'ai-welcome', 'Chat · Begrüssung'], ['KI-Assistent', 'ai-loading', 'Chat · Antwort wird erstellt'], ['KI-Assistent', 'ai-answer', 'Chat · Antwort schreiben'], ['KI-Assistent', 'ai-confirmation', 'Chat · Änderungsbestätigung'], ['KI-Assistent', 'ai-error', 'Chat · Verbindungsfehler'],
    ['Datenmigration', 'migration-preview', 'Joker-Datum · Vorschau'], ['Datenmigration', 'migration-locked', 'Joker-Datum · Datei gesperrt'], ['Datenmigration', 'migration-success', 'Joker-Datum · Abgeschlossen'],
    ...[...document.querySelectorAll('.modal:not(#debug-panel)')].filter(modal => modal.id !== 'problem-report-modal' || require('../src/featureFlags').problemReportsEnabled).map(modal => ['Alle Modals', `modal:${modal.id}`, modalLabel(modal)])
  ];
}
function choose(id) {
  if (!enabled) return;
  sessionStorage.setItem(KEY, id);
  // A clean document for each fixture resets handlers, busy states and timers.
  location.reload();
}
function reset() { sessionStorage.removeItem(KEY); location.reload(); }
function closePanel() { motion.closeDialog(panel); launcher.focus(); }
function openPanel() { motion.openDialog(panel); panel.querySelector('input').focus(); }

function createPanel() {
  launcher = document.createElement('button'); launcher.id = 'debug-launcher'; launcher.className = 'btn-2 debug-launcher';
  launcher.textContent = simulation ? 'Debug · Simulation' : 'Debug'; launcher.addEventListener('click', openPanel); document.body.append(launcher);
  panel = document.createElement('div'); panel.id = 'debug-panel'; panel.className = 'modal'; panel.setAttribute('role', 'dialog'); panel.setAttribute('aria-modal', 'true'); panel.setAttribute('aria-labelledby', 'debug-heading');
  panel.innerHTML = '<div class="modal-content debug-panel-content"><button class="debug-dismiss update-hidden" type="button" aria-label="Debug schliessen">Schliessen</button><h2 id="debug-heading">Debug-Szenarien</h2><p>Modals und App-Zustände mit Beispieldaten ansehen. In einer Simulation werden keine Dateien verändert, Updates installiert oder Nachrichten versendet.</p><div class="debug-panel-actions"><input type="search" placeholder="Szenario suchen …" aria-label="Szenario suchen"><button class="btn-2" id="debug-reset">Simulation beenden</button></div><div class="debug-scenario-list"></div></div>';
  document.body.append(panel);
  panel.querySelector('.debug-dismiss').addEventListener('click', closePanel);
  panel.querySelector('#debug-reset').addEventListener('click', reset);
  panel.addEventListener('click', e => {if (e.target === panel) closePanel();});
  const search = panel.querySelector('input'), list = panel.querySelector('.debug-scenario-list');
  function render() {
    list.replaceChildren(); const groups = new Map();
    for (const [group, id, label] of scenarioList()) {
      if (!`${group} ${label}`.toLocaleLowerCase('de').includes(search.value.toLocaleLowerCase('de'))) continue;
      if (!groups.has(group)) {
        const section = document.createElement('section'), heading = document.createElement('h3'), buttons = document.createElement('div'); heading.textContent = group; section.append(heading, buttons); list.append(section); groups.set(group, buttons);
      }
      const button = document.createElement('button'); button.className = 'btn-2'; button.dataset.debugScenario = id; button.textContent = label; button.addEventListener('click', () => choose(id)); groups.get(group).append(button);
    }
    if (!list.children.length) list.textContent = 'Keine passenden Szenarien.';
  }
  search.addEventListener('input', render); render();
  document.getElementById('debug-scenarios-btn').addEventListener('click', () => {motion.setDrawerOpen(false); openPanel();});
  document.addEventListener('keydown', e => {
    if (e.ctrlKey && e.shiftKey && e.key.toLowerCase() === 'd') {e.preventDefault(); openPanel();}
    if (!panel.classList.contains('modal-open')) return;
    if (e.key === 'Escape') {e.preventDefault(); e.stopImmediatePropagation(); closePanel();}
    if (e.key === 'Tab') {
      const controls = [...panel.querySelectorAll('button,input')].filter(node => node.getClientRects().length), first = controls[0], last = controls.at(-1);
      if (e.shiftKey && document.activeElement === first) {e.preventDefault(); last.focus();}
      else if (!e.shiftKey && document.activeElement === last) {e.preventDefault(); first.focus();}
    }
  }, true);
}
async function runScenario() {
  if (!simulation || !current) return;
  seedClasses();
  if (current === 'classes') return;
  if (current === 'attendance') { emit('session-setup-data', {className: 'Demo A', persons: people}); return; }
  if (['participation', 'repetition'].includes(current)) {emit('teaching-session-started', session()); if(current === 'repetition') {await waitForSession(); click('session-repetition-tab');} return;}
  if (current === 'assessment') { click('assessment-nav-btn'); return; }
  if (current === 'statistics') { click('stats-view-btn'); return; }
  if (current === 'settings' || current === 'modal:settings-modal') { click('settings-btn'); return; }
  if (current.startsWith('update-')) {
    const status = current.slice(7); emit('update-status', {status: status === 'progress' ? 'progress' : status, version: '99.9.9', percent: 62, message: status === 'error' ? 'Simulation: Download unterbrochen. Bitte erneut versuchen.' : undefined}); return;
  }
  if (current === 'notification') {click('debug-test-alert-btn'); return;}
  if (current === 'pending') {emit('pending-excel-status', 3); emit('excel-write-pending', {type: 'grade'}); return;}
  if (current.startsWith('migration-')) {
    seedMigration();
    if (current === 'migration-locked') emit('joker-date-migration-result', {success:false, reason:'excel-locked'});
    if (current === 'migration-success') emit('joker-date-migration-result', {success:true, dates:27, backupPath:'Demo / Klasse.backup.xlsx'});
    return;
  }
  if (current.startsWith('ai-')) {
    click('ai-open-btn');
    if (current !== 'ai-welcome') {await new Promise(resolve => setTimeout(resolve, 150)); const input = document.getElementById('ai-input'); input.value = 'Erfasse 4 Mitarbeitspunkte für Anna Müller.'; input.dispatchEvent(new Event('input')); document.getElementById('ai-form').requestSubmit();}
    return;
  }
  if (current.startsWith('modal:')) {
    const id = current.slice(6);
    if (id === 'joker-date-migration-modal') {seedMigration(); return;}
    if (id === 'settings-modal') {click('settings-btn'); return;}
    if (id === 'ai-modal') {click('ai-open-btn'); return;}
    if (id === 'person-modal') {emit('persons-list', people); return;}
    if (id === 'absence-modal') {emit('absence-persons-data', people); return;}
    if (id === 'backup-modal') {emit('backup-data', {entries: [{type:'grade', grade:5, personName:'Anna Müller', className:'Demo A', createdAt:new Date().toISOString(), excelWriteSucceeded:true, filePath:'Demo / Klasse.xlsx'}]}, {backupPath:'Demo / backups.json'}); return;}
    if (id === 'log-modal') {emit('log-data', 'DEBUG · Klassen geladen\nDEBUG · Session gestartet\nDEBUG · Excel-Speicherung simuliert', {logPath:'Demo / app.log'}); return;}
    if (id === 'probabilities-modal') {emit('probability-data', people.map(person=>({...person, probability:1/people.length, weight:1}))); return;}
    if (id === 'class-import-modal') {emit('class-list-import-preview', {detectedClassName:'Demo F', count:people.length, names:people.map(p=>p.name)}); return;}
    if (id === 'joker-migration-modal') {emit('joker-migration-required', {persons:25}, 'Demo / Klasse.xlsx'); return;}
    if (id === 'excel-editor-modal') {document.getElementById('class-name').textContent='Demo A'; emit('editor-persons-data', people); return;}
    if (id === 'problem-report-modal') {click('problem-report-btn'); emit('report-preview-data', {appId:'Repetierer', version:'99.9.9', environment:'debug', type:'Problembericht', consoleOutput:'Nur simulierte Einträge.', metadata:{os:'Windows', clientTimestamp:new Date().toISOString()}}); return;}
    if (id === 'assessment-confirm-modal') document.getElementById('assessment-confirm-text').textContent='Beurteilungen für 25 Schüler versenden? Dies ist eine lokale Vorschau.';
    if (id === 'assessment-delete-modal') document.getElementById('assessment-delete-text').textContent='Beurteilungsrunde «Herbstsemester 2026» löschen?';
    if (id === 'participation-modal' || id === 'session-end-modal') {
      emit('teaching-session-started', session());
      await waitForSession();
      id === 'participation-modal' ? document.querySelector('.participation-student')?.click() : click('end-session-btn'); return;
    }
    if (id === 'update-news-modal') {emit('update-news-data', {version:'99.9.9', notes:'Debug-Vorschau\n- Flüssigere Seitenwechsel\n- Joker-Datumsspalten und Datenmigration'}); return;}
    if (id === 'reload-conflict-modal') {emit('reload-excel-conflict', 3); return;}
    const modal = document.getElementById(id); if (modal) motion.openDialog(modal);
  }
}
async function waitForSession() {
  for (let i=0;i<100;i++) {
    if (document.querySelector('#session-header:not(.update-hidden)') && !document.getElementById('content').dataset.pageMotion) return;
    await new Promise(resolve=>setTimeout(resolve,20));
  }
  throw new Error('Die simulierte Session konnte nicht geöffnet werden.');
}
function authorize(value) {
  if (!value || enabled) return;
  enabled = true; current = sessionStorage.getItem(KEY) || '';
  if (current) isolate();
  // DOM modules finish registering their real handlers before fixtures arrive.
  setTimeout(() => {createPanel(); document.getElementById('debug-actions').classList.remove('update-hidden'); runScenario().catch(console.error);}, 80);
}
function install() {
  if (installed) return; installed = true;
  if (process.argv.includes('--repetierer-debug')) {current = sessionStorage.getItem(KEY) || ''; if (current) isolate();}
  ipcRenderer.on('debug-mode', (_event, value) => authorize(value));
}
module.exports = {install};
