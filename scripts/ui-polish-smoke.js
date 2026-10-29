// Exercise the real renderer in an isolated, hidden Electron window with controlled IPC.
// Run: npm run test:ui
const { app, BrowserWindow, ipcMain } = require('electron');
if (!app) {
	const env = { ...process.env };
	delete env.ELECTRON_RUN_AS_NODE;
	const result = require('node:child_process').spawnSync(require('electron'), [__filename], { env, stdio: 'inherit', windowsHide: true });
	if (result.error) console.error(result.error);
	process.exit(result.status ?? 1);
}
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const output = path.join(__dirname, '..', '.ui-polish-smoke');
fs.mkdirSync(output, { recursive: true });
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'repetierer-ui-')));
let win, aiState = 'ready', askResolve, overviewResolve, failOverview = false, failConfirm = true, expiredConfirm = false;
let failTeachingStart = false;
let asks = 0, fileRequests = 0, importRequests = 0, minimizeRequests = 0;
const errors = [];
const students = [{ id: 6, name: 'Anna Müller' }, { id: 7, name: 'Max Beispiel' }];
const session = { id: 'ui-test', className: 'L23a', presentStudents: students, absentStudents: [], participationEntries: [], repetitionEntries: [] };
const pause = ms => new Promise(resolve => setTimeout(resolve, ms));
const ok = data => ({ ok: true, data });
const {AssessmentStore} = require('../src/assessment/assessmentStore');
const {presentClass,summarizeRound} = require('../src/assessment/mainAssessment');
const assessmentFile=path.join(app.getPath('userData'),'assessment-test.json');
const workbook=path.join(app.getPath('userData'),'class.xlsx');
const assessmentStore=new AssessmentStore(assessmentFile);
const commentRound=assessmentStore.createRound(workbook,'L23a',students,{title:'Kommentare je Kriterium'});
let showCommentRound=false;
function commentView() {
	const {classRecord,round}=assessmentStore.getRound(workbook,'L23a',students,commentRound.id);
	return {classRecord:presentClass(classRecord),round,students:classRecord.students,summary:summarizeRound(round,classRecord.students),invitation:null,assessmentConfigured:false,deliveryValidation:{valid:false,incomplete:[]}};
}

ipcMain.handle('ai:status', () => ({ state: aiState, message: aiState === 'ready' ? 'Bereit' : 'Lokale KI ist noch nicht eingerichtet.' }));
ipcMain.handle('ai:ask', () => { asks++; return new Promise(resolve => { askResolve = resolve; }); });
ipcMain.handle('ai:confirm', () => {
	if (expiredConfirm) throw new Error('Bestätigung ist ungültig oder abgelaufen.');
	if (failConfirm) { failConfirm = false; throw new Error('Test: Verbindung unterbrochen'); }
	return { type: 'answer', text: 'Die bestätigte Aktion wurde abgeschlossen.' };
});
ipcMain.handle('assessment:get-overview', () => new Promise(resolve => { overviewResolve = () => resolve(failOverview ? { ok: false, error: { message: 'Test: Daten nicht erreichbar' } } : ok(showCommentRound ? presentClass(assessmentStore.getClassContext(workbook,'L23a',students)) : { name: 'L23a', rounds: [] })); }));
ipcMain.handle('assessment:get-round',()=>ok(commentView()));
ipcMain.handle('assessment:save-teacher',(_event,args)=>{
	assessmentStore.saveTeacherAssessment(workbook,'L23a',students,args.roundId,args.studentId,args.answers,args.comment,args.criterionComments);
	return ok(commentView());
});
ipcMain.handle('assessment:config-status', () => ok({ configured: false }));
ipcMain.handle('assessment:brevo-status', () => ok({ configured: false }));
const bundleFixture = path.join(app.getPath('userData'),'shared.repetierer');
const bundleValues = {ASSESSMENT_API_URL:'https://example.ch/api',ASSESSMENT_TEACHER_TOKEN:'rt_'+'a'.repeat(43),BREVO_SMTP_HOST:'smtp.example.ch',BREVO_SMTP_PORT:'587',BREVO_SMTP_USER:'teacher@example.ch',BREVO_SMTP_PASSWORD:'fixture-only',BREVO_FROM_EMAIL:'teacher@example.ch',BREVO_FROM_NAME:'Demo'};
ipcMain.handle('assessment:config-export',()=>ok(require('../src/assessment/integrationConfig').exportIntegrationFile(bundleFixture,bundleValues)));
ipcMain.handle('assessment:config-import',async()=>{
	require('../src/assessment/integrationConfig').exportIntegrationFile(bundleFixture,bundleValues);
	return ok(await require('../src/assessment/integrationConfig').importIntegrationEnv(bundleFixture,path.join(app.getPath('userData'),'imported.env'),{protectTeacherToken:token=>'encrypted:'+Buffer.from(token).toString('base64')}));
});
ipcMain.on('load-saved-file', event => event.reply('saved-file-missing', ''));
ipcMain.on('get-active-teaching-session', event => event.reply('active-teaching-session-data', null));
ipcMain.on('version', event => event.reply('version', '7.0.0'));
ipcMain.on('get-settings', event => event.reply('settings-data', { theme: 'midnight', probabilityDecreaseFactor: .8, neverSelectedBoostFactor: 1.5, visualEffectsEnabled: false, wiggersRuleDurationMinutes: 30 }, {}, {}));
ipcMain.on('file', () => fileRequests++);
ipcMain.on('import-class-list', () => importRequests++);
ipcMain.on('minimize', event => { minimizeRequests++; event.reply('window-restored'); });
ipcMain.on('class', event => event.reply('ready', false));
ipcMain.on('get-session-setup', event => event.reply('session-setup-data', { className: 'L23a', persons: students }));
ipcMain.on('start-teaching-session', event => event.reply(failTeachingStart?'teaching-session-start-failed':'teaching-session-started', failTeachingStart?{reason:'test-retry'}:session));
ipcMain.on('end-teaching-session', event => event.reply('teaching-session-ended'));
ipcMain.on('start', event => event.reply('name', ['Anna Müller', 0], students.map(item => item.name), students[0]));

const evaluate = expression => win.webContents.executeJavaScript(expression, true);
async function waitFor(expression, label) {
	for (let i = 0; i < 100; i++) {
		if (await evaluate(expression)) return;
		await pause(30);
	}
	throw new Error(`Zeitüberschreitung: ${label || expression}`);
}
const click = id => evaluate(`document.getElementById(${JSON.stringify(id)}).click()`);
const visible = id => evaluate(`getComputedStyle(document.getElementById(${JSON.stringify(id)})).display !== 'none'`);
async function shot(name) {
	await pause(250);
	win.webContents.invalidate();
	await pause(60);
	fs.writeFileSync(path.join(output, `${name}.png`), (await win.webContents.capturePage()).toPNG());
}
async function draft(text) {
	await evaluate(`(() => { const input = document.getElementById('ai-input'); input.value = ${JSON.stringify(text)}; input.dispatchEvent(new Event('input', { bubbles: true })); })()`);
}
async function key(keyCode, modifiers = []) {
	win.webContents.sendInputEvent({ type: 'keyDown', keyCode, modifiers });
	if (keyCode === 'Enter' && modifiers.includes('shift')) win.webContents.sendInputEvent({ type: 'char', keyCode: '\r', modifiers });
	win.webContents.sendInputEvent({ type: 'keyUp', keyCode, modifiers });
	await pause(50);
}

async function verifyModalCenters() {
	const results = await evaluate(`(() => {
		const style = document.createElement('style');
		style.textContent = '.modal, .modal .modal-content { animation: none !important; transform: none !important; }';
		document.head.appendChild(style);
		const results = [...document.querySelectorAll('.modal')].map(modal => {
			const previousStyle = modal.getAttribute('style'), previousClass = modal.className;
			modal.style.display = 'flex';
			modal.classList.add('modal-open');
			const rect = modal.querySelector('.modal-content').getBoundingClientRect();
			const result = { id: modal.id, x: Math.abs(rect.left + rect.width / 2 - innerWidth / 2), y: Math.abs(rect.top + rect.height / 2 - innerHeight / 2), inside: rect.left >= 0 && rect.top >= 0 && rect.right <= innerWidth && rect.bottom <= innerHeight };
			if (previousStyle === null) modal.removeAttribute('style'); else modal.setAttribute('style', previousStyle);
			modal.className = previousClass;
			return result;
		});
		style.remove();
		return results;
	})()`);
	assert.ok(results.length >= 20, 'All existing dialogs are covered; settings is now a page');
	assert.deepEqual(results.filter(item => item.x > 1 || item.y > 1 || !item.inside), [], 'Dialog boxes are centered and remain in the viewport');
	return results.length;
}

async function orbSnapshot() {
	return evaluate(`(() => {
		const canvas = document.getElementById('ai-thinking-orb');
		const pixels = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
		return { state: canvas.dataset.orbState, motion: canvas.dataset.orbMotion, ink: pixels.some((value, index) => index % 4 === 3 && value > 0), frame: canvas.toDataURL(), visibility: document.visibilityState };
	})()`);
}

async function main() {
	await app.whenReady();
	win = new BrowserWindow({ width: 1100, height: 800, frame: false, show: false, backgroundColor: '#05070b', webPreferences: { nodeIntegration: true, contextIsolation: false, spellcheck: false, backgroundThrottling: false, offscreen: true } });
	win.webContents.on('console-message', event => {
		const level = event.level;
		const text = event.message;
		if ((level === 'error' || level === 3) && !String(text).includes('Test:')) errors.push(text);
	});
	await win.loadFile(path.join(__dirname, '..', 'public', 'index.html'));
	await waitFor(`document.getElementById('session-file-empty') && !document.getElementById('ai-input').disabled`, 'Rendererstart');
	await waitFor(`Number(getComputedStyle(document.getElementById('container')).opacity) === 1`, 'Appanimation');
	await evaluate(`document.fonts.ready`);
	assert.match(await evaluate(`getComputedStyle(document.getElementById('session-start-heading')).fontFamily`), /Segoe UI/, 'Kristall uses the native system font');
	await waitFor(`document.documentElement.dataset.renderer === 'webgl'`, 'Kristall WebGL start');
	const modalCount = await verifyModalCenters();
 await win.webContents.debugger.attach('1.3');
 await win.webContents.debugger.sendCommand('DOM.enable'); await win.webContents.debugger.sendCommand('CSS.enable');
 const dom=await win.webContents.debugger.sendCommand('DOM.getDocument');const gearNode=await win.webContents.debugger.sendCommand('DOM.querySelector',{nodeId:dom.root.nodeId,selector:'#settings-btn'});
 const gearCenter=await evaluate(`(()=>{const r=document.querySelector('#settings-btn svg').getBoundingClientRect();return [r.x+r.width/2,r.y+r.height/2]})()`);
 await win.webContents.debugger.sendCommand('CSS.forcePseudoState',{nodeId:gearNode.nodeId,forcedPseudoClasses:['hover']});await pause(160);
 const hoverCenter=await evaluate(`(()=>{const r=document.querySelector('#settings-btn svg').getBoundingClientRect();return [r.x+r.width/2,r.y+r.height/2]})()`);
 assert.ok(Math.hypot(gearCenter[0]-hoverCenter[0],gearCenter[1]-hoverCenter[1])<.1,'Gear rotates without moving its centre');
 await win.webContents.debugger.sendCommand('CSS.forcePseudoState',{nodeId:gearNode.nodeId,forcedPseudoClasses:[]});win.webContents.debugger.detach();
	assert.equal(await evaluate(`document.documentElement.dataset.motionEngine`), 'gsap');
	assert.equal(await evaluate(`require('./vendor/ui-motion.js').gsap.version`), '3.15.0');

	// Drawer transitions survive an interrupted exit and never leave hidden controls focusable.
	assert.equal(await evaluate(`(()=>{document.getElementById('drawer-toggle').click();return require('./vendor/ui-motion.js').gsap.getTweensOf(document.getElementById('drawer')).length > 0;})()`), true, 'The actual drawer transition uses GSAP');
 await pause(30);
	await click('drawer-toggle');
	await pause(30);
	await click('drawer-toggle');
	await waitFor(`!document.getElementById('drawer').dataset.motionState`);
	assert.equal(await evaluate(`document.getElementById('drawer').classList.contains('drawer-open') && !document.getElementById('drawer').inert && Math.abs(document.getElementById('drawer').getBoundingClientRect().left) < 2`), true, 'Latest drawer action wins');
	await click('drawer-toggle');
	await waitFor(`!document.getElementById('drawer').dataset.motionState`);
	assert.equal(await evaluate(`document.getElementById('drawer').inert && document.getElementById('drawer').getBoundingClientRect().right < 0`), true);
	await click('minimize');
	await pause(450);
	assert.equal(minimizeRequests, 1, 'Minimize still reaches Electron IPC after its transition');
	await waitFor(`getComputedStyle(document.getElementById('container')).opacity === '1'`, 'Window restoration completes');
	assert.equal(await evaluate(`getComputedStyle(document.getElementById('container')).opacity`), '1', 'Window restoration clears the minimized visual state');
	assert.equal(await visible('session-file-empty'), true);
	assert.equal(await visible('repetition'), false);
	await click('session-open-file-btn');
	await click('session-import-btn');
	await waitFor('true');
	await pause(30);
	assert.equal(fileRequests, 1);
	assert.equal(importRequests, 1);
	await shot('home-empty');
	assert.equal(await evaluate(`document.getElementById('session-open-file-btn').dataset.glass`), 'pill');
	assert.equal(await evaluate(`getComputedStyle(document.getElementById('session-open-file-btn')).backgroundColor`), 'rgba(0, 0, 0, 0)', 'WebGL owns the glass fill');
	win.webContents.send('classes', ['L23a', 'M24b', 'K25c', 'B26a', 'B26b', 'L24a'], 'C:\\Beispiel\\Unterricht.xlsx');
	await waitFor(`document.querySelectorAll('#class-list button').length === 6`);
	assert.equal(await visible('session-file-empty'), false);
	assert.match(await evaluate(`document.getElementById('session-source-label').textContent`), /6 Klassen · Unterricht.xlsx/);
	await shot('home-classes');
	assert.notDeepEqual(fs.readFileSync(path.join(output, 'home-empty.png')), fs.readFileSync(path.join(output, 'home-classes.png')), 'Screenshots reflect current renderer state');
	await evaluate(`document.querySelector('#class-list button').click()`);
	await waitFor(`!document.getElementById('session-attendance-view').classList.contains('update-hidden')`);
	await shot('session-attendance');
 assert.equal(await evaluate(`document.querySelectorAll('#session-attendance-view input[type=checkbox]').length`),0);
 assert.equal(await evaluate(`document.documentElement.dataset.crystalOverlay`),'false');
 await evaluate(`document.querySelector('.attendance-person').click()`);
 assert.equal(await evaluate(`document.querySelector('.attendance-person').getAttribute('aria-pressed')`),'true');
 assert.match(await evaluate(`document.getElementById('session-absent-preview').textContent`),/1 abwesend/);
 await evaluate(`document.querySelector('.attendance-person').click()`);
	await click('confirm-session-start-btn');
	await waitFor(`document.querySelectorAll('.participation-student').length === 2`);
	await pause(400);
	await shot('session-participation');
 await evaluate(`document.querySelector('.participation-student').click()`); await pause(350);
 assert.equal(await evaluate(`document.querySelectorAll('.participation-score').length`), 6);
 assert.ok(await evaluate(`[...document.querySelectorAll('.participation-score')].every(button=>{const a=button.getBoundingClientRect();return [...button.children].every(c=>{const b=c.getBoundingClientRect();return b.left>=a.left && b.right<=a.right+1 && b.top>=a.top && b.bottom<=a.bottom+1;}) && getComputedStyle(button.querySelector('strong')).backgroundColor==='rgba(0, 0, 0, 0)' && +getComputedStyle(button.querySelector('strong')).fontWeight>=700;})`));
 await shot('participation-entry-refined'); await click('participation-cancel-btn'); await pause(250);
 assert.equal(await evaluate(`document.querySelector('.participation-student').dataset.glass`), 'pill');
 assert.equal(await evaluate(`document.getElementById('grade-input').hasAttribute('data-glass')`), false);
 assert.ok(await evaluate(`document.getElementById('end-session-btn').getBoundingClientRect().top < 130`));
	await click('session-repetition-tab');
	await waitFor(`document.getElementById('session-repetition-tab').getAttribute('aria-pressed') === 'true'`);
	assert.equal(await evaluate(`document.querySelectorAll('#class .neum, #class .session-file-panel, #class .session-classes-panel').length`), 0, 'Sidebar cards were removed from the markup');
	assert.equal(await evaluate(`(() => { const styles = [...document.querySelector('#class').children].map(item => getComputedStyle(item)); return styles.every(style => style.backgroundColor === 'rgba(0, 0, 0, 0)' && style.boxShadow === 'none' && parseFloat(style.borderTopWidth) === 0); })()`), true, 'File and classes remain unboxed');
	assert.ok(await evaluate(`parseFloat(getComputedStyle(document.querySelector('#class h2')).fontSize)`) < 16, 'Context labels remain secondary to the repetition heading');
	assert.equal(await visible('class'), false, 'The whole left toolbar is hidden during a session');
	assert.equal(await evaluate(`document.querySelectorAll('.session-live-dot, #session-attendance-count').length`), 0);
	assert.equal(await evaluate(`getComputedStyle(document.getElementById('repetition')).gridColumnStart`), '1', 'Repetition occupies the full session width');
	assert.equal(await visible('res'), false);
	win.webContents.debugger.attach('1.3');
	const {root} = await win.webContents.debugger.sendCommand('DOM.getDocument');
	const {nodeId} = await win.webContents.debugger.sendCommand('DOM.querySelector', {nodeId: root.nodeId, selector: '#repetition-help-btn'});
	await win.webContents.debugger.sendCommand('CSS.enable');
	await win.webContents.debugger.sendCommand('CSS.forcePseudoState', {nodeId, forcedPseudoClasses: ['hover']});
	await pause(200);
	assert.equal(await evaluate(`getComputedStyle(document.getElementById('repetition-help-btn')).transform`), 'none', 'Question marks stay still on hover');
	assert.equal(await evaluate(`getComputedStyle(document.getElementById('repetition-help-btn')).backgroundColor`), 'rgba(0, 0, 0, 0)', 'Help hover preserves the clear crystal material');
	assert.notEqual(await evaluate(`getComputedStyle(document.getElementById('repetition-help-btn')).boxShadow`), 'none');
	await win.webContents.debugger.sendCommand('CSS.forcePseudoState', {nodeId, forcedPseudoClasses: []});
	win.webContents.debugger.detach();
	await shot('session-repetition-ready');
 assert.equal(await evaluate(`document.getElementById('end-session-btn').dataset.glass`),'pill');
 assert.ok(await evaluate(`document.getElementById('end-session-btn').getBoundingClientRect().bottom <= document.getElementById('repetition').getBoundingClientRect().top`));
 assert.ok(await evaluate(`parseFloat(getComputedStyle(document.querySelector('.repetition-header h2')).fontSize) >= 28`));
 assert.ok(await evaluate(`['start-btn','manual-select-btn','probabilities-btn','undo-last-btn','redo-last-btn','repetition-help-btn'].every(id => Math.abs(document.getElementById(id).getBoundingClientRect().height-48)<1)`));
	await click('start-btn');
	await waitFor(`document.getElementById('name').textContent.includes('Anna')`);
	assert.equal(await visible('res'), true);
	assert.ok(await evaluate(`parseFloat(getComputedStyle(document.getElementById('name')).fontSize)`) >= 80, 'Short student names retain their prominent display size');
	win.webContents.send('name', ['Alexandermatthias Beispiel', 0], [], students[0]);
	await waitFor(`document.getElementById('name').textContent.includes('Alexandermatthias')`);
	assert.equal(await evaluate(`document.getElementById('name').scrollWidth <= document.querySelector('.repetition-stage').clientWidth + 1`), true, 'Long student names fit their actual stage');
	win.setSize(680, 560);
	await pause(100);
	assert.equal(await evaluate(`document.getElementById('name').scrollWidth <= document.querySelector('.repetition-stage').clientWidth + 1`), true, 'Long names rescale when the window becomes compact');
	win.setSize(1100, 800);
	win.webContents.send('name', ['Anna Müller', 0], [], students[0]);
	await waitFor(`document.getElementById('name').textContent.includes('Anna')`);
	await shot('session-repetition');
	win.setSize(680, 560);
	await pause(300);
	await shot('session-repetition-compact');
	assert.equal(await evaluate(`(() => { const rect = document.getElementById('res').getBoundingClientRect(); return rect.right <= innerWidth && rect.bottom <= innerHeight; })()`), true);
	assert.equal(await evaluate(`document.getElementById('grade-input').placeholder`), 'z.B. 5.5');
	assert.equal(await evaluate(`(() => { const controls = [...document.querySelectorAll('#res button, #grade-input')]; return controls.every(control => control.getBoundingClientRect().height >= 52 && parseFloat(getComputedStyle(control).fontSize) >= 17 && +getComputedStyle(control).fontWeight >= 600); })()`), true, 'Grading controls are larger and bolder, including at compact sizes');
	assert.equal(await evaluate(`(() => { const name = document.getElementById('name').getBoundingClientRect(), res = document.getElementById('res').getBoundingClientRect(); return name.bottom <= res.top; })()`), true, 'The enlarged grading bar does not overlap the student name');
	win.setSize(1100, 800);
	win.webContents.send('ui-settings', { theme: 'midnight', visualEffectsEnabled: true });
	await pause(50);
	const rouletteNames = ['Anna Müller', 'Anna Meier', 'Max Beispiel', 'Nora Test'];
	win.webContents.send('name', ['Anna Müller', 0], rouletteNames, students[0]);
	await waitFor(`document.querySelector('.repetition-stage').classList.contains('selection-running')`);
	assert.equal(await evaluate(`getComputedStyle(document.getElementById('name')).visibility`), 'hidden', 'The selected name cannot leak behind the roulette');
	assert.equal(await evaluate(`document.getElementById('res').inert`), true, 'Grading waits until the roulette stops');
	const reel = await evaluate(`[...document.querySelectorAll('#spinwheel-track > div')].map(item => item.textContent)`);
	for (let offset = 0; offset < 24; offset += rouletteNames.length) {
		assert.deepEqual([...reel.slice(offset, offset + rouletteNames.length)].sort(), rouletteNames.map(name => name.split(' ')[0]).sort(), 'Every shuffled pass includes all pupils, even with identical first names');
	}
	assert.equal(reel.at(-1), 'Anna', 'The roulette lands on the actual selected pupil');
	assert.ok(reel.every(name => ['Anna','Max','Nora'].includes(name)), 'Roulette displays first names only, preserving duplicate first names');
	assert.equal(await evaluate(`getComputedStyle(document.getElementById('selection-effect')).backgroundColor === getComputedStyle(document.getElementById('container')).backgroundColor`), true, 'Roulette uses the existing app surface');
	await shot('session-roulette');
	await waitFor(`!document.querySelector('.repetition-stage').classList.contains('selection-running')`);
	assert.equal(await evaluate(`getComputedStyle(document.getElementById('name')).visibility`), 'visible');
	assert.equal(await evaluate(`document.getElementById('res').inert`), false);
	assert.equal(await evaluate(`document.getElementById('name').textContent`), 'Anna');
	// An interrupted selection cannot reveal a stale winner or lock grading.
	win.webContents.send('name', ['Max Beispiel', 0], rouletteNames, students[1]);
	await waitFor(`document.querySelector('.repetition-stage').classList.contains('selection-running')`);
	win.webContents.send('ui-settings', { theme: 'midnight', visualEffectsEnabled: false });
	await pause(50);
	assert.equal(await evaluate(`document.querySelector('.repetition-stage').classList.contains('selection-running') || document.getElementById('res').inert`), false);
	await pause(2100);
	assert.equal(await evaluate(`document.getElementById('name').textContent`), 'Max');
	await click('end-session-btn');
	await click('confirm-session-end-btn');
	await waitFor(`!document.getElementById('session-empty-view').classList.contains('update-hidden')`);
	assert.equal(await evaluate(`!!document.getElementById('class-list').closest('#session-empty-view')`), true, 'The single class list returns home after the session');
	await pause(400);
	await evaluate(`document.getElementById('update-panel').classList.add('update-hidden')`);

	// Skeletons persist while the actual read is pending and disappear on completion.
	await click('assessment-nav-btn');
	await evaluate(`(() => { const select = document.getElementById('assessment-class-select'); select.value = 'L23a'; select.dispatchEvent(new Event('change')); })()`);
	await waitFor(`!!document.querySelector('#assessment-body .ui-skeleton-state')`);
	assert.equal(await evaluate(`document.getElementById('assessment-class-select').disabled`), true);
	await shot('assessment-loading');
	overviewResolve();
	await waitFor(`!!document.getElementById('assessment-create-round-btn')`);
	// Comments open only by choice, independently of ratings, and survive saving.
	showCommentRound=true;
	await evaluate(`document.getElementById('assessment-class-select').dispatchEvent(new Event('change'))`);
	await waitFor(`document.getElementById('assessment-body').getAttribute('aria-busy') === 'true'`);
	overviewResolve();
	await waitFor(`!!document.querySelector('[data-round-id]')`);
	await evaluate(`document.querySelector('[data-round-id]').click()`);
	await waitFor(`!!document.getElementById('assessment-teacher-btn')`);
	await click('assessment-teacher-btn');
	await waitFor(`!!document.querySelector('.teacher-assessment')`);
	assert.equal(await evaluate(`document.querySelectorAll('.teacher-criterion-comment[open]').length`),0);
	assert.equal(await evaluate(`document.querySelector('[data-teacher-comment]').checkVisibility()`),false);
	await evaluate(`(() => {const input=document.querySelector('[data-teacher-criterion]');input.value='4';input.dispatchEvent(new Event('input'));})()`);
	assert.equal(await evaluate(`document.querySelectorAll('.teacher-criterion-comment[open]').length`),0,'Rating does not open a comment');
	await evaluate(`document.querySelector('.teacher-criterion-comment summary').click();document.querySelector('[data-teacher-comment]').value='Individueller Kriteriumskommentar'`);
	assert.equal(await evaluate(`document.querySelectorAll('.teacher-criterion-comment[open]').length`),1);
	await evaluate(`document.querySelector('.teacher-criterion-comment summary').click()`);
	assert.equal(await evaluate(`document.querySelectorAll('.teacher-criterion-comment[open]').length`),0);
	await evaluate(`document.querySelectorAll('[data-teacher-criterion]').forEach(select=>{select.value='3';select.dispatchEvent(new Event('change'));})`);
	await click('teacher-save-btn');
	await waitFor(`document.querySelector('[data-teacher-comment]').value==='Individueller Kriteriumskommentar' && document.querySelectorAll('.teacher-criterion-comment[open]').length===0 && document.querySelector('.teacher-criterion-comment summary').textContent.includes('bearbeiten')`);
	const reloadedStore=new AssessmentStore(assessmentFile);
	assert.equal(Object.values(reloadedStore.getRound(workbook,'L23a',students,commentRound.id).round.assessments)[0].teacherCriterionComments.participation_frequency,'Individueller Kriteriumskommentar');
	await shot('assessment-criterion-comments');
	await click('assessment-back-btn');
	await waitFor(`!!document.getElementById('assessment-teacher-btn')`);
	await click('assessment-teacher-btn');
	await waitFor(`!!document.querySelector('.teacher-assessment')`);
	assert.equal(await evaluate(`document.querySelector('[data-teacher-comment]').value`),'Individueller Kriteriumskommentar');
	assert.equal(await evaluate(`document.querySelectorAll('.teacher-criterion-comment[open]').length`),0,'Existing comments also remain collapsed');
	// Resume the existing overview/loading checks.
	await click('assessment-back-btn');
	await waitFor(`!!document.getElementById('assessment-teacher-btn')`);
	await click('assessment-back-btn');
	await waitFor(`document.getElementById('assessment-body').getAttribute('aria-busy') === 'true'`);
	showCommentRound=false;
	overviewResolve();
	await waitFor(`!!document.getElementById('assessment-create-round-btn')`);
	assert.equal(await evaluate(`document.getElementById('assessment-body').getAttribute('aria-busy')`), 'false');
	assert.equal(await evaluate(`document.getElementById('assessment-body').inert`), false);
	await shot('assessment-overview');

	// A failed read presents a working retry instead of leaving the screen stuck.
	failOverview = true;
	await evaluate(`document.getElementById('assessment-class-select').dispatchEvent(new Event('change'))`);
	await waitFor(`!!document.querySelector('.ui-skeleton-state')`);
	overviewResolve();
	await waitFor(`!!document.getElementById('assessment-retry-btn')`);
	failOverview = false;
	await click('assessment-retry-btn');
	await waitFor(`!!document.querySelector('.ui-skeleton-state')`);
	overviewResolve();
	await waitFor(`!!document.getElementById('assessment-create-round-btn')`);
	await click('stats-view-btn');
	await waitFor(`!!document.querySelector('#stats-list .ui-skeleton-state')`);
	assert.equal(await visible('assessment-view'), false, 'Navigation shows one main view');
	await shot('stats-loading');
	win.webContents.send('class-statistics-data', []);
	await waitFor(`document.getElementById('stats-list').getAttribute('aria-busy') === 'false'`);
	assert.equal(await visible('stats-empty'), true);
	await click('stats-back-btn');
	await pause(400);

	// AI: focus containment, real activity, editable drafts and duplicate-send protection.
	await evaluate(`document.getElementById('ai-open-btn').focus()`);
	await click('ai-open-btn');
	await waitFor(`document.activeElement.id === 'ai-input'`);
	await click('ai-close-btn');
	await pause(30);
	await click('ai-open-btn');
	await pause(300);
	assert.equal(await evaluate(`document.getElementById('ai-modal').classList.contains('modal-open') && getComputedStyle(document.getElementById('ai-modal')).opacity === '1'`), true, 'Reopening cancels the old modal exit');
	assert.equal(await evaluate(`document.getElementById('ai-send-btn').disabled`), true);
	await shot('ai-empty');
 assert.equal(await evaluate(`!!document.querySelector('#ai-empty canvas.thinking-orb')`), true);
	await draft('Erste Zeile');
	await key('Enter', ['shift']);
	assert.match(await evaluate(`document.getElementById('ai-input').value`), /\n/);
	await evaluate(`document.getElementById('ai-close-btn').focus()`);
	await key('Tab', ['shift']);
	assert.equal(await evaluate(`!!document.activeElement.closest('#ai-modal')`), true);
	await draft('Fasse die aktuelle Session zusammen.');
	await evaluate(`document.getElementById('ai-input').focus()`);
	await evaluate(`document.getElementById('ai-form').addEventListener('submit',()=>{const m=document.querySelector('.ai-user:last-child'); const g=document.querySelector('.ai-send-flight'); const a=g?.getAnimations()[0];window.flightAudit={hidden:getComputedStyle(m).visibility,scroll:Math.abs(document.getElementById('ai-messages').scrollHeight-document.getElementById('ai-messages').clientHeight-document.getElementById('ai-messages').scrollTop),easing:a?.effect.getTiming().easing,loadingHidden:document.getElementById('ai-activity').classList.contains('update-hidden')};},{once:true})`);
 await key('Enter');
 await waitFor(`!document.getElementById('ai-activity').classList.contains('update-hidden')`);
 assert.equal(await evaluate(`flightAudit.hidden`),'hidden');
 assert.equal(await evaluate(`flightAudit.loadingHidden`),true,'Typing dots wait for the send flight to land');
 assert.equal(await evaluate(`!!document.querySelector('.ai-send-flight')`),false,'The message lands before typing dots appear');
 assert.ok(await evaluate(`flightAudit.scroll < 2`));
 assert.equal(await evaluate(`flightAudit.easing`),'cubic-bezier(0.7, 0, 0.84, 0)');
	assert.equal(await evaluate(`document.querySelectorAll('#ai-activity .typing-dots i').length`), 3);
	assert.equal(await evaluate(`document.getElementById('ai-input').disabled`), false);
	await draft('Meine nächste Frage');
	await evaluate(`document.getElementById('ai-form').dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }))`);
	assert.equal(asks, 1);
	win.webContents.send('ai-activity', { tool: 'getActiveTeachingSession', label: 'Sessiondaten werden gelesen …' });
	await waitFor(`document.getElementById('ai-activity-label').textContent.includes('Sessiondaten')`);
	assert.equal(await evaluate(`document.querySelector('#ai-activity .thinking-orb')`), null);
 await waitFor(`document.getElementById('ai-activity').dataset.activity === 'working'`);
 win.webContents.send('ai-activity', {tool:'searchStudents',label:'Schüler werden gesucht …'});
 await waitFor(`document.getElementById('ai-activity').dataset.activity === 'searching'`);
 win.webContents.send('ai-activity', {tool:'addParticipationPoints',label:'Bewertung wird vorbereitet …'});
 await waitFor(`document.getElementById('ai-activity').dataset.activity === 'solving'`);
	await shot('ai-thinking');
	askResolve({ type: 'answer', text: 'Die Klasse hat heute aktiv mitgearbeitet.\nDie lokalen Einträge wurden ausgewertet.' });
	await waitFor(`document.getElementById('ai-activity').classList.contains('update-hidden')`);
	assert.equal(await evaluate(`document.getElementById('ai-activity').dataset.busy`), 'false');
	assert.equal(await evaluate(`document.getElementById('ai-input').value`), 'Meine nächste Frage');
 await waitFor(`!!document.querySelector('.ai-assistant[data-typing]')`, 'Progressive answer rendering');
 assert.equal(await evaluate(`getComputedStyle(document.querySelector('.ai-assistant')).backgroundColor`), 'rgba(0, 0, 0, 0)');
 await waitFor(`!document.querySelector('.ai-assistant[data-typing]')`, 'Answer rendering completes');
 await shot('ai-answer');

	// Confirmations cannot be bypassed by another prompt, and failed confirmation can be retried.
	await key('Enter');
	await waitFor(`document.getElementById('ai-send-btn').disabled`);
	askResolve({ type: 'confirmation', token: 'test', title: 'Mitarbeit speichern?', lines: ['Anna Müller · 5 Punkte'] });
	await waitFor(`!!document.querySelector('.ai-confirm-card')`);
	await evaluate(`document.getElementById('ai-form').dispatchEvent(new Event('submit', { cancelable: true }))`);
	assert.equal(asks, 2);
	await evaluate(`document.querySelector('.ai-confirm-card .btn-1').click()`);
	await waitFor(`!!document.querySelector('.ai-error')`);
	assert.equal(await evaluate(`document.querySelector('.ai-confirm-card .btn-1').disabled`), false);
	await evaluate(`document.querySelector('.ai-confirm-card .btn-1').click()`);
	await waitFor(`!document.querySelector('.ai-confirm-card')`);
	expiredConfirm = true;
	await draft('Eine neue Aktion');
	await key('Enter');
	await waitFor(`!document.getElementById('ai-activity').classList.contains('update-hidden')`);
	askResolve({ type: 'confirmation', token: 'expired', title: 'Aktion bestätigen?', lines: ['Abgelaufene Bestätigung'] });
	await waitFor(`!!document.querySelector('.ai-confirm-card')`);
	await evaluate(`document.querySelector('.ai-confirm-card .btn-1').click()`);
	await waitFor(`!document.querySelector('.ai-confirm-card') && document.getElementById('ai-activity').classList.contains('update-hidden')`);
	await draft('Nach abgelaufener Bestätigung weiterfragen');
	assert.equal(await evaluate(`document.getElementById('ai-send-btn').disabled`), false);
	await draft('');
	await key('Escape');
	await waitFor(`getComputedStyle(document.getElementById('ai-modal')).display === 'none'`);
	assert.equal(await visible('ai-modal'), false);
	assert.equal(await evaluate(`document.activeElement.id`), 'ai-open-btn');
	aiState = 'unavailable';
	await click('ai-open-btn');
	await waitFor(`!document.getElementById('ai-configure-btn').classList.contains('update-hidden')`);
	assert.equal(await evaluate(`document.getElementById('ai-input').disabled`), true);
	await shot('ai-unavailable');
	await click('ai-configure-btn');
	await waitFor(`getComputedStyle(document.getElementById('ai-modal')).display === 'none' && !document.getElementById('settings-view').classList.contains('update-hidden')`);
	assert.equal(await evaluate(`!!document.activeElement.closest('#settings-view')`), true, 'Opening settings waits for the chat exit and retains focus inside the new dialog');
	await shot('settings');
	await click('close-settings-modal');
	await waitFor(`getComputedStyle(document.getElementById('settings-view')).display === 'none'`);
	aiState = 'ready';
	await waitFor(`getComputedStyle(document.getElementById('ai-modal')).display === 'none'`);

	// Reduced motion, both themes, and compact windows use the real computed styles.
	await win.webContents.debugger.attach('1.3');
	await click('ai-open-btn');
	await waitFor(`!document.getElementById('ai-input').disabled`);
	assert.equal(await evaluate(`(() => { document.getElementById('ai-close-btn').click(); return require('./vendor/ui-motion.js').gsap.getTweensOf(document.getElementById('ai-modal')).length > 0; })()`), true, 'The exit starts an actual GSAP transition before reduced motion interrupts it');
	await win.webContents.debugger.sendCommand('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-reduced-motion', value: 'reduce' }] });
	await waitFor(`getComputedStyle(document.getElementById('ai-modal')).display === 'none'`, 'Reduced-motion change finishes a pending exit');
	await click('ai-open-btn');
	await waitFor(`!document.getElementById('ai-input').disabled`);
	assert.equal(await evaluate(`require('./vendor/ui-motion.js').gsap.getTweensOf([document.getElementById('ai-modal'), document.querySelector('.ai-content')]).length`), 0, 'Reduced motion opens dialogs in their final state');
	await draft('Bewegungsreduktion prüfen');
	await evaluate(`document.getElementById('ai-form').dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }))`);
	await waitFor(`!document.getElementById('ai-activity').classList.contains('update-hidden')`);
	assert.equal(await evaluate(`getComputedStyle(document.querySelector('.typing-dots i')).animationName`), 'none');
	askResolve({ type: 'answer', text: 'Statische Rückmeldung geprüft.' });
	await waitFor(`document.getElementById('ai-activity').classList.contains('update-hidden')`);
	await click('ai-close-btn');
	assert.equal(await evaluate(`(() => { const state = window.createRepetiererLoadingState('Test'); document.body.appendChild(state); const value = getComputedStyle(state.querySelector('.skeleton-block'), '::after').animationName; state.remove(); return value; })()`), 'none');
	for (const theme of ['midnight', 'classic']) {
		await evaluate(`document.documentElement.dataset.theme = ${JSON.stringify(theme)}`);
		const contrast = await evaluate(`(() => {
			const control = getComputedStyle(document.getElementById('session-open-file-btn'));
			const canvas = document.createElement('canvas'); canvas.width = canvas.height = 1;
			const context = canvas.getContext('2d');
			const luminance = color => {
				context.clearRect(0, 0, 1, 1); context.fillStyle = color; context.fillRect(0, 0, 1, 1);
				const values = [...context.getImageData(0, 0, 1, 1).data].slice(0, 3).map(value => { const s = value / 255; return s <= .04045 ? s / 12.92 : ((s + .055) / 1.055) ** 2.4; });
				return values[0] * .2126 + values[1] * .7152 + values[2] * .0722;
			};
			const foreground = luminance(control.color);
            const r = window.repetiererKristall.renderer, gl = r.gl;
            r.draw();
            const box = document.getElementById('session-open-file-btn').getBoundingClientRect();
            const pixels = new Uint8Array(4);
            gl.readPixels(Math.floor((box.left + box.width / 2) * r.pixelRatio[0]), Math.floor((r.height - box.top - box.height / 2) * r.pixelRatio[1]), 1, 1, gl.RGBA, gl.UNSIGNED_BYTE, pixels);
            const background = luminance('rgb(' + [...pixels].slice(0,3).join(',') + ')');
			return (Math.max(foreground, background) + .05) / (Math.min(foreground, background) + .05);
		})()`);
		assert.ok(contrast >= 4.5, theme + ' primary controls meet normal text contrast');
		await shot(`home-${theme}`);
	}
	win.setSize(680, 560);
	await pause(100);
	await verifyModalCenters();
	await shot('home-compact');
	const homeBounds = await evaluate(`(() => { const rect = document.getElementById('session-empty-view').getBoundingClientRect(); return { left: rect.left, right: rect.right, viewport: innerWidth }; })()`);
	// Windows DPI scaling can round the CSS viewport down by a fractional pixel.
	assert.ok(homeBounds.left >= -1 && homeBounds.right <= homeBounds.viewport + 1, 'Compact entry remains within the viewport: ' + JSON.stringify(homeBounds));
	await click('ai-open-btn');
	await waitFor(`!document.getElementById('ai-input').disabled`);
	await shot('ai-compact');
	assert.equal(await evaluate(`(() => { const rect = document.querySelector('.ai-content').getBoundingClientRect(); return rect.left >= 0 && rect.right <= innerWidth && rect.bottom <= innerHeight; })()`), true);
	assert.equal(await evaluate(`(() => { const rect = document.getElementById('ai-input').getBoundingClientRect(); return rect.width > 100 && rect.bottom <= innerHeight; })()`), true);
	await require('./start-flow-ui-check.cjs')({win,evaluate,waitFor,shot,pause,setStartFailure:value=>{failTeachingStart=value;}});
	await require('./joker-date-ui-check.cjs')({win,evaluate,waitFor,shot,pause});
	await require('./settings-page-ui-check.cjs')({win,evaluate,waitFor,shot,pause});
	await require('./liquid-refinement-ui-check.cjs')({win,evaluate,waitFor,shot,pause,bundleFixture});
	assert.deepEqual(errors, [], 'Keine Rendererfehler');
	console.log(`UI smoke passed: hidden session sidebar, shuffled roulette without winner leakage, larger grading controls, real GSAP drawer/modal transitions and interrupted exits, ${modalCount} centered dialogs at two window sizes, repetition/name sizing, three-dot typing states/reduced motion, entry/session, read loading/error/retry, AI draft/multiline/confirm/retry/expiry, keyboard/focus, themes, reduced motion, compact layout.`);
	const splash = new BrowserWindow({width: 560, height: 260, frame: false, transparent: true, show: false, webPreferences: {nodeIntegration: true, contextIsolation: false, offscreen: true}});
	await splash.loadFile(path.join(__dirname, '..', 'public', 'splash.html'));
	assert.equal(await splash.webContents.executeJavaScript(`getComputedStyle(document.getElementById('container')).borderTopLeftRadius`), '18px', 'Only the intro gets rounder outside corners');
	assert.equal(await splash.webContents.executeJavaScript(`getComputedStyle(document.getElementById('container')).clipPath`), 'inset(0px round 18px)');
	await pause(500);
	fs.writeFileSync(path.join(output, 'intro-rounded.png'), (await splash.webContents.capturePage()).toPNG());
	splash.destroy();
	console.log(`Screenshots: ${output}`);
}

const watchdog = setTimeout(() => { console.error('UI smoke timed out.'); app.exit(1); }, 120000);
main().then(() => { clearTimeout(watchdog); app.exit(0); }).catch(error => { console.error(error); clearTimeout(watchdog); app.exit(1); });
