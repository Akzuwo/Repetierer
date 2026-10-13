// Capture the current UI through Electron's Chrome DevTools Protocol.
// Start Electron with --remote-debugging-port=9348 before running this script.
const fs = require('fs');
const path = require('path');
const assert = require('assert/strict');

const port = Number(process.env.REPETIERER_AI_CDP_PORT || 9348);
const outputDir = path.join(__dirname, '..', 'docs', 'ui-audit', 'current');

async function main() {
	let target;
	for (let attempt = 0; attempt < 20 && !target; attempt++) {
		try {
			const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json();
			target = targets.find(item => item.url.endsWith('/public/index.html'));
		} catch { /* Electron may still be starting. */ }
		if (!target) await new Promise(resolve => setTimeout(resolve, 250));
	}
	if (!target) throw new Error('Repetierer-Hauptfenster nicht gefunden.');
	const socket = new WebSocket(target.webSocketDebuggerUrl);
	await new Promise((resolve, reject) => {
		socket.addEventListener('open', resolve, { once: true });
		socket.addEventListener('error', reject, { once: true });
	});
	let nextId = 0;
	const send = (method, params = {}) => new Promise((resolve, reject) => {
		const id = ++nextId;
		const timer = setTimeout(() => {
			socket.removeEventListener('message', onMessage);
			reject(new Error(`CDP-Zeitüberschreitung: ${method}`));
		}, 10000);
		const onMessage = event => {
			const message = JSON.parse(event.data);
			if (message.id !== id) return;
			clearTimeout(timer);
			socket.removeEventListener('message', onMessage);
			if (message.error || message.result?.exceptionDetails) reject(new Error(JSON.stringify(message.error || message.result.exceptionDetails)));
			else resolve(message.result);
		};
		socket.addEventListener('message', onMessage);
		socket.send(JSON.stringify({ id, method, params }));
	});
	const evaluate = async expression => (await send('Runtime.evaluate', { expression, returnByValue: true })).result.value;
	const disableMotion = () => evaluate(`(() => { const style = document.createElement('style'); style.textContent = '*,*::before,*::after{animation:none!important;transition:none!important}.modal.modal-open{opacity:1!important}.modal.modal-open .modal-content{opacity:1!important;transform:none!important}'; document.head.appendChild(style); return true; })()`);
	const shot = async name => {
		await evaluate(`(() => {
			const replace = (selector, value) => { const item = document.querySelector(selector); if (item) item.textContent = value; };
			replace('#brevo-status', 'Eingerichtet · Absender konfiguriert');
			replace('#file-path', 'Beispiel.xlsx');
			replace('#settings-location', 'Lokale Einstellungen');
			replace('#backup-location', 'Lokaler Speicherort');
			replace('#log-location', 'Lokaler Speicherort');
			replace('#log-list', '');
			replace('#problem-report-summary', '');
			replace('#problem-report-details', '');
			replace('#assessment-confirm-text', 'Vorschau der Versandaktion.');
			replace('#assessment-delete-text', 'Diese Beurteilungsrunde wird lokal gelöscht.');
			replace('#assessment-delete-sent-warning', '');
			replace('#assessment-body', 'Bitte wähle eine Klasse.');
			replace('#session-class-name', 'Klasse 1');
			if (!${JSON.stringify(name.endsWith('-layout'))}) { replace('#name', 'Schüler'); replace('#participation-student-list', ''); }
			document.querySelectorAll('#class-list button').forEach((item, index) => { item.textContent = 'Klasse ' + (index + 1); });
			document.querySelectorAll('#assessment-class-select option').forEach((item, index) => { if (index) item.textContent = 'Klasse ' + index; });
			document.querySelectorAll('#session-attendance-list, #person-list, #absence-list, #backup-list, #probability-list, #excel-editor-list, #class-import-preview, #participation-score-list').forEach(item => { item.textContent = ''; });
			return true;
		})()`);
		await new Promise(resolve => setTimeout(resolve, 120));
		const { data } = await send('Page.captureScreenshot', { format: 'png', captureBeyondViewport: false });
		fs.writeFileSync(path.join(outputDir, `${name}.png`), Buffer.from(data, 'base64'));
	};
	fs.mkdirSync(outputDir, { recursive: true });
	await send('Page.enable');
	await send('Page.reload', { ignoreCache: true });
	await new Promise(resolve => setTimeout(resolve, 1200));
	await disableMotion();
	await evaluate(`(() => { document.querySelectorAll('.modal').forEach(item => { item.style.display = 'none'; item.classList.remove('modal-open'); }); document.getElementById('drawer').classList.remove('drawer-open'); document.getElementById('drawer-scrim').classList.remove('scrim-visible'); return true; })()`);
	const homeState = await evaluate(`(() => ({ title: document.getElementById('title').textContent, authorInHeader: !!document.querySelector('#title-bar #title-author'), versionInHeader: !!document.querySelector('#title-bar #version'), emptyVisible: getComputedStyle(document.getElementById('session-empty-view')).display !== 'none', repetitionVisible: getComputedStyle(document.getElementById('repetition')).display !== 'none' }))()`);
	assert.equal(homeState.title, 'Repetierer');
	assert.equal(homeState.authorInHeader || homeState.versionInHeader, false);
	assert.equal(homeState.emptyVisible, true);
	assert.equal(homeState.repetitionVisible, false);
	await shot('home');
	await evaluate(`(() => { document.getElementById('drawer-toggle').click(); return true; })()`);
	await shot('drawer-main');
	for (const [name, id] of [['drawer-data', 'drawer-data-view-btn'], ['drawer-protocol', 'drawer-protocol-view-btn']]) {
		await evaluate(`document.getElementById(${JSON.stringify(id)}).click()`);
		await shot(name);
		await evaluate(`document.getElementById('drawer-back-btn').click()`);
	}
	await evaluate(`document.getElementById('drawer-toggle').click()`);
	await evaluate(`document.getElementById('settings-btn').click()`);
	await shot('settings');
	await evaluate(`(() => { document.querySelector('#settings-modal .settings-scroll-area').scrollTop = 10000; return true; })()`);
	await shot('settings-integrations');
	await evaluate(`(() => { document.getElementById('settings-modal').style.display = 'none'; document.getElementById('settings-modal').classList.remove('modal-open'); document.getElementById('ai-open-btn').click(); return true; })()`);
	await shot('ai-empty');
	await evaluate(`document.querySelector('[data-ai-prompt]').click()`);
	assert.equal(await evaluate(`document.getElementById('ai-input').value`), 'Fasse die aktuelle Session zusammen.');
	await evaluate(`(() => { document.getElementById('ai-modal').style.display = 'none'; document.getElementById('ai-modal').classList.remove('modal-open'); document.getElementById('assessment-nav-btn').click(); return true; })()`);
	assert.equal(await evaluate(`document.getElementById('title').textContent`), 'Repetierer');
	await shot('assessment');
	await evaluate(`(() => {
		document.getElementById('assessment-view').classList.add('update-hidden');
		document.getElementById('content').classList.add('session-active');
		document.getElementById('session-header').classList.remove('update-hidden');
		document.getElementById('session-empty-view').classList.add('update-hidden');
		document.getElementById('repetition').classList.remove('update-hidden');
		document.getElementById('session-class-name').textContent = 'Klasse 1';
		document.getElementById('session-attendance-count').textContent = '18 anwesend';
		const name = document.getElementById('name'); name.textContent = 'Anna Müller'; name.classList.remove('disabled'); name.style.fontSize = '5.5rem';
		document.querySelectorAll('#res .disabled').forEach(item => item.classList.remove('disabled'));
		document.querySelector('#grade input').value = '5.5';
		return true;
	})()`);
	await shot('repetition-layout');
	await evaluate(`(() => {
		document.getElementById('repetition').classList.add('update-hidden');
		document.getElementById('participation-view').classList.remove('update-hidden');
		document.getElementById('participation-student-list').innerHTML = '<button class="participation-student"><strong>Anna Müller</strong><span>2 Einträge</span></button><button class="participation-student"><strong>Max Beispiel</strong><span>1 Eintrag</span></button>';
		return true;
	})()`);
	await shot('participation-layout');
	await evaluate(`(() => {
		document.getElementById('content').classList.remove('session-active');
		document.getElementById('session-header').classList.add('update-hidden');
		document.getElementById('participation-view').classList.add('update-hidden');
		document.getElementById('stats-view').classList.remove('update-hidden');
		document.getElementById('stats-empty').classList.remove('update-hidden');
		return true;
	})()`);
	await shot('statistics-layout');
	await send('Page.reload', { ignoreCache: true });
	await new Promise(resolve => setTimeout(resolve, 1200));
	await disableMotion();
	const modals = [
		'person-modal', 'absence-modal', 'backup-modal', 'log-modal', 'problem-report-modal',
		'reload-conflict-modal', 'class-import-modal', 'teacher-help-modal', 'joker-migration-modal',
		'joker-migration-help-modal', 'probabilities-modal', 'excel-editor-modal', 'update-news-modal',
		'settings-help-modal', 'session-start-modal', 'session-end-modal', 'participation-modal',
		'assessment-confirm-modal', 'assessment-delete-modal'
	];
	for (const id of modals) {
		await evaluate(`(() => { document.querySelectorAll('.modal').forEach(item => { item.style.display = 'none'; item.classList.remove('modal-open'); }); const modal = document.getElementById(${JSON.stringify(id)}); modal.style.display = 'block'; modal.classList.add('modal-open'); return true; })()`);
		const bounds = await evaluate(`(() => { const modal = document.getElementById('${id}'); const rect = modal.querySelector('.modal-content').getBoundingClientRect(); return { left: rect.left, right: rect.right, width: innerWidth, opacity: getComputedStyle(modal).opacity }; })()`);
		assert.ok(bounds.left >= 0 && bounds.right <= bounds.width, `${id} liegt ausserhalb des Fensters`);
		assert.equal(bounds.opacity, '1', `${id} ist nicht sichtbar`);
		await shot(id);
	}
	await evaluate(`(() => { document.querySelectorAll('.modal').forEach(item => { item.style.display = 'none'; item.classList.remove('modal-open'); }); return true; })()`);
	socket.close();
	console.log(`Screenshots gespeichert: ${outputDir}`);
}

main().catch(error => { console.error(error.message); process.exitCode = 1; });
