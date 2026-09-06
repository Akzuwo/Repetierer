const fs = require('fs');

const port = Number(process.env.REPETIERER_SMOKE_PORT || 9333);
const mode = process.env.REPETIERER_SMOKE_MODE || 'create';
const statePath = process.env.REPETIERER_SMOKE_STATE;

async function main() {
	const target = await waitForTarget();
	const socket = new WebSocket(target.webSocketDebuggerUrl);
	await new Promise((resolve, reject) => {
		socket.addEventListener('open', resolve, { once: true });
		socket.addEventListener('error', reject, { once: true });
	});
	let nextId = 1;
	const evaluate = expression => new Promise((resolve, reject) => {
		const id = nextId++;
		const onMessage = event => {
			const message = JSON.parse(event.data);
			if (message.id !== id) return;
			socket.removeEventListener('message', onMessage);
			if (message.error || message.result.exceptionDetails) reject(new Error(JSON.stringify(message.error || message.result.exceptionDetails)));
			else resolve(message.result.result.value);
		};
		socket.addEventListener('message', onMessage);
		socket.send(JSON.stringify({ id, method: 'Runtime.evaluate', params: { expression, awaitPromise: true, returnByValue: true } }));
	});

	const classes = await evaluate(`new Promise((resolve, reject) => {
		const started = Date.now();
		const check = () => {
			const values = Array.from(document.querySelectorAll('#assessment-class-select option')).map(option => option.value).filter(Boolean);
			if (values.length) return resolve(values);
			if (Date.now() - started > 15000) return reject(new Error('Keine Klassen geladen'));
			setTimeout(check, 200);
		};
		check();
	})`);
	const className = classes[0];
	if (mode === 'create') {
		const result = await evaluate(`(async () => {
			const { ipcRenderer } = require('electron');
			const api = await ipcRenderer.invoke('assessment:api-test');
			if (!api.ok) return { ok: false, stage: 'api', error: api.error };
			const created = await ipcRenderer.invoke('assessment:create-round', { className: ${JSON.stringify(className)}, title: 'Release Smoke ' + Date.now() });
			if (!created.ok) return { ok: false, stage: 'create', error: created.error };
			const round = created.data.rounds[0];
			const view = await ipcRenderer.invoke('assessment:get-round', { className: ${JSON.stringify(className)}, roundId: round.id });
			if (!view.ok) return { ok: false, stage: 'round', error: view.error };
			const link = new URL(view.data.invitation.text.split('\\n').find(line => line.startsWith('https://')));
			const entryKey = Array.from(link.searchParams.keys()).find(key => key.startsWith('entry.'));
			const sync = await ipcRenderer.invoke('assessment:sync-responses', { className: ${JSON.stringify(className)}, roundId: round.id });
			return { ok: sync.ok, stage: sync.ok ? 'complete' : 'sync', error: sync.error, className: ${JSON.stringify(className)}, roundId: round.id, externalRoundId: round.externalRoundId, entryKey, entryValue: entryKey && link.searchParams.get(entryKey), studentCount: view.data.students.length };
		})()`);
		if (!result.ok) throw new Error(JSON.stringify(result));
		fs.writeFileSync(statePath, JSON.stringify(result), 'utf8');
		console.log(JSON.stringify(result));
	} else if (mode === 'verify') {
		const previous = JSON.parse(fs.readFileSync(statePath, 'utf8'));
		const result = await evaluate(`(async () => {
			const { ipcRenderer } = require('electron');
			const before = await ipcRenderer.invoke('assessment:get-overview', { className: ${JSON.stringify(previous.className)} });
			if (!before.ok) return { ok: false, stage: 'reload', error: before.error };
			const persisted = before.data.rounds.some(round => round.id === ${JSON.stringify(previous.roundId)});
			const studentCountBefore = before.data.students.length;
			const removed = await ipcRenderer.invoke('assessment:delete-round', { className: ${JSON.stringify(previous.className)}, roundId: ${JSON.stringify(previous.roundId)} });
			if (!removed.ok) return { ok: false, stage: 'delete', error: removed.error };
			return { ok: persisted && !removed.data.rounds.some(round => round.id === ${JSON.stringify(previous.roundId)}) && removed.data.students.length === studentCountBefore, persisted, deleted: !removed.data.rounds.some(round => round.id === ${JSON.stringify(previous.roundId)}), studentsPreserved: removed.data.students.length === studentCountBefore };
		})()`);
		if (!result.ok) throw new Error(JSON.stringify(result));
		console.log(JSON.stringify(result));
	} else {
		const result = await evaluate(`new Promise((resolve, reject) => {
			document.getElementById('assessment-nav-btn').click();
			const select = document.getElementById('assessment-class-select');
			select.value = ${JSON.stringify(className)};
			select.dispatchEvent(new Event('change'));
			const started = Date.now();
			const openRound = () => {
				const card = document.querySelector('[data-round-id]');
				if (!card) {
					if (Date.now() - started > 15000) return reject(new Error('Rundenübersicht nicht geladen'));
					return setTimeout(openRound, 200);
				}
				card.click();
				const waitForRound = () => {
					const invite = document.getElementById('assessment-invitation-btn');
					if (!invite) return setTimeout(waitForRound, 200);
					resolve({
						ok: true,
						heading: document.getElementById('assessment-heading').innerText,
						buttons: Array.from(document.querySelectorAll('.assessment-toolbar button')).map(button => button.innerText),
						hasRoster: !!document.querySelector('.assessment-roster'),
						hasInvitation: !!document.querySelector('.assessment-invitation')
					});
				};
				waitForRound();
			};
			openRound();
		})`);
		if (!result.ok) throw new Error(JSON.stringify(result));
		console.log(JSON.stringify(result));
	}
	socket.close();
}

async function waitForTarget() {
	const deadline = Date.now() + 20000;
	while (Date.now() < deadline) {
		try {
			const targets = await (await fetch(`http://127.0.0.1:${port}/json`)).json();
			const target = targets.find(item => item.type === 'page' && item.title === 'Repetierer');
			if (target) return target;
		} catch (error) {}
		await new Promise(resolve => setTimeout(resolve, 250));
	}
	throw new Error('Installierte Repetierer-Oberfläche wurde nicht erreichbar.');
}

main().catch(error => {
	console.error(error.message);
	process.exit(1);
});
