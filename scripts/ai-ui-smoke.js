// Start Electron with --remote-debugging-port=9348, then run this local UI smoke test.
const port = Number(process.env.REPETIERER_AI_CDP_PORT || 9348);

async function main() {
	const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json();
	const target = targets.find(item => item.url.endsWith('/public/index.html'));
	if (!target) throw new Error('Repetierer-Hauptfenster nicht gefunden.');
	const socket = new WebSocket(target.webSocketDebuggerUrl);
	await new Promise((resolve, reject) => { socket.addEventListener('open', resolve, { once: true }); socket.addEventListener('error', reject, { once: true }); });
	let id = 0;
	const send = (method, params) => new Promise((resolve, reject) => {
		const requestId = ++id;
		const listener = event => {
			const message = JSON.parse(event.data);
			if (message.id !== requestId) return;
			socket.removeEventListener('message', listener);
			if (message.error || message.result?.exceptionDetails) reject(new Error(JSON.stringify(message.error || message.result.exceptionDetails)));
			else resolve(message.result);
		};
		socket.addEventListener('message', listener);
		socket.send(JSON.stringify({ id: requestId, method, params }));
	});
	const evaluate = async expression => (await send('Runtime.evaluate', { expression, returnByValue: true })).result.value;
	const before = await evaluate(`(() => { const button = document.getElementById('ai-open-btn'); const rect = button.getBoundingClientRect(); return { appRegion: getComputedStyle(button).webkitAppRegion, x: rect.x + rect.width / 2, y: rect.y + rect.height / 2, modalDisplay: getComputedStyle(document.getElementById('ai-modal')).display }; })()`);
	if (before.appRegion !== 'no-drag') throw new Error(`KI-Button ist keine klickbare Header-Region: ${before.appRegion}`);
	await send('Input.dispatchMouseEvent', { type: 'mousePressed', button: 'left', clickCount: 1, x: before.x, y: before.y });
	await send('Input.dispatchMouseEvent', { type: 'mouseReleased', button: 'left', clickCount: 1, x: before.x, y: before.y });
	await new Promise(resolve => setTimeout(resolve, 200));
	const after = await evaluate(`(() => { const modal = document.getElementById('ai-modal'); return { display: getComputedStyle(modal).display, open: modal.classList.contains('modal-open'), status: document.getElementById('ai-chat-status').textContent, inputDisabled: document.getElementById('ai-input').disabled }; })()`);
	console.log(JSON.stringify({ before, after }));
	socket.close();
	if (!after.open || after.display === 'none' || after.inputDisabled) process.exitCode = 1;
}

main().catch(error => { console.error(error.message); process.exitCode = 1; });
