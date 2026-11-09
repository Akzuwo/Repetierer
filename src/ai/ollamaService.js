const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');
const { Readable } = require('stream');
const { pipeline } = require('stream/promises');

const MODEL = 'qwen3:4b-instruct';
const BASE = 'http://127.0.0.1:11434';
const INSTALLER = 'https://ollama.com/download/OllamaSetup.exe';

function createOllamaService({ fetchImpl = fetch, platform = process.platform, localAppData = process.env.LOCALAPPDATA || path.join(os.homedir(), 'AppData', 'Local') } = {}) {
	let setupRunning = false;
	async function request(endpoint, options = {}, timeout = 12000) {
		let response;
		try { response = await fetchImpl(`${BASE}${endpoint}`, { ...options, signal: AbortSignal.timeout(timeout) }); }
		catch (error) { throw new Error('Ollama ist nicht erreichbar.'); }
		if (!response.ok) {
			const detail = (await response.text()).slice(0, 400);
			if (response.status === 404 && /model/i.test(detail)) throw new Error(`Modell ${MODEL} ist nicht installiert.`);
			throw new Error(`Ollama-Fehler (${response.status}): ${detail}`);
		}
		return response;
	}
	async function isReachable() { try { await request('/api/tags'); return true; } catch { return false; } }
	async function hasModel() {
		const data = await (await request('/api/tags')).json();
		return (data.models || []).some(model => model.name === MODEL || model.model === MODEL);
	}
	async function status() {
		if (!await isReachable()) return platform === 'win32' && findExecutable() ? { state: 'error', message: 'Ollama ist installiert, aber nicht erreichbar.' } : { state: 'not-installed', message: 'Nicht installiert' };
		return await hasModel() ? { state: 'ready', message: 'Bereit' } : { state: 'model-missing', message: 'Modell nicht vorhanden' };
	}
	function findExecutable() {
		if (platform !== 'win32') return null;
		const candidates = [path.join(localAppData, 'Programs', 'Ollama', 'ollama.exe'), path.join(localAppData, 'Ollama', 'ollama.exe')];
		return candidates.find(candidate => fs.existsSync(candidate)) || null;
	}
	function freeBytes(directory) {
		try { const info = fs.statfsSync(directory, { bigint: true }); return info.bavail * info.bsize; }
		catch { return null; }
	}
	async function waitUntilReady(seconds = 90) {
		for (let i = 0; i < seconds; i += 2) {
			if (await isReachable()) return;
			await new Promise(resolve => setTimeout(resolve, 2000));
		}
		throw new Error('Ollama ist nach dem Start nicht erreichbar.');
	}
	async function startInstalled() {
		if (await isReachable()) return;
		const executable = findExecutable();
		if (!executable) throw new Error('Ollama ist nicht installiert.');
		const child = spawn(executable, ['serve'], { detached: true, stdio: 'ignore', windowsHide: true });
		child.unref();
		await waitUntilReady();
	}
	async function installWindows(onProgress) {
		if (platform !== 'win32') throw new Error('Die automatische Installation ist derzeit nur unter Windows verfügbar.');
		const free = freeBytes(os.homedir());
		if (free !== null && free < 7n * 1024n ** 3n) throw new Error('Zu wenig Speicherplatz: Für Ollama und das Modell werden mindestens 7 GB frei benötigt.');
		onProgress({ state: 'installing', message: 'Ollama wird heruntergeladen …' });
		let response;
		try { response = await fetchImpl(INSTALLER, { signal: AbortSignal.timeout(10 * 60 * 1000) }); }
		catch { throw new Error('Keine Internetverbindung oder Ollama-Download fehlgeschlagen.'); }
		if (!response.ok) throw new Error(`Ollama-Download fehlgeschlagen (${response.status}).`);
		if (!response.url.startsWith('https://')) throw new Error('Ollama-Download verwendet keine sichere Verbindung.');
		const installerPath = path.join(os.tmpdir(), 'Repetierer-OllamaSetup.exe');
		try { await pipeline(Readable.fromWeb(response.body), fs.createWriteStream(installerPath)); }
		catch { throw new Error('Ollama-Download fehlgeschlagen oder zu wenig Speicherplatz.'); }
		const signature = Buffer.alloc(2);
		const descriptor = fs.openSync(installerPath, 'r');
		try { fs.readSync(descriptor, signature, 0, 2, 0); }
		finally { fs.closeSync(descriptor); }
		if (signature.toString('ascii') !== 'MZ') throw new Error('Der Ollama-Installer ist ungültig.');
		onProgress({ state: 'installing', message: 'Ollama wird installiert …' });
		await new Promise((resolve, reject) => {
			const child = spawn(installerPath, ['/SILENT', '/NORESTART'], { windowsHide: true });
			child.once('error', reject);
			child.once('exit', code => code === 0 ? resolve() : reject(new Error(`Installation fehlgeschlagen (Code ${code}).`)));
		});
		await startInstalled();
	}
	async function pullModel(onProgress) {
		const free = freeBytes(os.homedir());
		if (free !== null && free < 3n * 1024n ** 3n) throw new Error('Zu wenig Speicherplatz für das Modell (mindestens 3 GB frei benötigt).');
		onProgress({ state: 'downloading', message: 'Modell wird heruntergeladen …', percent: 0 });
		let response;
		try { response = await request('/api/pull', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ model: MODEL, stream: true }) }, 60 * 60 * 1000); }
		catch (error) { throw new Error(`Modell-Download fehlgeschlagen: ${error.message}`); }
		let buffer = '';
		let lastProgress = '';
		try {
			for await (const chunk of response.body) {
				buffer += Buffer.from(chunk).toString('utf8');
				const lines = buffer.split('\n'); buffer = lines.pop();
				for (const line of lines) {
					if (!line.trim()) continue;
					const item = JSON.parse(line);
					if (item.error) throw new Error(item.error);
					const percent = item.total ? Math.min(100, Math.round(100 * item.completed / item.total)) : undefined;
					const progressKey = `${item.status || ''}:${percent === undefined ? '' : percent}`;
					if (progressKey !== lastProgress) onProgress({ state: 'downloading', message: item.status || 'Modell wird heruntergeladen …', percent });
					lastProgress = progressKey;
				}
			}
		} catch (error) { throw new Error(`Modell-Download fehlgeschlagen: ${error.message}`); }
		if (!await hasModel()) throw new Error('Modell-Download fehlgeschlagen: Modell nicht vorhanden.');
	}
	async function chat(messages, tools = []) {
		if (!await isReachable()) throw new Error('Ollama ist nicht erreichbar.');
		if (!await hasModel()) throw new Error(`Modell ${MODEL} ist nicht installiert.`);
		const response = await request('/api/chat', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ model: MODEL, messages, tools, stream: false, think: false }) }, 120000);
		return (await response.json()).message;
	}
	async function setup(onProgress = () => {}) {
		if (setupRunning) throw new Error('Die Einrichtung läuft bereits.');
		setupRunning = true;
		try {
			if (!await isReachable()) {
				if (findExecutable()) { onProgress({ state: 'installing', message: 'Ollama wird gestartet …' }); await startInstalled(); }
				else await installWindows(onProgress);
			}
			if (!await hasModel()) await pullModel(onProgress);
			onProgress({ state: 'downloading', message: 'Funktionstest läuft …', percent: 100 });
			const result = await chat([{ role: 'user', content: 'Antworte mit OK.' }]);
			if (!result || !result.content) throw new Error('Der Modell-Funktionstest ist fehlgeschlagen.');
			return { state: 'ready', message: 'Bereit' };
		} finally { setupRunning = false; }
	}
	return { status, setup, chat, hasModel, isReachable };
}

module.exports = { createOllamaService, MODEL };
