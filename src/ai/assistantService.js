const { randomUUID } = require('crypto');

const SYSTEM_PROMPT = `Du bist der lokale Assistent von Repetierer. Antworte standardmässig kurz und auf Deutsch.
Repetierer-Daten sind die einzige zuverlässige Quelle für Schüler, Klassen, Noten, Mitarbeit und Sessions. Erfinde niemals Schüler, Noten, IDs oder andere Repetierer-Daten. Verwende bei Bedarf die verfügbaren Tools und hole nur die Daten, die für die konkrete Frage nötig sind. Wenn die Anfrage einen Schüler beim Namen nennt, rufe immer zuerst searchStudents auf. Rate niemals eine studentId. Eine ID darf nur aus einer eindeutigen Suche oder dem aktuellen App-Kontext stammen.
Verwende für Änderungen ausschliesslich die bereitgestellten Tools. Behaupte niemals, eine Änderung vorgenommen zu haben, bevor das entsprechende Tool erfolgreich ausgeführt wurde. Bei mehrdeutigen Personen zuerst suchen und auflösen; frage nach einer Präzisierung, falls mehrere Treffer bleiben. Verwende die Klassen-ID und Schüler-ID aus Tool-Ergebnissen. Führe keine Aktionen ausserhalb der bereitgestellten Tools aus.
Die Excel-Repetitionsnoten und die Sitzungsdaten können unterschiedliche Historien enthalten. Wenn du nach Repetitionsnoten gefragt wirst, verwende getStudentGrades. Die aktuelle UI-Kontextinformation ist nur eine Navigationshilfe.`;

function createAiAssistantService({ ollama, makeRegistry, onActivity = () => {} }) {
	let messages = [{ role: 'system', content: SYSTEM_PROMPT }];
	let pending = null;
	let busy = false;
	function reset() { messages = [{ role: 'system', content: SYSTEM_PROMPT }]; pending = null; }
	function trim() { if (messages.length > 40) messages = [messages[0], ...messages.slice(-30)]; }
	async function run(registry, count = 0) {
		for (let step = count; step < 12; step++) {
			const message = await ollama.chat(messages, registry.definitions);
			if (!message) throw new Error('Ollama hat keine Antwort geliefert.');
			const calls = Array.isArray(message.tool_calls) ? message.tool_calls : [];
			messages.push({ role: 'assistant', content: message.content || '', ...(calls.length ? { tool_calls: calls } : {}) });
			if (!calls.length) { trim(); return { type: 'answer', text: String(message.content || 'Keine Antwort erhalten.') }; }
			for (let index = 0; index < calls.length; index++) {
				const call = calls[index];
				const name = call && call.function && call.function.name;
				onActivity({ tool: name, label: activityLabel(name) });
				let result;
				try {
					const args = typeof call.function.arguments === 'string' ? JSON.parse(call.function.arguments) : call.function.arguments;
					result = await registry.execute(name, args);
				} catch (error) { result = { error: error.message }; }
				if (result && result.confirmation) {
					pending = { token: randomUUID(), action: result.confirmation, registry, step: step + 1, remaining: calls.slice(index + 1) };
					return { type: 'confirmation', token: pending.token, title: result.confirmation.title, lines: result.confirmation.lines };
				}
				messages.push({ role: 'tool', tool_name: name, content: JSON.stringify(result) });
			}
		}
		throw new Error('Zu viele aufeinanderfolgende Tool-Aufrufe.');
	}
	async function ask(input) {
		if (busy || pending) throw new Error('Bitte zuerst die laufende Anfrage abschliessen.');
		if (typeof input !== 'string' || !input.trim() || input.length > 2000) throw new Error('Bitte eine kurze Frage eingeben.');
		busy = true;
		try {
			const registry = makeRegistry();
			const context = await registry.execute('getCurrentAppContext', {});
			messages.push({ role: 'user', content: `App-Kontext: ${JSON.stringify(context)}\nAnfrage: ${input.trim()}` });
			return await run(registry);
		}
		finally { busy = false; }
	}
	async function decide(token, approved) {
		if (busy || !pending || pending.token !== token) throw new Error('Bestätigung ist ungültig oder abgelaufen.');
		busy = true;
		const current = pending; pending = null;
		try {
			let result;
			if (!approved) result = { cancelled: true };
			else {
				try { result = await current.registry.commit(current.action); }
				catch (error) { result = { error: error.message }; }
			}
			messages.push({ role: 'tool', tool_name: current.action.name, content: JSON.stringify(result) });
			if (result.cancelled) return { type: 'answer', text: 'Abgebrochen. Es wurde nichts geändert.' };
			if (result.error) return { type: 'answer', text: `Änderung fehlgeschlagen: ${result.error}` };
			if (result.pendingExcel) return { type: 'answer', text: 'Die Note wurde lokal vorgemerkt. Schliessen Sie Excel und wählen Sie «Excel aktualisieren», um sie in die Datei zu schreiben.' };
			for (let index = 0; index < current.remaining.length; index++) {
				const call = current.remaining[index];
				const name = call && call.function && call.function.name;
				let next;
				try {
					const args = typeof call.function.arguments === 'string' ? JSON.parse(call.function.arguments) : call.function.arguments;
					next = await current.registry.execute(name, args);
				} catch (error) { next = { error: error.message }; }
				if (next && next.confirmation) {
					pending = { token: randomUUID(), action: next.confirmation, registry: current.registry, step: current.step, remaining: current.remaining.slice(index + 1) };
					return { type: 'confirmation', token: pending.token, title: next.confirmation.title, lines: next.confirmation.lines };
				}
				messages.push({ role: 'tool', tool_name: name, content: JSON.stringify(next) });
			}
			return await run(current.registry, current.step);
		} finally { busy = false; }
	}
	return { ask, decide, reset };
}

function activityLabel(name) {
	const labels = { searchStudents: 'Suche Schüler …', getStudent: 'Lese Schülerdaten …', getStudentGrades: 'Lese Noten …', getStudentParticipation: 'Lese Mitarbeit …', getClassStudents: 'Lese Klasse …', getCurrentSession: 'Prüfe Session …', createRepetitionGrade: 'Prüfe Note …', addParticipationPoints: 'Prüfe Mitarbeit …' };
	return labels[name] || 'Führe App-Aktion aus …';
}

module.exports = { createAiAssistantService, SYSTEM_PROMPT };
