// Read-only smoke test against an installed local Ollama model and synthetic pupils.
const { createOllamaService } = require('../src/ai/ollamaService.js');
const { createAiToolRegistry } = require('../src/ai/toolRegistry.js');
const { createAiAssistantService } = require('../src/ai/assistantService.js');

const pupils = [{ id: 0, name: 'Anna Müller', grades: 2, joker: 1 }, { id: 1, name: 'Ben Müller', grades: 0, joker: 1 }, { id: 2, name: 'Lia Meier', grades: 1, joker: 0 }];
const session = { id: 'synthetic-session', className: '6a', startedAt: '2026-09-16T08:00:00Z', presentStudents: pupils, absentStudents: [], participationEntries: [], repetitionEntries: [] };
const program = {
	getClasses: () => ['6a'], getClassPersons: () => pupils, getCurrentClass: () => '6a',
	getStudentGradeHistory: (classId, studentId) => ({ classId, studentId, grades: studentId === 0 ? [{ grade: 5, date: '01.09.2026', position: 1 }, { grade: 5.5, date: '15.09.2026', position: 2 }] : studentId === 2 ? [{ grade: 4.5, date: '10.09.2026', position: 1 }] : [] })
};
const sessions = { getActiveSession: () => session, readData: () => ({ sessions: [session] }) };
const cases = [
	{ prompt: 'Welche Note hatte Anna Müller zuletzt?', tools: ['searchStudents', 'getStudentGrades'] },
	{ prompt: 'Zeig alle Repetitionsnoten von Anna Müller.', tools: ['searchStudents', 'getStudentGrades'] },
	{ prompt: 'Öffne Anna Müller.', tools: ['searchStudents', 'openStudent'] },
	{ prompt: 'Gib Anna Müller eine 5.5.', tools: ['searchStudents', 'createRepetitionGrade'], type: 'confirmation' },
	{ prompt: 'Welche Note hatte Müller zuletzt?', tools: ['searchStudents'] },
	{ prompt: 'Welche Note hatte Nora Frei zuletzt?', tools: ['searchStudents'] },
	{ prompt: 'Sag einfach Hallo.', tools: [] },
	{ prompt: 'Wer aus dieser Klasse hat noch keine Repetitionsnote?', tools: ['getClassStudents'] },
	{ prompt: 'Gib Anna Müller zwei Mitarbeitspunkte.', tools: ['searchStudents', 'addParticipationPoints'], type: 'confirmation' },
	{ prompt: 'Öffne die Auswertung der Klasse 6a.', tools: ['openEvaluation'] }
];

async function main() {
	const ollama = createOllamaService();
	if ((await ollama.status()).state !== 'ready') throw new Error('Zielmodell ist lokal nicht bereit.');
	let failed = 0;
	for (const item of cases.filter(candidate => !process.argv[2] || candidate.prompt.includes(process.argv[2]))) {
		const called = [];
		const assistant = createAiAssistantService({ ollama, makeRegistry: () => {
			const registry = createAiToolRegistry({ program, sessions, commitGrade: () => { throw new Error('Smoke test may not write.'); }, notifyNavigation: () => {}, getUiContext: () => ({ page: 'repetition', classId: '6a', studentId: null, sessionId: session.id }) });
			const execute = registry.execute;
			registry.execute = async (name, args) => { if (name !== 'getCurrentAppContext') called.push(name); return execute(name, args); };
			return registry;
		} });
		try {
			const result = await assistant.ask(item.prompt);
			if (result.type === 'confirmation') await assistant.decide(result.token, false);
			const ok = item.tools.every(name => called.includes(name)) && (!item.type || item.type === result.type) && (item.tools.length > 0 || called.length === 0);
			if (!ok) failed++;
			console.log(JSON.stringify({ prompt: item.prompt, ok, tools: called, result: result.type === 'answer' ? result.text : result.type }));
		} catch (error) { failed++; console.log(JSON.stringify({ prompt: item.prompt, ok: false, error: error.message })); }
	}
	if (failed) process.exitCode = 1;
}

main().catch(error => { console.error(error.message); process.exitCode = 1; });
