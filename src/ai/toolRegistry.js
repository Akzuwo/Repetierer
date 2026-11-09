const MAX_RESULTS = 20;

function createAiToolRegistry({ program, sessions, commitGrade, notifyNavigation, getUiContext, notifyWrite = () => {} }) {
	const searched = new Set();
	const definitions = [
		tool('searchStudents', 'Suche Personen nach Namen. Bei mehrdeutigen Namen muss der Benutzer präzisieren.', { query: str(), classId: str(true) }),
		tool('getStudent', 'Details einer Person anhand Klassenname und ID.', { classId: str(), studentId: integer() }),
		tool('getStudentGrades', 'Repetitionsnoten einer Person aus Excel.', { classId: str(), studentId: integer() }),
		tool('getStudentParticipation', 'Mitarbeit einer Person aus Unterrichtssitzungen.', { classId: str(), studentId: integer() }),
		tool('getClassStudents', 'Personen einer Klasse, einschliesslich Anzahl Repetitionsnoten.', { classId: str() }),
		tool('getCurrentSession', 'Aktive Unterrichtssitzung.', {}),
		tool('getCurrentAppContext', 'Kleiner aktueller UI-Kontext.', {}),
		tool('openStudent', 'Öffnet eine Person in der App.', { classId: str(), studentId: integer() }),
		tool('openClass', 'Öffnet eine Klasse in der App.', { classId: str() }),
		tool('openEvaluation', 'Öffnet die Auswertung einer Klasse.', { classId: str() }),
		tool('openCurrentSession', 'Öffnet die aktive Sitzung.', {}),
		tool('createRepetitionGrade', 'Trägt eine Schweizer Repetitionsnote von 1 bis 6 ein. Benötigt eine aktive Sitzung.', { classId: str(), studentId: integer(), grade: number() }),
		tool('addParticipationPoints', 'Erfasst 1 bis 6 Mitarbeitspunkte in der aktiven Sitzung.', { classId: str(), studentId: integer(), points: integer() })
	];
	const writes = new Set(['createRepetitionGrade', 'addParticipationPoints']);
	const key = (classId, studentId) => `${classId}::${studentId}`;
	function classes() { return program.getClasses(); }
	function validClass(classId) {
		if (!classes().includes(classId)) throw new Error('Unbekannte Klasse.');
		return classId;
	}
	function student(classId, studentId) {
		validClass(classId);
		const context = getUiContext();
		if (!searched.has(key(classId, studentId)) && !(context.classId === classId && context.studentId === studentId)) throw new Error('Person zuerst mit searchStudents eindeutig suchen; keine Schüler-ID raten.');
		const person = program.getClassPersons(classId).find(item => item.id === studentId);
		if (!person) throw new Error('Unbekannte Person in dieser Klasse.');
		return person;
	}
	function sessionRows(classId) {
		return sessions.readData().sessions.filter(item => item.className === classId);
	}
	function validate(name, args) {
		const definition = definitions.find(item => item.function.name === name);
		if (!definition) throw new Error(`Unbekanntes Tool: ${name}`);
		if (!args || typeof args !== 'object' || Array.isArray(args)) throw new Error('Tool-Argumente müssen ein Objekt sein.');
		const props = definition.function.parameters.properties;
		for (const property of Object.keys(args)) if (!props[property]) throw new Error(`Unbekanntes Argument: ${property}`);
		for (const [property, schema] of Object.entries(props)) {
			const value = args[property];
			if (value === undefined && !definition.function.parameters.required.includes(property)) continue;
			if (schema.type === 'string' && (typeof value !== 'string' || !value.trim() || value.length > 120)) throw new Error(`Ungültiges Argument: ${property}`);
			if (schema.type === 'integer' && (!Number.isInteger(value) || value < 0)) throw new Error(`Ungültiges Argument: ${property}`);
			if (schema.type === 'number' && (typeof value !== 'number' || !Number.isFinite(value))) throw new Error(`Ungültiges Argument: ${property}`);
		}
		return args;
	}
	function prepareWrite(name, args) {
		const person = student(args.classId, args.studentId);
		const active = sessions.getActiveSession();
		if (!active || active.className !== args.classId || program.getCurrentClass() !== args.classId) throw new Error('Für diese Klasse muss eine aktive Sitzung geöffnet sein.');
		if (!active.presentStudents.some(item => item.id === args.studentId)) throw new Error('Die Person ist in der aktiven Sitzung nicht anwesend.');
		if (name === 'createRepetitionGrade') {
			if (args.grade < 1 || args.grade > 6) throw new Error('Note muss zwischen 1 und 6 liegen.');
			if (person.grades >= 6) throw new Error('Für diese Person sind bereits sechs Repetitionsnoten erfasst.');
			return { name, args, title: 'Repetitionsnote eintragen', lines: [person.name, `Klasse ${args.classId}`, `Note ${args.grade}`] };
		}
		if (args.points < 1 || args.points > 6) throw new Error('Mitarbeitspunkte müssen ganzzahlig zwischen 1 und 6 liegen.');
		return { name, args, title: 'Mitarbeit erfassen', lines: [person.name, `Klasse ${args.classId}`, `${args.points} Punkte`] };
	}
	async function execute(name, rawArgs) {
		const args = validate(name, rawArgs);
		if (writes.has(name)) return { confirmation: prepareWrite(name, args) };
		switch (name) {
			case 'searchStudents': {
				if (args.classId) validClass(args.classId);
				const query = args.query.toLocaleLowerCase('de-CH');
				const matches = (args.classId ? [args.classId] : classes()).flatMap(classId => program.getClassPersons(classId).filter(person => person.name.toLocaleLowerCase('de-CH').includes(query)).map(person => ({ classId, studentId: person.id, name: person.name })));
				if (matches.length === 1) searched.add(key(matches[0].classId, matches[0].studentId));
				return { matches: matches.slice(0, MAX_RESULTS), total: matches.length, ambiguous: matches.length > 1 };
			}
			case 'getStudent': { const person = student(args.classId, args.studentId); return { classId: args.classId, studentId: person.id, name: person.name, gradeCount: person.grades, remainingJokers: person.joker }; }
			case 'getStudentGrades': return program.getStudentGradeHistory(args.classId, student(args.classId, args.studentId).id);
			case 'getStudentParticipation': { const person = student(args.classId, args.studentId); return { classId: args.classId, studentId: person.id, entries: sessionRows(args.classId).flatMap(session => (session.participationEntries || []).filter(entry => entry.studentId === person.id).map(entry => ({ points: entry.points, timestamp: entry.timestamp, sessionId: session.id }))) }; }
			case 'getClassStudents': return { classId: validClass(args.classId), students: program.getClassPersons(args.classId).map(person => ({ studentId: person.id, name: person.name, gradeCount: person.grades })) };
			case 'getCurrentSession': { const active = sessions.getActiveSession(); return active ? { id: active.id, classId: active.className, startedAt: active.startedAt, presentCount: active.presentStudents.length, absentCount: active.absentStudents.length } : null; }
			case 'getCurrentAppContext': return getUiContext();
			case 'openStudent': { const person = student(args.classId, args.studentId); notifyNavigation({ page: 'student', classId: args.classId, studentId: person.id, name: person.name, gradeCount: person.grades, remainingJokers: person.joker }); return { opened: true, classId: args.classId, studentId: person.id, name: person.name }; }
			case 'openClass': validClass(args.classId); notifyNavigation({ page: 'class', classId: args.classId }); return { opened: true, classId: args.classId };
			case 'openEvaluation': validClass(args.classId); notifyNavigation({ page: 'evaluation', classId: args.classId }); return { opened: true, classId: args.classId };
			case 'openCurrentSession': { const active = sessions.getActiveSession(); if (!active) throw new Error('Keine aktive Sitzung.'); notifyNavigation({ page: 'session', classId: active.className }); return { opened: true, sessionId: active.id }; }
		}
	}
	async function commit(action) {
		validate(action.name, action.args);
		prepareWrite(action.name, action.args);
		const { classId, studentId } = action.args;
		if (action.name === 'createRepetitionGrade') return commitGrade(classId, studentId, action.args.grade);
		const active = sessions.getActiveSession();
		const result = sessions.addParticipation(active.id, studentId, action.args.points);
		if (!result.success) throw new Error('Mitarbeit konnte nicht erfasst werden.');
		notifyWrite('participation', result.session);
		return { saved: true, classId, studentId, points: action.args.points, session: result.session };
	}
	return { definitions, validate, execute, commit };
}

function str(optional = false) { return { type: 'string', optional }; }
function integer() { return { type: 'integer' }; }
function number() { return { type: 'number' }; }
function tool(name, description, properties) { return { type: 'function', function: { name, description, parameters: { type: 'object', properties, required: Object.keys(properties).filter(key => !properties[key].optional), additionalProperties: false } } }; }

module.exports = { createAiToolRegistry };
