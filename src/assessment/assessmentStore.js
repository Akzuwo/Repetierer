const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const { matchResponse, normalizeEmail, normalizeName, selectLatestResponses } = require('./matching.js');

const SCHEMA_VERSION = 2;
const DEFAULT_CRITERIA = [
	{ id: 'participation', label: 'Aktive Beteiligung' },
	{ id: 'preparation', label: 'Vorbereitung' },
	{ id: 'quality', label: 'Qualität der Beiträge' },
	{ id: 'reliability', label: 'Zuverlässigkeit' }
];

class AssessmentStore {
	constructor(filePath) {
		this.filePath = filePath;
		this.data = this.read();
	}

	read() {
		try {
			if (!fs.existsSync(this.filePath)) return emptyData();
			return migrate(JSON.parse(fs.readFileSync(this.filePath, 'utf8')));
		} catch (error) {
			const wrapped = new Error('Mitmachnoten-Daten konnten nicht gelesen werden.');
			wrapped.code = 'ASSESSMENT_DATA_INVALID';
			wrapped.cause = error;
			throw wrapped;
		}
	}

	save() {
		this.data.updatedAt = new Date().toISOString();
		atomicWriteJson(this.filePath, this.data);
	}

	getClassContext(workbookPath, className, sourceStudents) {
		const classRecord = this.ensureClass(workbookPath, className, sourceStudents);
		return clone(classRecord);
	}

	ensureClass(workbookPath, className, sourceStudents) {
		const workbookId = hashPath(workbookPath);
		if (!this.data.workbooks[workbookId]) {
			this.data.workbooks[workbookId] = {
				id: workbookId,
				path: path.resolve(workbookPath),
				classes: {}
			};
		}
		const workbook = this.data.workbooks[workbookId];
		workbook.path = path.resolve(workbookPath);
		const classId = hashValue(`${workbookId}\n${className}`);
		if (!workbook.classes[classId]) {
			workbook.classes[classId] = { id: classId, name: className, students: [], rounds: [] };
		}
		const classRecord = workbook.classes[classId];
		classRecord.name = className;
		syncStudents(classRecord, sourceStudents);
		return classRecord;
	}

	createRound(workbookPath, className, sourceStudents, input) {
		const classRecord = this.ensureClass(workbookPath, className, sourceStudents);
		const title = String(input && input.title || '').trim();
		if (!title) throw validationError('Der Titel der Beurteilungsrunde fehlt.');
		if (classRecord.rounds.some(round => normalizeName(round.title) === normalizeName(title))) {
			throw validationError('Für diese Klasse existiert bereits eine Runde mit diesem Titel.');
		}
		const deadline = normalizeDeadline(input && input.deadline);
		const round = {
			id: crypto.randomUUID(),
			classId: classRecord.id,
			externalRoundId: createExternalRoundId(className, title),
			title,
			createdAt: new Date().toISOString(),
			deadline,
			criteria: clone(DEFAULT_CRITERIA),
			lastSyncAt: null,
			lastSyncStatus: 'never',
			lastSyncResult: null,
			assessments: {},
			externalResponses: [],
			responseAssignments: {},
			preferredResponseIds: {},
			unmatchedResponses: [],
			duplicates: []
		};
		for (const student of classRecord.students) round.assessments[student.id] = emptyAssessment(student.id);
		classRecord.rounds.unshift(round);
		this.save();
		return clone(round);
	}

	getRound(workbookPath, className, sourceStudents, roundId) {
		const classRecord = this.ensureClass(workbookPath, className, sourceStudents);
		const round = classRecord.rounds.find(item => item.id === roundId);
		if (!round) throw notFoundError('Beurteilungsrunde nicht gefunden.');
		ensureRoundStudents(round, classRecord.students);
		return { classRecord, round };
	}

	deleteRound(workbookPath, className, sourceStudents, roundId) {
		const classRecord = this.ensureClass(workbookPath, className, sourceStudents);
		const index = classRecord.rounds.findIndex(item => item.id === roundId);
		if (index < 0) throw notFoundError('Beurteilungsrunde nicht gefunden.');
		const [round] = classRecord.rounds.splice(index, 1);
		this.save();
		return {
			id: round.id,
			externalRoundId: round.externalRoundId,
			hadSentResults: Object.values(round.assessments || {}).some(assessment => assessment.mailStatus === 'sent')
		};
	}

	applyResponses(workbookPath, className, sourceStudents, roundId, responses, syncDetails = {}) {
		const { classRecord, round } = this.getRound(workbookPath, className, sourceStudents, roundId);
		const known = new Set(round.externalResponses.map(getResponseId));
		let added = 0;
		for (const response of Array.isArray(responses) ? responses : []) {
			const responseId = getResponseId(response);
			if (!responseId || known.has(responseId)) continue;
			round.externalResponses.push(normalizeStoredResponse(response, round.externalRoundId));
			known.add(responseId);
			added++;
		}
		rebuildStudentResponses(classRecord, round);
		round.lastSyncAt = new Date().toISOString();
		round.lastSyncStatus = 'success';
		round.lastSyncResult = {
			added,
			invalid: Number(syncDetails.invalid || 0),
			ignoredRound: Number(syncDetails.ignoredRound || 0)
		};
		this.save();
		return clone(round);
	}

	markSyncFailed(workbookPath, className, sourceStudents, roundId) {
		const { round } = this.getRound(workbookPath, className, sourceStudents, roundId);
		round.lastSyncStatus = 'failed';
		this.save();
	}

	assignResponse(workbookPath, className, sourceStudents, roundId, responseId, studentId) {
		const { classRecord, round } = this.getRound(workbookPath, className, sourceStudents, roundId);
		const response = round.externalResponses.find(item => getResponseId(item) === responseId);
		if (!response) throw notFoundError('Umfrageantwort nicht gefunden.');
		if (!classRecord.students.some(student => student.id === studentId)) throw notFoundError('Schüler nicht gefunden.');
		round.responseAssignments[responseId] = studentId;
		rebuildStudentResponses(classRecord, round);
		this.save();
		return clone(round);
	}

	selectResponse(workbookPath, className, sourceStudents, roundId, responseId) {
		const { classRecord, round } = this.getRound(workbookPath, className, sourceStudents, roundId);
		const response = round.externalResponses.find(item => getResponseId(item) === responseId);
		if (!response) throw notFoundError('Umfrageantwort nicht gefunden.');
		const email = normalizeEmail(response.email);
		if (!email) throw validationError('Die Antwort besitzt keine gültige E-Mail-Zuordnung.');
		round.preferredResponseIds[email] = responseId;
		rebuildStudentResponses(classRecord, round);
		this.save();
		return clone(round);
	}

	saveTeacherAssessment(workbookPath, className, sourceStudents, roundId, studentId, answers, comment) {
		const { round } = this.getRound(workbookPath, className, sourceStudents, roundId);
		const assessment = round.assessments[studentId];
		if (!assessment) throw notFoundError('Schüler nicht gefunden.');
		const normalizedAnswers = {};
		for (const criterion of round.criteria) {
			const value = Number(answers && answers[criterion.id]);
			if (!Number.isInteger(value) || value < 1 || value > 5) {
				throw validationError(`Für „${criterion.label}“ ist ein Wert von 1 bis 5 erforderlich.`);
			}
			normalizedAnswers[criterion.id] = value;
		}
		assessment.teacherAnswers = normalizedAnswers;
		assessment.teacherComment = String(comment || '').trim();
		assessment.teacherCompletedAt = new Date().toISOString();
		this.save();
		return clone(round);
	}

	setMailStatus(workbookPath, className, sourceStudents, roundId, studentId, status, details) {
		const { round } = this.getRound(workbookPath, className, sourceStudents, roundId);
		const assessment = round.assessments[studentId];
		if (!assessment) throw notFoundError('Schüler nicht gefunden.');
		assessment.mailStatus = status;
		assessment.mailAttempts = Number(assessment.mailAttempts || 0) + (details && details.incrementAttempt ? 1 : 0);
		assessment.sentAt = status === 'sent' ? new Date().toISOString() : assessment.sentAt || null;
		assessment.lastMailError = details && details.error ? String(details.error).slice(0, 500) : null;
		this.save();
		return clone(assessment);
	}
}

function applyStudentResponse(classRecord, round, studentId, response, method) {
	const assessment = round.assessments[studentId] || emptyAssessment(studentId);
	assessment.studentEmail = normalizeEmail(response.email);
	assessment.studentAnswers = clone(response.answers || {});
	assessment.studentComment = String(response.comment || '').trim();
	assessment.submittedAt = response.timestamp || response.lastSubmittedTime || response.submittedAt || response.createTime || null;
	assessment.externalResponseId = getResponseId(response) || null;
	delete assessment.responseId;
	assessment.matchMethod = method;
	round.assessments[studentId] = assessment;
	const student = classRecord.students.find(item => item.id === studentId);
	if (student && assessment.studentEmail) student.email = assessment.studentEmail;
}

function rebuildStudentResponses(classRecord, round) {
	const { selected, duplicates } = selectLatestResponses(round.externalResponses, round.preferredResponseIds);
	const unmatched = [];
	for (const response of selected) {
		const responseId = getResponseId(response);
		const manualStudentId = round.responseAssignments[responseId];
		const match = manualStudentId ? { status: 'matched', studentId: manualStudentId, method: 'manual' } : matchResponse(response, classRecord.students);
		if (match.status !== 'matched' || !classRecord.students.some(student => student.id === match.studentId)) {
			unmatched.push(Object.assign({}, response, { candidateIds: match.candidateIds || [] }));
			continue;
		}
		applyStudentResponse(classRecord, round, match.studentId, response, match.method);
	}
	round.unmatchedResponses = unmatched;
	round.duplicates = duplicates;
}

function syncStudents(classRecord, sourceStudents) {
	const incoming = Array.isArray(sourceStudents) ? sourceStudents : [];
	for (const source of incoming) {
		const sourceId = Number(source.id);
		const normalized = normalizeName(source.name);
		let student = classRecord.students.find(item => item.sourceId === sourceId && normalizeName(item.name) === normalized);
		if (!student) student = classRecord.students.find(item => normalizeName(item.name) === normalized);
		if (!student) {
			student = { id: crypto.randomUUID(), sourceId, name: String(source.name), email: '' };
			classRecord.students.push(student);
		} else {
			student.sourceId = sourceId;
			student.name = String(source.name);
		}
	}
	const activeNames = new Set(incoming.map(source => normalizeName(source.name)));
	classRecord.students.forEach(student => { student.active = activeNames.has(normalizeName(student.name)); });
	classRecord.students = classRecord.students.filter(student => student.active || classRecord.rounds.some(round => round.assessments[student.id]));
}

function ensureRoundStudents(round, students) {
	if (!round.assessments) round.assessments = {};
	for (const student of students) {
		if (student.active !== false && !round.assessments[student.id]) round.assessments[student.id] = emptyAssessment(student.id);
	}
}

function emptyAssessment(studentId) {
	return {
		studentId,
		studentEmail: '',
		studentAnswers: null,
		teacherAnswers: null,
		studentComment: '',
		teacherComment: '',
		submittedAt: null,
		teacherCompletedAt: null,
		mailStatus: 'not_sent',
		sentAt: null,
		lastMailError: null,
		mailAttempts: 0
	};
}

function emptyData() {
	return { schemaVersion: SCHEMA_VERSION, updatedAt: null, workbooks: {} };
}

function migrate(input) {
	if (!input || typeof input !== 'object') return emptyData();
	if (!input.schemaVersion) input.schemaVersion = 1;
	if (input.schemaVersion > SCHEMA_VERSION) throw new Error('Neuere Mitmachnoten-Datenversion wird nicht unterstützt.');
	if (!input.workbooks || typeof input.workbooks !== 'object') input.workbooks = {};
	for (const workbook of Object.values(input.workbooks)) {
		if (!workbook.classes) workbook.classes = {};
		for (const classRecord of Object.values(workbook.classes)) {
			if (!Array.isArray(classRecord.students)) classRecord.students = [];
			if (!Array.isArray(classRecord.rounds)) classRecord.rounds = [];
			for (const round of classRecord.rounds) {
				round.classId = round.classId || classRecord.id;
				round.externalRoundId = round.externalRoundId || createExternalRoundId(classRecord.name, round.title, round.id);
				if (!round.criteria) round.criteria = clone(DEFAULT_CRITERIA);
				if (!round.assessments) round.assessments = {};
				if (!Array.isArray(round.externalResponses)) round.externalResponses = migrateResponses(round, classRecord);
				if (!round.responseAssignments || typeof round.responseAssignments !== 'object') round.responseAssignments = {};
				if (!round.preferredResponseIds || typeof round.preferredResponseIds !== 'object') round.preferredResponseIds = {};
				if (!Array.isArray(round.unmatchedResponses)) round.unmatchedResponses = [];
				if (!Array.isArray(round.duplicates)) round.duplicates = [];
				delete round.form;
				for (const assessment of Object.values(round.assessments)) {
					if (assessment.mailStatus === 'sending') {
						assessment.mailStatus = 'failed';
						assessment.lastMailError = 'Der vorherige Versand wurde unterbrochen.';
					}
				}
			}
		}
	}
	input.schemaVersion = SCHEMA_VERSION;
	return input;
}

function migrateResponses(round, classRecord) {
	const responses = [];
	for (const [studentId, assessment] of Object.entries(round.assessments || {})) {
		const responseId = assessment.externalResponseId || assessment.responseId;
		if (!responseId || !assessment.studentAnswers) continue;
		const student = classRecord.students.find(item => item.id === studentId);
		responses.push({
			externalResponseId: String(responseId),
			timestamp: assessment.submittedAt || round.lastSyncAt || round.createdAt,
			externalRoundId: round.externalRoundId,
			name: student ? student.name : '',
			email: assessment.studentEmail || '',
			answers: clone(assessment.studentAnswers),
			comment: assessment.studentComment || ''
		});
		assessment.externalResponseId = String(responseId);
		delete assessment.responseId;
	}
	return responses;
}

function normalizeStoredResponse(response, externalRoundId) {
	return {
		externalResponseId: getResponseId(response),
		timestamp: response.timestamp || response.lastSubmittedTime || response.submittedAt || response.createTime || null,
		externalRoundId,
		name: String(response.name || '').trim(),
		email: normalizeEmail(response.email),
		answers: clone(response.answers || {}),
		comment: String(response.comment || '').trim()
	};
}

function getResponseId(response) {
	return String(response && (response.externalResponseId || response.responseId) || '');
}

function createExternalRoundId(className, title, id) {
	const suffix = id ? String(id).replace(/[^a-z0-9]/gi, '').slice(0, 6).toLowerCase() : crypto.randomBytes(3).toString('hex');
	return `assessment-${slug(className)}-${slug(title)}-${suffix}`;
}

function slug(value) {
	return normalizeName(value).replace(/\s+/g, '-').replace(/^-+|-+$/g, '').slice(0, 40) || 'runde';
}

function atomicWriteJson(filePath, value) {
	fs.mkdirSync(path.dirname(filePath), { recursive: true });
	const tempPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
	try {
		fs.writeFileSync(tempPath, JSON.stringify(value, null, 2), { encoding: 'utf8', mode: 0o600 });
		fs.renameSync(tempPath, filePath);
	} finally {
		if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
	}
}

function normalizeDeadline(value) {
	if (!value) return null;
	const text = String(value).trim();
	if (!/^\d{4}-\d{2}-\d{2}$/.test(text) || Number.isNaN(Date.parse(`${text}T00:00:00Z`))) {
		throw validationError('Die Abgabefrist ist ungültig.');
	}
	return text;
}

function hashPath(value) {
	const resolved = path.resolve(String(value || ''));
	return hashValue(process.platform === 'win32' ? resolved.toLowerCase() : resolved);
}

function hashValue(value) {
	return crypto.createHash('sha256').update(String(value)).digest('hex').slice(0, 24);
}

function validationError(message) {
	const error = new Error(message);
	error.code = 'VALIDATION_ERROR';
	return error;
}

function notFoundError(message) {
	const error = new Error(message);
	error.code = 'NOT_FOUND';
	return error;
}

function clone(value) {
	return JSON.parse(JSON.stringify(value));
}

module.exports = {
	AssessmentStore,
	DEFAULT_CRITERIA,
	SCHEMA_VERSION,
	atomicWriteJson,
	createExternalRoundId,
	emptyAssessment,
	hashPath,
	migrate
};
