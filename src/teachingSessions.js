const fs = require('fs');
const path = require('path');

const MIN_VISITED_SESSIONS = 2;
const MIN_PARTICIPATION_RATINGS = 3;

function createTeachingSessionStore(filePath) {
	function readData() {
		try {
			if (!fs.existsSync(filePath)) return { sessions: [] };
			const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));
			return { sessions: Array.isArray(data.sessions) ? data.sessions : [] };
		} catch (error) {
			return { sessions: [] };
		}
	}

	function writeData(data) {
		fs.mkdirSync(path.dirname(filePath), { recursive: true });
		fs.writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf8');
	}

	function getActiveSession() {
		const sessions = readData().sessions;
		return sessions.slice().reverse().find(session => !session.endedAt) || null;
	}

	function startSession(className, persons, absentIds) {
		if (!className || !Array.isArray(persons) || persons.length === 0) {
			return { success: false, reason: 'invalid-session-data' };
		}
		if (getActiveSession()) return { success: false, reason: 'session-already-active' };

		const absentIdSet = new Set((absentIds || []).map(normalizeId));
		const snapshots = persons.map(person => ({ id: person.id, name: String(person.name || '') }));
		const startedAt = new Date().toISOString();
		const session = {
			id: createId(),
			className: String(className),
			startedAt,
			date: startedAt,
			endedAt: null,
			presentStudents: snapshots.filter(person => !absentIdSet.has(normalizeId(person.id))),
			absentStudents: snapshots.filter(person => absentIdSet.has(normalizeId(person.id))),
			participationEntries: [],
			repetitionEntries: []
		};
		const data = readData();
		data.sessions.push(session);
		writeData(data);
		return { success: true, session };
	}

	function endSession(sessionId) {
		return updateSession(sessionId, session => {
			if (session.endedAt) return false;
			session.endedAt = new Date().toISOString();
			return true;
		});
	}

	function addParticipation(sessionId, personId, points) {
		const numericPoints = Number(points);
		return updateSession(sessionId, session => {
			const person = session.presentStudents.find(candidate => normalizeId(candidate.id) === normalizeId(personId));
			if (!person || !Number.isInteger(numericPoints) || numericPoints < 1 || numericPoints > 6) return false;
			session.participationEntries.push({
				id: createId(),
				studentId: person.id,
				studentName: person.name,
				points: numericPoints,
				timestamp: new Date().toISOString(),
				sessionId: session.id
			});
			return true;
		});
	}

	function addRepetition(sessionId, entry) {
		return updateSession(sessionId, session => {
			if (!entry || !entry.personName) return false;
			session.repetitionEntries.push({
				id: createId(),
				studentId: entry.personId,
				studentName: entry.personName,
				type: entry.type || 'grade',
				grade: entry.grade,
				timestamp: new Date().toISOString(),
				sessionId: session.id
			});
			return true;
		});
	}

	function updateSession(sessionId, mutator) {
		const data = readData();
		const session = data.sessions.find(candidate => candidate.id === sessionId);
		if (!session || !mutator(session)) return { success: false, reason: 'invalid-session-operation' };
		writeData(data);
		return { success: true, session };
	}

	function getParticipationSummary(className) {
		const sessions = readData().sessions.filter(session => !className || session.className === className);
		const students = new Map();
		sessions.forEach(session => {
			[...(session.presentStudents || []), ...(session.absentStudents || [])].forEach(person => {
				const key = studentKey(session.className, person);
				if (!students.has(key)) students.set(key, createSummaryRow(session.className, person));
			});
			(session.presentStudents || []).forEach(person => {
				const key = studentKey(session.className, person);
				if (!students.has(key)) students.set(key, createSummaryRow(session.className, person));
				students.get(key).visitedSessions += 1;
			});
			const participants = new Set();
			(session.participationEntries || []).forEach(entry => {
				const person = { id: entry.studentId, name: entry.studentName };
				const key = studentKey(session.className, person);
				if (!students.has(key)) students.set(key, createSummaryRow(session.className, person));
				const row = students.get(key);
				row.ratingCount += 1;
				row.pointsTotal += Number(entry.points) || 0;
				participants.add(key);
			});
			participants.forEach(key => { students.get(key).sessionsWithParticipation += 1; });
		});

		return Array.from(students.values()).map(row => ({
			className: row.className,
			studentId: row.studentId,
			studentName: row.studentName,
			visitedSessions: row.visitedSessions,
			sessionsWithParticipation: row.sessionsWithParticipation,
			ratingCount: row.ratingCount,
			averagePoints: row.ratingCount ? row.pointsTotal / row.ratingCount : null,
			participationRate: row.visitedSessions ? row.sessionsWithParticipation / row.visitedSessions : null,
			dataSufficient: row.visitedSessions >= MIN_VISITED_SESSIONS && row.ratingCount >= MIN_PARTICIPATION_RATINGS
		})).sort((a, b) => a.studentName.localeCompare(b.studentName, 'de'));
	}

	return { readData, getActiveSession, startSession, endSession, addParticipation, addRepetition, getParticipationSummary };
}

function createSummaryRow(className, person) {
	return { className, studentId: person.id, studentName: person.name, visitedSessions: 0, sessionsWithParticipation: 0, ratingCount: 0, pointsTotal: 0 };
}

function studentKey(className, person) {
	return `${className}::${normalizeId(person.id)}`;
}

function normalizeId(value) {
	return String(value);
}

function createId() {
	return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

module.exports = { createTeachingSessionStore, MIN_VISITED_SESSIONS, MIN_PARTICIPATION_RATINGS };
