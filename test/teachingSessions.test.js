const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { createTeachingSessionStore } = require('../src/teachingSessions.js');

function makeStore(t) {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'repetierer-sessions-'));
	t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
	return createTeachingSessionStore(path.join(dir, 'sessions.json'));
}

test('starts and ends a session with attendance snapshots', t => {
	const store = makeStore(t);
	const result = store.startSession('L23a', [{ id: 1, name: 'Ada' }, { id: 2, name: 'Ben' }], [2]);
	assert.equal(result.success, true);
	assert.deepEqual(result.session.presentStudents.map(person => person.name), ['Ada']);
	assert.deepEqual(result.session.absentStudents.map(person => person.name), ['Ben']);
	assert.equal(store.getActiveSession().id, result.session.id);
	assert.equal(store.endSession(result.session.id).success, true);
	assert.equal(store.getActiveSession(), null);
});

test('stores multiple participation ratings and repetition entries in one session', t => {
	const store = makeStore(t);
	const session = store.startSession('L23a', [{ id: 1, name: 'Ada' }], []).session;
	assert.equal(store.addParticipation(session.id, 1, 4).success, true);
	assert.equal(store.addParticipation(session.id, 1, 6).success, true);
	assert.equal(store.addParticipation(session.id, 1, 0).success, false);
	assert.equal(store.addRepetition(session.id, { personId: 1, personName: 'Ada', type: 'grade', grade: 5 }).success, true);
	const stored = store.getActiveSession();
	assert.deepEqual(stored.participationEntries.map(entry => entry.points), [4, 6]);
	assert.equal(stored.repetitionEntries.length, 1);
	assert.equal(stored.participationEntries[0].sessionId, session.id);
});

test('calculates participation metrics without treating missing ratings as zero', t => {
	const store = makeStore(t);
	let session = store.startSession('L23a', [{ id: 1, name: 'Ada' }, { id: 2, name: 'Ben' }], []).session;
	store.addParticipation(session.id, 1, 3);
	store.addParticipation(session.id, 1, 5);
	store.endSession(session.id);
	session = store.startSession('L23a', [{ id: 1, name: 'Ada' }, { id: 2, name: 'Ben' }], []).session;
	store.addParticipation(session.id, 1, 4);
	const rows = store.getParticipationSummary('L23a');
	const ada = rows.find(row => row.studentName === 'Ada');
	const ben = rows.find(row => row.studentName === 'Ben');
	assert.equal(ada.visitedSessions, 2);
	assert.equal(ada.sessionsWithParticipation, 2);
	assert.equal(ada.ratingCount, 3);
	assert.equal(ada.averagePoints, 4);
	assert.equal(ada.participationRate, 1);
	assert.equal(ada.dataSufficient, true);
	assert.equal(ben.averagePoints, null);
	assert.equal(ben.participationRate, 0);
	assert.equal(ben.dataSufficient, false);
});
