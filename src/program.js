const { init, read, read_editor_persons, write_grade, write_joker, edit_persons, apply_entries, undo_entry, getJokerMigrationStatus, migrate_jokers } = require('./excelReader.js');
const { getAppSettings, getWiggersRulePenalties, saveWiggersRulePenalty, removeWiggersRulePenalty } = require('./storage.js');
let cls;
let persons;
let person;
let filePath;
const absences = new Set();
const wiggersRulePenalties = new Map();

// set the file
function setFile(file, callback) {
	filePath = file;
	init(file, (r) => {
		callback(r);
	});
}

// set the class
function setClass(clss, callback) {
	cls = clss;
	persons = read(cls);
	callback(!persons);
}

// get all persons for manual selection
function getPersons() {
	return filterAbsentPersons(persons);
}

function getEditorPersons() {
	if (!cls) return [];
	return read_editor_persons(cls) || [];
}

function getAbsencePersons() {
	return getEditorPersons().map(e => ({
		id: e.id,
		name: e.name,
		grades: e.grades,
		joker: Number(e.joker || 0),
		absent: isPersonAbsent(e.id)
	}));
}

function setAbsences(absentIds, callback) {
	if (!cls) {
		callback({ success: false, reason: 'no-class-selected' });
		return;
	}

	const absentIdSet = new Set((absentIds || []).map(id => Number(id)));
	const prefix = getAbsenceKeyPrefix();
	Array.from(absences).forEach(key => {
		if (key.startsWith(prefix)) absences.delete(key);
	});
	absentIdSet.forEach(id => {
		if (Number.isFinite(id)) absences.add(getAbsenceKey(id));
	});

	callback({
		success: true,
		count: absentIdSet.size,
		persons: getAbsencePersons()
	});
}

function getProbabilities() {
	const weightedPersons = getWeightedPersons();
	const totalWeight = weightedPersons.reduce((sum, e) => sum + e.weight, 0);

	return weightedPersons.map(e => ({
		id: e.person.id,
		name: e.person.name,
		grades: e.person.grades,
		joker: Number(e.person.joker || 0),
		weight: e.weight,
		boostedNeverSelected: e.boostedNeverSelected,
		wiggersRuleActive: e.wiggersRuleActive,
		wiggersRuleRemainingMinutes: e.wiggersRuleRemainingMinutes,
		probability: totalWeight > 0 ? e.weight / totalWeight : 0
	}));
}

// select a specific person manually
function selectSpecificPerson(personId) {
	person = filterAbsentPersons(persons).find(p => p.id === personId);
	return person ? getPersonResult(person) : null;
}

// select a random person (based on the amount of grades)
function selectPerson() {
	if (!persons || persons.length === 0) return null;

	const weightedPersons = getWeightedPersons();
	if (weightedPersons.length === 0) return null;
	const totalWeight = weightedPersons.reduce((sum, e) => sum + e.weight, 0);
	let randomWeight = Math.random() * totalWeight;
	person = null;

	for (let i = 0; i < weightedPersons.length; i++) {
		randomWeight -= weightedPersons[i].weight;
		if (randomWeight <= 0) {
			person = weightedPersons[i].person;
			break;
		}
	}

	if (!person) person = weightedPersons[weightedPersons.length - 1].person;
	return getPersonResult(person);
}

function getWeightedPersons() {
	const settings = getAppSettings();
	const storedPenalties = getWiggersRulePenalties();
	const decreaseFactor = settings.probabilityDecreaseFactor;

	if (!persons || persons.length === 0) return [];

	return filterAbsentPersons(persons).map(e => {
		let weight = Math.pow(decreaseFactor, 6 - e.grades);
		const boostedNeverSelected = settings.boostNeverSelected && e.grades === 0;
		if (boostedNeverSelected) {
			weight *= settings.neverSelectedBoostFactor;
		}
		const wiggersRule = getWiggersRuleState(e.id, settings, storedPenalties);
		if (wiggersRule.active) {
			weight *= settings.wiggersRulePenaltyFactor;
		}

		return {
			person: e,
			weight: weight,
			boostedNeverSelected: boostedNeverSelected,
			wiggersRuleActive: wiggersRule.active,
			wiggersRuleRemainingMinutes: wiggersRule.remainingMinutes
		};
	});
}

function getPersonResult(selectedPerson) {
	const remainingJokers = Number(selectedPerson.joker || 0);
	const jokerUnavailable = remainingJokers <= 0;
	return [selectedPerson.name, jokerUnavailable];
}

// save the new grade to the file
function saveGrade(grade, callback) {
	const settings = getAppSettings();
	const awardExtraJoker = settings.extraJokerAfterThreeGrades && person.grades < 3 && person.grades + 1 >= 3;
	const entry = {
		type: 'grade',
		className: cls,
		filePath: filePath,
		personId: person.id,
		personName: person.name,
		grade: grade,
		awardExtraJoker: awardExtraJoker,
		date: getFormattedDate()
	};

	write_grade(cls, person, grade, (p, result) => {
		if (p) persons = p;
		callback(undefined, {
			...entry,
			excelWriteSucceeded: !!(result && result.success),
			excelWriteError: result && result.error,
			excelWriteReason: result && result.reason
		})
	}, awardExtraJoker);
}

// save the joker to the file
function setJoker(callback) {
	const entry = {
		type: 'joker',
		className: cls,
		filePath: filePath,
		personId: person.id,
		personName: person.name,
		date: getFormattedDate()
	};

    write_joker(cls, person, (p, result) => {
        if (p) persons = p;
		if (result && (result.success || ['excel-locked', 'excel-write-failed'].includes(result.reason))) {
			activateWiggersRule(person.id);
		}
		if (callback) {
			callback(undefined, {
				...entry,
				excelWriteSucceeded: !!(result && result.success),
				excelWriteError: result && result.error,
				excelWriteReason: result && result.reason
			});
		}
    });

}

function saveEditorPersons(editorPersons, callback) {
	if (!cls) {
		callback({ success: false, reason: 'no-class-selected' });
		return;
	}

	edit_persons(cls, editorPersons, (p, result) => {
		if (result && result.success) persons = p || [];
		callback(result || { success: false, reason: 'unknown' });
	});
}

function migrateJokers(callback) {
	migrate_jokers(result => {
		if (result && result.success && cls) {
			persons = read(cls);
		}
		callback(result);
	});
}

function applyPendingExcelEntries(entries, callback) {
	if (!cls) {
		callback({ success: false, reason: 'no-class-selected', applied: [] });
		return;
	}

	apply_entries(cls, entries, (p, result) => {
		if (p) persons = p;
		callback(result || { success: false, reason: 'unknown', applied: [] });
	});
}

function undoLastExcelEntry(entry, callback) {
	if (!cls) {
		callback({ success: false, reason: 'no-class-selected' });
		return;
	}

	undo_entry(cls, entry, (p, result) => {
		if (p) persons = p;
		if (result && result.success && entry && entry.type === 'joker') {
			deactivateWiggersRule(entry.personId);
		}
		callback(result || { success: false, reason: 'unknown' });
	});
}

function redoExcelEntry(entry, callback) {
	if (!cls) {
		callback({ success: false, reason: 'no-class-selected', applied: [] });
		return;
	}

	apply_entries(cls, [entry], (p, result) => {
		if (p) persons = p;
		if (result && result.success && entry && entry.type === 'joker') {
			activateWiggersRule(entry.personId);
		}
		callback(result || { success: false, reason: 'unknown', applied: [] });
	});
}

function reloadExcel(callback) {
	if (!filePath) {
		callback({ success: false, reason: 'no-file-selected' });
		return;
	}

	init(filePath, worksheets => {
		if (!worksheets) {
			callback({ success: false, reason: 'excel-read-failed' });
			return;
		}

		if (cls) {
			persons = read(cls);
		}

		callback({
			success: true,
			worksheets: worksheets,
			filePath: filePath,
			className: cls,
			hasClass: !!cls
		});
	});
}

function getCurrentFilePath() {
	return filePath;
}

function getFormattedDate() {
	const currentDate = new Date();
	return `${String(currentDate.getDate()).padStart(2, '0')}.${String(currentDate.getMonth() + 1).padStart(2, '0')}.${currentDate.getFullYear()}`;
}

function filterAbsentPersons(personList) {
	if (!personList) return personList;
	return personList.filter(e => !isPersonAbsent(e.id));
}

function isPersonAbsent(personId) {
	if (!cls) return false;
	return absences.has(getAbsenceKey(personId));
}

function getAbsenceKeyPrefix() {
	return `${cls}::${getFormattedDate()}::`;
}

function getAbsenceKey(personId) {
	return `${getAbsenceKeyPrefix()}${Number(personId)}`;
}

function activateWiggersRule(personId) {
	const settings = getAppSettings();
	if (!settings.wiggersRuleEnabled) return;
	const expiresAt = Date.now() + (settings.wiggersRuleDurationMinutes * 60 * 1000);
	const key = getWiggersRuleKey(personId);
	wiggersRulePenalties.set(key, expiresAt);
	saveWiggersRulePenalty(key, expiresAt);
}

function activateWiggersRuleForEntry(entry) {
	if (entry && entry.type === 'joker') {
		activateWiggersRule(entry.personId);
	}
}

function deactivateWiggersRule(personId) {
	const key = getWiggersRuleKey(personId);
	wiggersRulePenalties.delete(key);
	removeWiggersRulePenalty(key);
}

function deactivateWiggersRuleForEntry(entry) {
	if (entry && entry.type === 'joker') {
		deactivateWiggersRule(entry.personId);
	}
}

function getWiggersRuleState(personId, settings, storedPenalties) {
	if (!settings.wiggersRuleEnabled) {
		return { active: false, remainingMinutes: 0 };
	}

	const key = getWiggersRuleKey(personId);
	const expiresAt = wiggersRulePenalties.get(key) || (storedPenalties && storedPenalties[key]);
	if (!expiresAt) return { active: false, remainingMinutes: 0 };

	const remainingMs = expiresAt - Date.now();
	if (remainingMs <= 0) {
		wiggersRulePenalties.delete(key);
		removeWiggersRulePenalty(key);
		return { active: false, remainingMinutes: 0 };
	}

	wiggersRulePenalties.set(key, expiresAt);

	return {
		active: true,
		remainingMinutes: Math.ceil(remainingMs / 60000)
	};
}

function getWiggersRuleKey(personId) {
	return `${filePath || ''}::${cls || ''}::${Number(personId)}`;
}

module.exports = {
	setFile: setFile,
	setClass: setClass,
	selectPerson: selectPerson,
	saveGrade: saveGrade,
	setJoker: setJoker,
	saveEditorPersons: saveEditorPersons,
	applyPendingExcelEntries: applyPendingExcelEntries,
	reloadExcel: reloadExcel,
	getCurrentFilePath: getCurrentFilePath,
	getJokerMigrationStatus: getJokerMigrationStatus,
	migrateJokers: migrateJokers,
	getPersons: getPersons,
	getEditorPersons: getEditorPersons,
	getProbabilities: getProbabilities,
	selectSpecificPerson: selectSpecificPerson,
	getAbsencePersons: getAbsencePersons,
	setAbsences: setAbsences,
	undoLastExcelEntry: undoLastExcelEntry,
	redoExcelEntry: redoExcelEntry,
	activateWiggersRuleForEntry: activateWiggersRuleForEntry,
	deactivateWiggersRuleForEntry: deactivateWiggersRuleForEntry,
}
