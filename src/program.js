const { init, read, write_grade, write_joker, apply_entries, getJokerMigrationStatus, migrate_jokers } = require('./excelReader.js');
const { getAppSettings } = require('./storage.js');
let cls;
let persons;
let person;
let filePath;

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
	return persons;
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
		probability: totalWeight > 0 ? e.weight / totalWeight : 0
	}));
}

// select a specific person manually
function selectSpecificPerson(personId) {
	person = persons.find(p => p.id === personId);
	return person ? getPersonResult(person) : null;
}

// select a random person (based on the amount of grades)
function selectPerson() {
	if (!persons || persons.length === 0) return null;

	const weightedPersons = getWeightedPersons();
	const totalWeight = weightedPersons.reduce((sum, e) => sum + e.weight, 0);
	let randomWeight = Math.random() * totalWeight;

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
	const decreaseFactor = settings.probabilityDecreaseFactor;

	if (!persons || persons.length === 0) return [];

	return persons.map(e => {
		let weight = Math.pow(decreaseFactor, 6 - e.grades);
		const boostedNeverSelected = settings.boostNeverSelected && e.grades === 0;
		if (boostedNeverSelected) {
			weight *= settings.neverSelectedBoostFactor;
		}

		return {
			person: e,
			weight: weight,
			boostedNeverSelected: boostedNeverSelected
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

module.exports = {
	setFile: setFile,
	setClass: setClass,
	selectPerson: selectPerson,
	saveGrade: saveGrade,
	setJoker: setJoker,
	applyPendingExcelEntries: applyPendingExcelEntries,
	reloadExcel: reloadExcel,
	getCurrentFilePath: getCurrentFilePath,
	getJokerMigrationStatus: getJokerMigrationStatus,
	migrateJokers: migrateJokers,
	getPersons: getPersons,
	getProbabilities: getProbabilities,
	selectSpecificPerson: selectSpecificPerson,
}
