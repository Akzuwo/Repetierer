const { init, read, write_grade, write_joker } = require('./excelReader.js');
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
		weight: e.weight,
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
		if (settings.boostNeverSelected && e.grades === 0) {
			weight *= settings.neverSelectedBoostFactor;
		}

		return {
			person: e,
			weight: weight
		};
	});
}

function getPersonResult(selectedPerson) {
	const settings = getAppSettings();
	const jokerCount = Number(selectedPerson.joker || 0);
	let jokerUsed = jokerCount >= 1;

	if (settings.extraJokerAfterThreeGrades && selectedPerson.grades >= 3 && jokerCount < 2) {
		jokerUsed = false;
	}

	return [selectedPerson.name, jokerUsed];
}

// save the new grade to the file
function saveGrade(grade, callback) {
	write_grade(cls, person, grade, (p) => {
		persons = p;
		callback(undefined, {
			type: 'grade',
			className: cls,
			filePath: filePath,
			personId: person.id,
			personName: person.name,
			grade: grade,
			excelWriteSucceeded: !!p
		})
	});
}

// save the joker to the file
function setJoker(callback) {
	const settings = getAppSettings();
	const allowExtraJoker = settings.extraJokerAfterThreeGrades && person.grades >= 3;

    write_joker(cls, person, (p) => {
        persons  = p;
		if (callback) {
			callback(undefined, {
				type: 'joker',
				className: cls,
				filePath: filePath,
				personId: person.id,
				personName: person.name,
				excelWriteSucceeded: !!p
			});
		}
    }, allowExtraJoker);

}

module.exports = {
	setFile: setFile,
	setClass: setClass,
	selectPerson: selectPerson,
	saveGrade: saveGrade,
	setJoker: setJoker,
	getPersons: getPersons,
	getProbabilities: getProbabilities,
	selectSpecificPerson: selectSpecificPerson,
}
