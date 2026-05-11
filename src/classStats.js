function buildClassStatistics(workbook) {
	if (!workbook || !Array.isArray(workbook.worksheets)) return [];

	return workbook.worksheets
		.map(worksheet => buildWorksheetStatistics(worksheet))
		.filter(Boolean);
}

function buildWorksheetStatistics(worksheet) {
	if (!worksheet || getCellValue(worksheet.getCell('A1')) !== 'repetierer') return null;

	const persons = readPersons(worksheet);
	if (persons.length === 0) {
		return {
			className: worksheet.name || 'Klasse',
			personCount: 0,
			repetitionsTotal: 0,
			averageGrade: null,
			topRepetitionPerson: null,
			topRepetitionAdditionalCount: 0,
			averageRepetitionsPerPerson: null
		};
	}

	const repetitionsTotal = persons.reduce((sum, person) => sum + person.repetitions, 0);
	const validGrades = persons.flatMap(person => person.validGrades);
	const topRepetitions = Math.max(...persons.map(person => person.repetitions));
	const topPersons = persons.filter(person => person.repetitions === topRepetitions);

	return {
		className: worksheet.name || 'Klasse',
		personCount: persons.length,
		repetitionsTotal: repetitionsTotal,
		averageGrade: validGrades.length > 0 ? roundToTwo(validGrades.reduce((sum, grade) => sum + grade, 0) / validGrades.length) : null,
		topRepetitionPerson: topPersons.length > 0 ? {
			name: topPersons[0].name,
			repetitions: topPersons[0].repetitions
		} : null,
		topRepetitionAdditionalCount: Math.max(0, topPersons.length - 1),
		averageRepetitionsPerPerson: roundToTwo(repetitionsTotal / persons.length)
	};
}

function readPersons(worksheet) {
	const persons = [];

	for (let row = 6; row <= 30; row++) {
		const name = getCellValue(worksheet.getCell('A' + row));
		if (!name) break;

		const gradeCells = [];
		for (let column = 2; column <= 7; column++) {
			gradeCells.push(worksheet.getCell(row, column));
		}

		persons.push({
			name: String(name),
			repetitions: gradeCells.filter(cell => hasCellValue(cell)).length,
			validGrades: gradeCells
				.map(cell => parseGrade(getCellValue(cell)))
				.filter(grade => grade !== null)
		});
	}

	return persons;
}

function hasCellValue(cell) {
	const value = getCellValue(cell);
	return value !== null && value !== undefined && String(value).trim() !== '';
}

function getCellValue(cell) {
	if (!cell) return null;
	const value = cell.value;
	if (value && typeof value === 'object') {
		if (Object.prototype.hasOwnProperty.call(value, 'result')) return value.result;
		if (Array.isArray(value.richText)) return value.richText.map(part => part.text || '').join('');
		if (Object.prototype.hasOwnProperty.call(value, 'text')) return value.text;
	}
	return value;
}

function parseGrade(value) {
	if (value === null || value === undefined) return null;
	const normalized = typeof value === 'number'
		? value
		: Number(String(value).trim().replace(',', '.'));
	if (!Number.isFinite(normalized) || normalized < 1 || normalized > 6) return null;
	return normalized;
}

function roundToTwo(value) {
	return Math.round((value + Number.EPSILON) * 100) / 100;
}

module.exports = {
	buildClassStatistics: buildClassStatistics
};
