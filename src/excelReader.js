const ExcelJS = require('exceljs');
let f;
let wb;
let ws;

// initialize the workbook
function init(file, callback) {
	f = file;
	wb = new ExcelJS.Workbook();
	wb.xlsx.readFile(f)
		.then(() => {
			let worksheets = [];
			wb.worksheets.forEach(v => {
				worksheets.push(v.name);
			})
			callback(worksheets);
		})
		.catch(() => {
			callback();
		})
}

// read and extract persons
function read(clss) {
	ws = wb.getWorksheet(clss);

	const currentDate = new Date();
	const formattedDate = `${String(currentDate.getDate()).padStart(2, '0')}.${String(currentDate.getMonth() + 1).padStart(2, '0')}.${currentDate.getFullYear()}`;

	// check file
	if (ws.getCell('A1').value !== 'repetierer') return;

	// loop
	let persons = [];
	for (let i = 0; i < 25; i++) {
		let name = ws.getCell('A' + (i + 6));
		let joker_date = ws.getCell('Q' + (i + 6))

		if (!name.value) break;

		// check if a person used the joker on the current date
		if (joker_date.value == formattedDate) continue;

		let grades = countGrades(i)
        let joker = ws.getCell('H' + (i + 6));
        if (grades < 6)
			persons.push({
				id: i,
				name: name.value,
				grades: grades,
                joker: joker.value
			})
	}

	if (persons.length === 0) return;

	return persons;

	function countGrades(i) {
		let n = 0;
		[1,2,3,4,5,6].map(x => {
			if (ws.getCell(i + 6, x + 1).value) n++;
		});
		return n;
	}
}

function getFormattedDate() {
	const currentDate = new Date();
	return `${String(currentDate.getDate()).padStart(2, '0')}.${String(currentDate.getMonth() + 1).padStart(2, '0')}.${currentDate.getFullYear()}`;
}

function reloadAfterWrite(clss, callback, result) {
	init(f, (a) => {
		if (a) {
			callback(read(clss), result);
		} else {
			callback(undefined, Object.assign({}, result, { success: false, reason: 'reload-failed' }));
		}
	});
}

function writeWorkbook(clss, callback, result) {
	wb.xlsx.writeFile(f)
		.then(() => reloadAfterWrite(clss, callback, Object.assign({}, result, { success: true })))
		.catch(error => {
			const failedResult = Object.assign({}, result, {
				success: false,
				reason: 'excel-write-failed',
				error: error && error.message ? error.message : String(error)
			});

			init(f, () => callback(undefined, failedResult));
		});
}

// write grade to file
function write_grade(clss, person, grade, callback) {

	// check file
	if (ws.getCell('A1').value !== 'repetierer') {
		callback(undefined, { success: false, reason: 'invalid-file' });
		return;
	}

	ws.getCell(person.id + 6, person.grades + 2).value = grade;
	ws.getCell(person.id + 6, person.grades + 11).value = getFormattedDate();
	writeWorkbook(clss, callback, { type: 'grade' });
}
function write_joker(clss, person, callback, allowExtraJoker) {

	// check file
	if (ws.getCell('A1').value !== 'repetierer') {
		callback(undefined, { success: false, reason: 'invalid-file' });
		return;
	}

	const jokerCell = ws.getCell('H' + (person.id + 6));
	const currentJokerValue = Number(jokerCell.value || 0);
	if (currentJokerValue >= 1 && !allowExtraJoker) {
		callback(undefined, { success: false, reason: 'joker-already-used' });
		return;
	}
	if (currentJokerValue >= 2) {
		callback(undefined, { success: false, reason: 'joker-already-used' });
		return;
	}
	jokerCell.value = currentJokerValue === 1 && allowExtraJoker ? 2 : 1;

	ws.getCell('Q' + (person.id + 6)).value = getFormattedDate();
	writeWorkbook(clss, callback, { type: 'joker' });
}

function apply_entries(clss, entries, callback) {
	if (!entries || entries.length === 0) {
		callback(read(clss), { success: true, applied: [] });
		return;
	}

	const applied = [];
	entries.forEach(entry => {
		const entryWorksheet = wb.getWorksheet(entry.className || clss);
		if (!entryWorksheet || entryWorksheet.getCell('A1').value !== 'repetierer') return;

		const row = Number(entry.personId) + 6;
		if (!Number.isFinite(row)) return;

		if (entry.type === 'grade') {
			const gradeIndex = countGradesForRow(entryWorksheet, row);
			if (gradeIndex >= 6) return;
			entryWorksheet.getCell(row, gradeIndex + 2).value = entry.grade;
			entryWorksheet.getCell(row, gradeIndex + 11).value = entry.date || getFormattedDate();
			applied.push(entry.id);
		}

		if (entry.type === 'joker') {
			const jokerCell = entryWorksheet.getCell('H' + row);
			const currentJokerValue = Number(jokerCell.value || 0);
			if (currentJokerValue >= 2) return;
			jokerCell.value = currentJokerValue + 1;
			entryWorksheet.getCell('Q' + row).value = entry.date || getFormattedDate();
			applied.push(entry.id);
		}
	});

	if (applied.length === 0) {
		callback(read(clss), { success: true, applied: [] });
		return;
	}

	writeWorkbook(clss, callback, { type: 'pending-entries', applied: applied });

	function countGradesForRow(entryWorksheet, row) {
		let n = 0;
		[1,2,3,4,5,6].map(x => {
			if (entryWorksheet.getCell(row, x + 1).value) n++;
		});
		return n;
	}
}
module.exports = {
	init: init,
	read: read,
	write_grade: write_grade,
    write_joker: write_joker,
	apply_entries: apply_entries
}
