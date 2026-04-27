const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
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
        let joker = Number(ws.getCell('H' + (i + 6)).value || 0);
        if (grades < 6)
			persons.push({
				id: i,
				name: name.value,
				grades: grades,
                joker: joker
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

function read_editor_persons(clss) {
	const worksheet = wb.getWorksheet(clss);
	if (!worksheet || worksheet.getCell('A1').value !== 'repetierer') return;

	const persons = [];
	for (let row = 6; row <= 30; row++) {
		const name = worksheet.getCell('A' + row).value;
		if (!name) break;

		persons.push({
			id: row - 6,
			name: String(name),
			grades: countGradesForRow(worksheet, row),
			joker: Number(worksheet.getCell('H' + row).value || 0)
		});
	}

	return persons;
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

function getFileLockStatus(filePath) {
	const targetPath = filePath || f;
	const ownerLockPath = getExistingOwnerLockPath(targetPath);
	if (ownerLockPath) {
		return {
			locked: true,
			error: `Excel lock file exists: ${path.basename(ownerLockPath)}`,
			code: 'EXCEL_OWNER_LOCK'
		};
	}

	if (process.platform === 'win32') {
		const windowsLockStatus = getWindowsFileLockStatus(targetPath);
		if (windowsLockStatus) return windowsLockStatus;
	}

	try {
		const handle = fs.openSync(targetPath, 'r+');
		fs.closeSync(handle);
		return { locked: false };
	} catch (error) {
		return {
			locked: true,
			error: error && error.message ? error.message : String(error),
			code: error && error.code
		};
	}
}

function getExistingOwnerLockPath(targetPath) {
	const directory = path.dirname(targetPath);
	const basename = path.basename(targetPath);
	const candidates = [
		path.join(directory, `~$${basename}`),
		path.join(directory, `~$${basename.slice(2)}`)
	];

	return candidates.find(candidate => fs.existsSync(candidate));
}

function getWindowsFileLockStatus(targetPath) {
	const script = [
		'$targetPath = $env:REPETIERER_LOCK_CHECK_PATH',
		'$ErrorActionPreference = "Stop"',
		'$stream = $null',
		'try {',
		'  $stream = [System.IO.File]::Open($targetPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)',
		'  exit 0',
		'} catch [System.IO.IOException] {',
		'  [Console]::Error.WriteLine($_.Exception.Message)',
		'  exit 2',
		'} catch {',
		'  [Console]::Error.WriteLine($_.Exception.Message)',
		'  exit 3',
		'} finally {',
		'  if ($stream) { $stream.Close() }',
		'}'
	].join('; ');
	const encodedScript = Buffer.from(script, 'utf16le').toString('base64');

	const result = spawnSync('powershell.exe', [
		'-NoProfile',
		'-NonInteractive',
		'-ExecutionPolicy',
		'Bypass',
		'-EncodedCommand',
		encodedScript
	], {
		encoding: 'utf8',
		env: Object.assign({}, process.env, { REPETIERER_LOCK_CHECK_PATH: targetPath }),
		windowsHide: true
	});

	if (result.error) return null;
	if (result.status === 0) return { locked: false };

	return {
		locked: true,
		error: (result.stderr || result.stdout || 'File is locked').trim(),
		code: result.status === 2 ? 'EXCEL_EXCLUSIVE_LOCK' : 'EXCEL_LOCK_CHECK_FAILED'
	};
}

function writeWorkbook(clss, callback, result) {
	const lockStatus = getFileLockStatus(f);
	if (lockStatus.locked) {
		callback(undefined, Object.assign({}, result, {
			success: false,
			reason: 'excel-locked',
			error: lockStatus.error
		}));
		return;
	}

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

function getJokerMigrationStatus() {
	const status = {
		needsMigration: false,
		emptyCells: 0,
		usedJokerCells: 0,
		sheets: []
	};

	if (!wb) return status;

	wb.worksheets.forEach(worksheet => {
		if (!worksheet || worksheet.getCell('A1').value !== 'repetierer') return;
		const sheetStatus = {
			name: worksheet.name,
			emptyCells: 0,
			usedJokerCells: 0
		};

		for (let row = 6; row <= 30; row++) {
			if (!worksheet.getCell('A' + row).value) break;
			const value = worksheet.getCell('H' + row).value;
			if (value === null || value === undefined || value === '') {
				sheetStatus.emptyCells++;
			} else if (Number(value) === 1) {
				sheetStatus.usedJokerCells++;
			}
		}

		if (sheetStatus.emptyCells > 0 || sheetStatus.usedJokerCells > 0) {
			status.sheets.push(sheetStatus);
			status.emptyCells += sheetStatus.emptyCells;
			status.usedJokerCells += sheetStatus.usedJokerCells;
		}
	});

	status.needsMigration = status.emptyCells > 0;
	return status;
}

function migrate_jokers(callback) {
	const lockStatus = getFileLockStatus(f);
	if (lockStatus.locked) {
		callback({ success: false, reason: 'excel-locked', error: lockStatus.error });
		return;
	}

	let migratedCells = 0;
	wb.worksheets.forEach(worksheet => {
		if (!worksheet || worksheet.getCell('A1').value !== 'repetierer') return;

		for (let row = 6; row <= 30; row++) {
			if (!worksheet.getCell('A' + row).value) break;
			const jokerCell = worksheet.getCell('H' + row);
			const value = jokerCell.value;

			if (value === null || value === undefined || value === '') {
				jokerCell.value = 1;
				migratedCells++;
			} else if (Number(value) === 1) {
				jokerCell.value = 0;
				migratedCells++;
			}
		}
	});

	wb.xlsx.writeFile(f)
		.then(() => init(f, worksheets => callback({ success: !!worksheets, worksheets: worksheets, migratedCells: migratedCells })))
		.catch(error => {
			callback({
				success: false,
				reason: 'excel-write-failed',
				error: error && error.message ? error.message : String(error)
			});
		});
}

// write grade to file
function write_grade(clss, person, grade, callback, awardExtraJoker) {

	// check file
	if (ws.getCell('A1').value !== 'repetierer') {
		callback(undefined, { success: false, reason: 'invalid-file' });
		return;
	}

	const lockStatus = getFileLockStatus(f);
	if (lockStatus.locked) {
		callback(undefined, { success: false, reason: 'excel-locked', error: lockStatus.error });
		return;
	}

	ws.getCell(person.id + 6, person.grades + 2).value = grade;
	ws.getCell(person.id + 6, person.grades + 11).value = getFormattedDate();
	if (awardExtraJoker) {
		const jokerCell = ws.getCell('H' + (person.id + 6));
		jokerCell.value = Number(jokerCell.value || 0) + 1;
	}
	writeWorkbook(clss, callback, { type: 'grade' });
}
function write_joker(clss, person, callback) {

	// check file
	if (ws.getCell('A1').value !== 'repetierer') {
		callback(undefined, { success: false, reason: 'invalid-file' });
		return;
	}

	const lockStatus = getFileLockStatus(f);
	if (lockStatus.locked) {
		callback(undefined, { success: false, reason: 'excel-locked', error: lockStatus.error });
		return;
	}

	const jokerCell = ws.getCell('H' + (person.id + 6));
	const remainingJokers = Number(jokerCell.value || 0);
	if (remainingJokers <= 0) {
		callback(undefined, { success: false, reason: 'joker-already-used' });
		return;
	}
	jokerCell.value = remainingJokers - 1;

	ws.getCell('Q' + (person.id + 6)).value = getFormattedDate();
	writeWorkbook(clss, callback, { type: 'joker' });
}

function edit_persons(clss, editorPersons, callback) {
	const worksheet = wb.getWorksheet(clss);
	if (!worksheet || worksheet.getCell('A1').value !== 'repetierer') {
		callback(undefined, { success: false, reason: 'invalid-file' });
		return;
	}

	const lockStatus = getFileLockStatus(f);
	if (lockStatus.locked) {
		callback(undefined, { success: false, reason: 'excel-locked', error: lockStatus.error });
		return;
	}

	const existingRows = new Map();
	for (let row = 6; row <= 30; row++) {
		if (!worksheet.getCell('A' + row).value) break;
		existingRows.set(row - 6, copyRowValues(worksheet, row));
	}

	const normalizedPersons = (editorPersons || [])
		.map(person => ({
			id: person.id === null || person.id === undefined ? null : Number(person.id),
			name: String(person.name || '').trim(),
			joker: Number(person.joker)
		}));

	const hasInvalidPerson = normalizedPersons.some(person =>
		!person.name ||
		!Number.isInteger(person.joker) ||
		person.joker < 0
	);
	if (hasInvalidPerson) {
		callback(undefined, { success: false, reason: 'invalid-editor-data' });
		return;
	}

	if (normalizedPersons.length > 25) {
		callback(undefined, { success: false, reason: 'too-many-persons' });
		return;
	}

	for (let row = 6; row <= 30; row++) {
		clearRosterRow(worksheet, row);
	}

	normalizedPersons.forEach((person, index) => {
		const targetRow = index + 6;
		const sourceValues = Number.isFinite(person.id) && existingRows.has(person.id) ? existingRows.get(person.id) : {};
		writeRowValues(worksheet, targetRow, sourceValues);
		worksheet.getCell('A' + targetRow).value = person.name;
		worksheet.getCell('H' + targetRow).value = person.joker;
	});

	writeWorkbook(clss, callback, { type: 'edit-persons' });
}

function apply_entries(clss, entries, callback) {
	if (!entries || entries.length === 0) {
		callback(read(clss), { success: true, applied: [] });
		return;
	}

	const lockStatus = getFileLockStatus(f);
	if (lockStatus.locked) {
		callback(undefined, {
			success: false,
			reason: 'excel-locked',
			error: lockStatus.error,
			applied: []
		});
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
			if (entry.awardExtraJoker && gradeIndex < 3 && gradeIndex + 1 >= 3) {
				const jokerCell = entryWorksheet.getCell('H' + row);
				jokerCell.value = Number(jokerCell.value || 0) + 1;
			}
			applied.push(entry.id);
		}

		if (entry.type === 'joker') {
			const jokerCell = entryWorksheet.getCell('H' + row);
			const remainingJokers = Number(jokerCell.value || 0);
			if (remainingJokers <= 0) return;
			jokerCell.value = remainingJokers - 1;
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

function countGradesForRow(worksheet, row) {
	let n = 0;
	[1,2,3,4,5,6].map(x => {
		if (worksheet.getCell(row, x + 1).value) n++;
	});
	return n;
}

function copyRowValues(worksheet, row) {
	const values = {};
	for (let column = 1; column <= 17; column++) {
		values[column] = worksheet.getCell(row, column).value;
	}
	return values;
}

function writeRowValues(worksheet, row, values) {
	for (let column = 1; column <= 17; column++) {
		worksheet.getCell(row, column).value = Object.prototype.hasOwnProperty.call(values, column) ? values[column] : null;
	}
}

function clearRosterRow(worksheet, row) {
	for (let column = 1; column <= 17; column++) {
		worksheet.getCell(row, column).value = null;
	}
}
module.exports = {
	init: init,
	read: read,
	read_editor_persons: read_editor_persons,
	write_grade: write_grade,
    write_joker: write_joker,
	edit_persons: edit_persons,
	apply_entries: apply_entries,
	getJokerMigrationStatus: getJokerMigrationStatus,
	migrate_jokers: migrate_jokers,
	getFileLockStatus: getFileLockStatus
}
