const fs = require('fs');
const path = require('path');

function getReleaseNotesPath(rootDir) {
	return path.join(rootDir, 'release-notes.txt');
}

function extractReleaseNotesSection(content, version) {
	if (!content || !version) return '';

	const lines = String(content).replace(/\r\n/g, '\n').split('\n');
	const headerPattern = /^##\s+(.+?)\s*$/;
	let inSection = false;
	const sectionLines = [];

	for (const line of lines) {
		const headerMatch = line.match(headerPattern);
		if (headerMatch) {
			if (inSection) break;
			inSection = headerMatch[1].trim() === version;
			if (inSection) sectionLines.push(line);
			continue;
		}

		if (inSection) sectionLines.push(line);
	}

	return sectionLines.join('\n').trim();
}

function readReleaseNotesForVersion(rootDir, version) {
	const releaseNotesPath = getReleaseNotesPath(rootDir);
	if (!fs.existsSync(releaseNotesPath)) return '';
	return extractReleaseNotesSection(fs.readFileSync(releaseNotesPath, 'utf8'), version);
}

function stripReleaseNotesHeader(section) {
	return String(section || '').replace(/^##\s+.+?(?:\n|$)/, '').trim();
}

module.exports = {
	extractReleaseNotesSection,
	getReleaseNotesPath,
	readReleaseNotesForVersion,
	stripReleaseNotesHeader
};
