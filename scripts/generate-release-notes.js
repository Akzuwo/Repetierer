const fs = require('node:fs');
const path = require('node:path');

function formatReleaseNotes(notes, version) {
  const entry = Object.hasOwn(notes, version) ? notes[version] : null;
  if (!entry || typeof entry.title !== 'string' || !entry.title.trim() ||
      !Array.isArray(entry.changes) || !entry.changes.length ||
      entry.changes.some(change => typeof change !== 'string' || !change.trim())) {
    throw new Error(`Buildnotes.json: Version ${version} benötigt einen Titel und eine nicht leere Liste von Änderungen.`);
  }
  const changes = entry.changes.map(change => `- ${change.trim().replace(/\r\n/g, '\n').replace(/\n/g, '\n  ')}`);
  return `## ${version} – ${entry.title.trim()}\n\n${changes.join('\n')}\n`;
}

function main(rootDir = path.resolve(__dirname, '..'), env = process.env) {
  const metadata = JSON.parse(fs.readFileSync(path.join(rootDir, 'package.json'), 'utf8'));
  if (typeof metadata.version !== 'string' || !metadata.version.trim()) throw new Error('Keine Version in package.json gefunden.');
  const versionSource = env.GITHUB_REF_TYPE === 'tag' && env.GITHUB_REF_NAME ? env.GITHUB_REF_NAME : metadata.version;
  const version = versionSource.replace(/^v/, '');
  if (version !== metadata.version) throw new Error(`Release-Tag ${versionSource} passt nicht zur Paketversion ${metadata.version}.`);
  let notes;
  try { notes = JSON.parse(fs.readFileSync(path.join(rootDir, 'Buildnotes.json'), 'utf8')); }
  catch (error) { throw new Error(`Buildnotes.json konnte nicht gelesen werden: ${error.message}`); }
  if (!notes || typeof notes !== 'object' || Array.isArray(notes)) throw new Error('Buildnotes.json muss ein Objekt mit Versionseinträgen enthalten.');
  const releaseBody = formatReleaseNotes(notes, version);
  const tag = `v${version}`;
  fs.writeFileSync(path.join(rootDir, 'release_notes.md'), releaseBody, 'utf8');
  if (env.GITHUB_OUTPUT) fs.appendFileSync(env.GITHUB_OUTPUT, `version=${version}\ntag=${tag}\nrelease_notes_path=release_notes.md\n`, 'utf8');
  console.log(`Release Notes für ${version} aus Buildnotes.json nach release_notes.md geschrieben.`);
  return {version, tag, releaseBody};
}

module.exports = {formatReleaseNotes, main};
if (require.main === module) {
  try { main(); } catch (error) { console.error(error.message); process.exitCode = 1; }
}
