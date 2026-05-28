const fs = require("fs");
const path = require("path");
const { readReleaseNotesForVersion } = require("../src/releaseNotes.js");

const rootDir = path.resolve(__dirname, "..");
const packageJsonPath = path.join(rootDir, "package.json");
const generatedReleaseNotesPath = path.join(rootDir, "release_notes.md");

function getVersionSource() {
  const refType = process.env.GITHUB_REF_TYPE;
  const refName = process.env.GITHUB_REF_NAME;

  if (refType === "tag" && refName) {
    return refName;
  }

  const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, "utf8"));
  if (!packageJson.version) {
    throw new Error("Keine Version in package.json gefunden.");
  }
  return packageJson.version;
}

function normalizeVersion(versionSource) {
  return versionSource.replace(/^v/, "");
}

function writeGithubOutput(values) {
  const outputPath = process.env.GITHUB_OUTPUT;
  if (!outputPath) return;

  const output = Object.entries(values)
    .map(([key, value]) => `${key}=${value}`)
    .join("\n");
  fs.appendFileSync(outputPath, `${output}\n`, "utf8");
}

function main() {
  const versionSource = getVersionSource();
  const version = normalizeVersion(versionSource);
  const tag = versionSource.startsWith("v") ? versionSource : `v${version}`;
  const releaseBody = readReleaseNotesForVersion(rootDir, version);

  if (!releaseBody) {
    throw new Error(`Keine Release Notes fuer Version ${version} in release-notes.txt gefunden.`);
  }

  fs.writeFileSync(generatedReleaseNotesPath, `${releaseBody}\n`, "utf8");
  writeGithubOutput({ version, tag, release_notes_path: "release_notes.md" });

  console.log(`Release Notes fuer ${version} aus release-notes.txt nach release_notes.md geschrieben.`);
}

main();
