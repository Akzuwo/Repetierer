const fs = require("fs");
const path = require("path");

const rootDir = path.resolve(__dirname, "..");
const buildnotesPath = path.join(rootDir, "Buildnotes.json");
const packageJsonPath = path.join(rootDir, "package.json");
const releaseNotesPath = path.join(rootDir, "release_notes.md");

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

function readBuildnotes() {
  if (!fs.existsSync(buildnotesPath)) {
    return {};
  }

  return JSON.parse(fs.readFileSync(buildnotesPath, "utf8"));
}

function createReleaseBody(version, note) {
  const lines = [`# Repetierer v${version}`, ""];

  if (!note) {
    lines.push("Keine Buildnotes für diese Version gefunden.");
    return `${lines.join("\n")}\n`;
  }

  if (note.title) {
    lines.push(`## ${note.title}`, "");
  }

  if (Array.isArray(note.changes) && note.changes.length > 0) {
    for (const change of note.changes) {
      lines.push(`- ${change}`);
    }
  } else {
    lines.push("Keine Buildnotes für diese Version gefunden.");
  }

  return `${lines.join("\n")}\n`;
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
  const buildnotes = readBuildnotes();
  const releaseBody = createReleaseBody(version, buildnotes[version]);

  fs.writeFileSync(releaseNotesPath, releaseBody, "utf8");
  writeGithubOutput({ version, tag, release_notes_path: "release_notes.md" });

  console.log(`Release Notes fuer ${version} nach release_notes.md geschrieben.`);
}

main();
