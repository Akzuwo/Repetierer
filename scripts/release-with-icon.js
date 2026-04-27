const { spawnSync } = require("child_process");
const fs = require("fs");
const path = require("path");
const rcedit = require("rcedit");

const rootDir = path.resolve(__dirname, "..");
const outputDir = path.join(rootDir, "builds");
const unpackedDir = path.join(outputDir, "win-unpacked");
const resourcesDir = path.join(unpackedDir, "resources");
const exePath = path.join(unpackedDir, "Repetierer.exe");
const iconPath = path.join(rootDir, "public", "icons", "icon.ico");
const appUpdatePath = path.join(resourcesDir, "app-update.yml");

function run(command, args) {
  const result = spawnSync(command, args, {
    cwd: rootDir,
    env: process.env,
    stdio: "inherit",
    shell: false,
  });

  if (result.status !== 0) {
    const reason = result.error ? result.error.message : `exit code ${result.status}`;
    throw new Error(`${command} ${args.join(" ")} failed with ${reason}`);
  }
}

function requireFile(filePath, label) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`${label} fehlt: ${filePath}`);
  }
}

function writeAppUpdateConfig() {
  const packageJson = JSON.parse(fs.readFileSync(path.join(rootDir, "package.json"), "utf8"));
  const githubPublishConfig = (packageJson.build.publish || []).find(config => config.provider === "github");

  if (!githubPublishConfig) {
    throw new Error("Keine GitHub-Publish-Konfiguration in package.json gefunden.");
  }

  const yaml = [
    "provider: github",
    `owner: ${githubPublishConfig.owner}`,
    `repo: ${githubPublishConfig.repo}`,
    `releaseType: ${githubPublishConfig.releaseType || "release"}`,
    "updaterCacheDirName: repetierer-updater",
    "",
  ].join("\n");

  fs.mkdirSync(resourcesDir, { recursive: true });
  fs.writeFileSync(appUpdatePath, yaml, "utf8");
  console.log(`Schreibe Update-Konfiguration nach ${path.relative(rootDir, appUpdatePath)}`);
}

async function main() {
  const node = process.execPath;
  const electronBuilderCli = path.join(rootDir, "node_modules", "electron-builder", "out", "cli", "cli.js");
  const publishMode = process.env.PUBLISH_MODE || "always";

  requireFile(iconPath, "App Icon");
  requireFile(electronBuilderCli, "electron-builder CLI");

  run(node, [
    electronBuilderCli,
    "--win",
    "dir",
    "--config.win.signAndEditExecutable=false",
    "--publish",
    "never",
  ]);

  requireFile(exePath, "Entpackte Repetierer.exe");

  console.log(`Setze App Icon fuer ${path.relative(rootDir, exePath)}`);
  await rcedit(exePath, {
    icon: iconPath,
  });

  writeAppUpdateConfig();
  requireFile(appUpdatePath, "Update-Konfiguration");

  run(node, [
    electronBuilderCli,
    "--prepackaged",
    unpackedDir,
    "--win",
    "nsis",
    "--config.win.signAndEditExecutable=false",
    "--publish",
    publishMode,
  ]);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
