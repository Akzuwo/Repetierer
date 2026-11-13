const fs = require('node:fs');
const path = require('node:path');
const {spawnSync} = require('node:child_process');

function publishRelease(root = path.resolve(__dirname, '..'), env = process.env, spawn = spawnSync) {
  const metadata = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  const publisher = metadata.build.publish.find(entry => entry.provider === 'github');
  const repository = `${publisher.owner}/${publisher.repo}`;
  const tag = `v${metadata.version}`;
  if (env.GITHUB_REPOSITORY !== repository || env.RELEASE_TAG !== tag ||
      !/^[a-f0-9]{40}$/.test(env.GITHUB_SHA || '')) {
    throw new Error('Repository, Release-Tag oder Workflow-Commit passen nicht zur Veröffentlichung.');
  }
  const notes = path.join(root, 'release_notes.md');
  const builds = path.join(root, 'builds');
  const installers = fs.readdirSync(builds).filter(name =>
    name.startsWith(`Repetierer-${metadata.version}-Build-`) && name.endsWith('.exe'));
  if (installers.length !== 1) throw new Error('Genau ein Installer für die aktuelle Version wird erwartet.');
  const assets = [installers[0], `${installers[0]}.blockmap`, 'latest.yml'].map(name => path.join(builds, name));
  for (const file of [notes, ...assets]) {
    if (!fs.statSync(file).isFile() || fs.statSync(file).size === 0) throw new Error(`Release-Datei fehlt oder ist leer: ${file}`);
  }

  function gh(args, allowMissing = false) {
    const result = spawn('gh', args, {cwd: root, env, encoding: 'utf8', windowsHide: true, shell: false});
    if (result.error) throw result.error;
    if (result.status !== 0) {
      if (allowMissing && /\bHTTP 404\b/.test(result.stderr || '')) return null;
      throw new Error(`GitHub-Schritt ${args.slice(0, 2).join(' ')} fehlgeschlagen: ${result.stderr || result.stdout}`);
    }
    return result.stdout;
  }

  // Materialize the Git ref explicitly before creating a release. Never move an existing tag.
  const tagEndpoint = `repos/${repository}/git/ref/tags/${encodeURIComponent(tag)}`;
  if (gh(['api', tagEndpoint], true) === null) {
    gh(['api', `repos/${repository}/git/refs`, '--method', 'POST',
      '-f', `ref=refs/tags/${tag}`, '-f', `sha=${env.GITHUB_SHA}`]);
    gh(['api', tagEndpoint]);
  }
  const release = gh(['api', `repos/${repository}/releases/tags/${encodeURIComponent(tag)}`], true);
  if (release === null) {
    gh(['release', 'create', tag, '--repo', repository, '--verify-tag', '--target', env.GITHUB_SHA,
      '--draft', '--title', metadata.version, '--notes-file', notes]);
  }
  // A failed upload leaves a new release as a draft; reruns can replace the same assets.
  gh(['release', 'upload', tag, ...assets, '--repo', repository, '--clobber']);
  gh(['release', 'edit', tag, '--repo', repository, '--notes-file', notes, '--draft=false']);
  console.log(`Release ${tag} mit Installer, Blockmap, latest.yml und Buildnotes.json veröffentlicht.`);
}

module.exports = {publishRelease};
if (require.main === module) {
  try { publishRelease(); } catch (error) { console.error(error.message); process.exitCode = 1; }
}
