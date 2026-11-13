const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {publishRelease} = require('../scripts/publish-release');

function fixture(t, {missingTag = false, existingRelease = false, fail = ''} = {}) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'repetierer publish '));
  t.after(() => fs.rmSync(root, {recursive: true, force: true}));
  fs.mkdirSync(path.join(root, 'builds'));
  fs.writeFileSync(path.join(root, 'package.json'), JSON.stringify({version: '7.0.0', build: {
    publish: [{provider: 'github', owner: 'Akzuwo', repo: 'Repetierer'}]
  }}));
  for (const name of ['Repetierer-7.0.0-Build-28.13112026.exe', 'Repetierer-7.0.0-Build-28.13112026.exe.blockmap', 'latest.yml']) {
    fs.writeFileSync(path.join(root, 'builds', name), 'fixture');
  }
  fs.writeFileSync(path.join(root, 'release_notes.md'), '## Änderungen\n\n- Aus Buildnotes.json');
  const env = {GITHUB_REPOSITORY: 'Akzuwo/Repetierer', RELEASE_TAG: 'v7.0.0', GITHUB_SHA: 'a'.repeat(40)};
  const calls = [];
  const spawn = (command, args, options) => {
    assert.equal(command, 'gh'); assert.equal(options.shell, false); assert.equal(options.cwd, root);
    calls.push(args);
    if (fail && args.includes(fail)) return {status: 1, stderr: 'HTTP 403: forbidden'};
    if (args[1].includes('/git/ref/tags/') && missingTag) {
      missingTag = false; return {status: 1, stderr: 'gh: Not Found (HTTP 404)'};
    }
    if (args[1].includes('/releases/tags/') && !existingRelease) return {status: 1, stderr: 'gh: Not Found (HTTP 404)'};
    return {status: 0, stdout: '{}'};
  };
  return {root, env, spawn, calls, run: () => publishRelease(root, env, spawn)};
}

test('missing tag is created at the workflow SHA before a draft is created, uploaded and published', t => {
  const f = fixture(t, {missingTag: true}); f.run();
  assert.deepEqual(f.calls.map(args => args.slice(0, 2)), [
    ['api', 'repos/Akzuwo/Repetierer/git/ref/tags/v7.0.0'],
    ['api', 'repos/Akzuwo/Repetierer/git/refs'],
    ['api', 'repos/Akzuwo/Repetierer/git/ref/tags/v7.0.0'],
    ['api', 'repos/Akzuwo/Repetierer/releases/tags/v7.0.0'],
    ['release', 'create'], ['release', 'upload'], ['release', 'edit']
  ]);
  assert.ok(f.calls[1].includes(`sha=${f.env.GITHUB_SHA}`));
  assert.ok(f.calls[4].includes('--verify-tag')); assert.ok(f.calls[4].includes('--draft'));
  assert.ok(f.calls[5].includes(path.join(f.root, 'builds', 'latest.yml')));
  assert.ok(f.calls[6].includes('--draft=false')); assert.ok(f.calls[6].includes('--notes-file'));
});

test('reruns reuse existing tags and releases without deleting earlier build assets', t => {
  const f = fixture(t, {existingRelease: true}); f.run();
  assert.deepEqual(f.calls.map(args => args[1]), [
    'repos/Akzuwo/Repetierer/git/ref/tags/v7.0.0', 'repos/Akzuwo/Repetierer/releases/tags/v7.0.0', 'upload', 'edit'
  ]);
  assert.ok(f.calls[2].includes('--clobber'));
});

test('permission errors do not trigger tag or release creation', t => {
  const f = fixture(t, {fail: 'repos/Akzuwo/Repetierer/git/ref/tags/v7.0.0'});
  assert.throws(f.run, /HTTP 403/); assert.equal(f.calls.length, 1);
});

test('upload failures do not publish a draft or overwrite its release notes', t => {
  const f = fixture(t, {fail: 'upload'});
  assert.throws(f.run, /HTTP 403/); assert.equal(f.calls.at(-1)[1], 'upload');
});

test('missing update metadata stops before any GitHub mutation', t => {
  const f = fixture(t); fs.unlinkSync(path.join(f.root, 'builds', 'latest.yml'));
  assert.throws(f.run, /ENOENT/); assert.equal(f.calls.length, 0);
});

test('wrong repository, tag or missing source SHA stops before GitHub access', t => {
  for (const [key, value] of [['GITHUB_REPOSITORY', 'other/repo'], ['RELEASE_TAG', 'v8.0.0'], ['GITHUB_SHA', '']]) {
    const f = fixture(t); f.env[key] = value;
    assert.throws(f.run, /passen nicht/); assert.equal(f.calls.length, 0);
  }
});

test('workflow builds without a publisher and serializes publication after notes generation', () => {
  const yaml = require('js-yaml');
  const workflow = yaml.load(fs.readFileSync(path.join(__dirname, '../.github/workflows/release.yml'), 'utf8'));
  const steps = workflow.jobs['release-windows'].steps;
  const build = steps.findIndex(step => step.run === 'npm run release:ci');
  assert.ok(steps.findIndex(step => step.run === 'node scripts/generate-release-notes.js') < build);
  assert.equal(steps[build].env.PUBLISH_MODE, 'never');
  assert.equal(steps[build + 1].run, 'node scripts/publish-release.js');
  assert.equal(steps[build + 1].env.RELEASE_TAG, '${{ steps.release-notes.outputs.tag }}');
  assert.equal(workflow.concurrency['cancel-in-progress'], false);
});
