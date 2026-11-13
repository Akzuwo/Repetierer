const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs'),os=require('node:os'),path=require('node:path');
const {spawnSync}=require('node:child_process');
const {formatReleaseNotes}=require('../scripts/generate-release-notes');
const root=path.resolve(__dirname,'..');

function runFixture({tag,notes={ '7.0.0':{title:'Änderungen',changes:['Erste Änderung','Zweite Änderung']} },version='7.0.0'}={}){
 const directory=fs.mkdtempSync(path.join(os.tmpdir(),'repetierer release notes '));
 fs.mkdirSync(path.join(directory,'scripts'));
 fs.copyFileSync(path.join(root,'scripts/generate-release-notes.js'),path.join(directory,'scripts/generate-release-notes.js'));
 fs.writeFileSync(path.join(directory,'package.json'),JSON.stringify({version}));
 fs.writeFileSync(path.join(directory,'Buildnotes.json'),typeof notes==='string'?notes:JSON.stringify(notes));
 fs.writeFileSync(path.join(directory,'release-notes.txt'),'Dieser alte Text darf nicht verwendet werden.');
 const output=path.join(directory,'github-output.txt');
 const result=spawnSync(process.execPath,[path.join(directory,'scripts/generate-release-notes.js')],{
  cwd:directory,encoding:'utf8',windowsHide:true,
  env:{...process.env,GITHUB_REF_TYPE:tag?'tag':'branch',GITHUB_REF_NAME:tag||'master',GITHUB_OUTPUT:output,RELEASE_NOTES_OVERRIDE:'Dieser manuelle Text darf nicht verwendet werden.'}
 });
 return {...result,directory,output};
}

test('manual and tag workflows generate the JSON title and changes, preserving Unicode',()=>{
 for(const tag of [undefined,'v7.0.0']){
  const run=runFixture({tag});assert.equal(run.status,0,run.stderr);
  assert.equal(fs.readFileSync(path.join(run.directory,'release_notes.md'),'utf8'),'## 7.0.0 – Änderungen\n\n- Erste Änderung\n- Zweite Änderung\n');
  assert.equal(fs.readFileSync(run.output,'utf8'),'version=7.0.0\ntag=v7.0.0\nrelease_notes_path=release_notes.md\n');
  assert.match(run.stdout,/aus Buildnotes\.json/);
 }
});

test('missing/malformed version entries stop generation without falling back to older notes',()=>{
 for(const notes of [{},{'6.1.4':{title:'Alt',changes:['Alt']}},{'7.0.0':{title:'',changes:['Text']}},{'7.0.0':{title:'Titel',changes:[]}},{'7.0.0':{title:'Titel',changes:['Text',null]}},'{broken',[]]){
  const run=runFixture({notes});assert.equal(run.status,1);assert.match(run.stderr,/Buildnotes\.json/);
  assert.equal(fs.existsSync(path.join(run.directory,'release_notes.md')),false);
  assert.equal(fs.existsSync(run.output),false);
 }
});

test('a tag for another package version fails before generating release metadata',()=>{
 const run=runFixture({tag:'v8.0.0'});assert.equal(run.status,1);assert.match(run.stderr,/passt nicht zur Paketversion/);
 assert.equal(fs.existsSync(run.output),false);
});

test('multiline changes remain inside their Markdown list entry',()=>{
 assert.equal(formatReleaseNotes({'1.0.0':{title:'Titel',changes:['Zeile 1\r\nZeile 2']}},'1.0.0'),'## 1.0.0 – Titel\n\n- Zeile 1\n  Zeile 2\n');
});

test('the actual current version includes every change from the tracked Buildnotes.json',()=>{
 const version=require('../package.json').version,notes=JSON.parse(fs.readFileSync(path.join(root,'Buildnotes.json'),'utf8'));
 const body=formatReleaseNotes(notes,version);assert.ok(body.includes(notes[version].title));
 for(const change of notes[version].changes)assert.ok(body.includes('- '+change),change);
});
