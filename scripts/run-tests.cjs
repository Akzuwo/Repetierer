const fs=require('node:fs'),path=require('node:path'),{spawnSync}=require('node:child_process');
const root=path.resolve(__dirname,'..');
// Private backend tests can remain on the operator's disk, but are not part
// of the public desktop suite. This also works from a source ZIP without Git.
const files=fs.readdirSync(path.join(root,'test')).filter(name=>name.endsWith('.test.js') && !name.startsWith('cloudflare')).sort().map(name=>path.join('test',name));
if(!files.length)throw new Error('No app tests found');
const result=spawnSync(process.execPath,['--test','--test-concurrency=1',...process.argv.slice(2),...files],{cwd:root,stdio:'inherit',windowsHide:true});
process.exit(result.status ?? 1);
