// Check the public Git index, not ignored private operator files. Findings
// contain only a file/line/rule; credential values are never printed.
const fs=require('node:fs'),path=require('node:path'),{execFileSync,spawnSync}=require('node:child_process');
const root=path.resolve(__dirname,'..');
const files=execFileSync('git',['ls-files','--cached','-z'],{cwd:root,encoding:'utf8'}).split('\0').filter(Boolean);
const privatePath=/(^|\/)(?:cloudflare|issue-handler|node_modules|builds|\.wrangler|\.cloudflare-private|\.cloudflare-test|\.ui-polish-smoke|\.smoke-user-data|\.kristall-check)(\/|$)|^scripts\/cloudflare-|^test\/cloudflareE2E\.test\.js$|^docs\/cloudflare-deployment\.md$|(^|\/)\.env(?:\.|$)|\.(?:env|repetierer|pem|key|pfx|p12|sqlite|db)$/i;
const rules=[
 ['private-key',/-----BEGIN (?:RSA |EC |DSA |OPENSSH )?PRIVATE KEY-----/g],
 ['github-token',/\b(?:gh[pousr]_[A-Za-z0-9_]{20,}|github_pat_[A-Za-z0-9_]{40,})\b/g],
 ['assessment-token',/\brt_[A-Za-z0-9_-]{43}\b/g],
 ['smtp-token',/\bxsmtpsib-[A-Za-z0-9_-]{20,}\b/g],
 ['aws-access-key',/\b(?:AKIA|ASIA)[A-Z0-9]{16}\b/g],
 ['credential-url',/https?:\/\/[^\s/@:]+:[^\s/@]+@/g]
];
const findings=[];
for(const file of files){
 if(privatePath.test(file))findings.push({file,rule:'private-file-tracked'});
 if(/(^|\/)(?:settings|backup|migrations|unterrichtssessions|mitmachnoten|(?:analytics|report)-config\.local)\.json$/i.test(file) || /\.(?:sqlite|db)(?:-.*)?$/i.test(file) || (/\.(?:xlsx|xls|csv)$/i.test(file) && file!=='Excel Template.xlsx'))findings.push({file,rule:'runtime-or-personal-data-tracked'});
 const buffer=fs.readFileSync(path.join(root,file));if(buffer.includes(0))continue;
 const text=buffer.toString('utf8');
 for(const [rule,pattern] of rules)for(const match of text.matchAll(pattern))findings.push({file,line:text.slice(0,match.index).split('\n').length,rule});
}
for(const file of ['.env.production','operator.env','credentials.repetierer','settings.json','backup.json','unterrichtssessions.json','cloudflare/wrangler.jsonc','scripts/cloudflare-teacher.js','test/cloudflareE2E.test.js','builds/old-build.exe']){
 if(spawnSync('git',['check-ignore','--no-index','--quiet','--',file],{cwd:root}).status!==0)findings.push({file,rule:'private-path-not-ignored'});
}
if(findings.length){console.error(JSON.stringify(findings,null,2));process.exitCode=1;}
else console.log(`REPOSITORY_CHECK_PASS: ${files.length} indexed files, private backend/credentials ignored, no high-confidence credential literals`);
