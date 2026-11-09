const fs = require('node:fs');
const path = require('node:path');
const { buildSync } = require('esbuild');
const root = path.join(__dirname, '..');
const output = path.join(root, 'public', 'vendor');
fs.mkdirSync(output, { recursive: true });
buildSync({ entryPoints: [path.join(__dirname, 'ui-vendor-entry.mjs')], outfile: path.join(output, 'ui-motion.js'), bundle: true, format: 'cjs', platform: 'browser', target: 'chrome140', minify: true, legalComments: 'eof', banner: { js: '// Local bundle: thinking-orbs 0.3.2 (MIT) / GSAP 3.15.0 (Standard License). Licenses alongside this file.' } });
fs.copyFileSync(path.join(root, 'node_modules/thinking-orbs/LICENSE'), path.join(output, 'thinking-orbs.LICENSE.txt'));
console.log('UI motion bundle built locally.');
