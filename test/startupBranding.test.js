const fs = require('fs');
const path = require('path');
const test = require('node:test');
const assert = require('node:assert/strict');

const projectRoot = path.join(__dirname, '..');

test('produktives Intro verwendet die Ein-Text-Animation aus Variante 1', () => {
	const html = fs.readFileSync(path.join(projectRoot, 'public', 'splash.html'), 'utf8');
	const script = fs.readFileSync(path.join(projectRoot, 'public', 'splash.js'), 'utf8');
	const css = fs.readFileSync(path.join(projectRoot, 'public', 'style.css'), 'utf8');

	assert.equal((html.match(/>Repetierer<\/text>/g) || []).length, 1);
	assert.match(html, /class="startup-loader-word"[\s\S]+mask="url\(#startup-line-reveal-mask\)"/);
	assert.match(html, /class="startup-fill-motion"[^>]+begin="indefinite"[^>]+dur="1\.2s"/);
	assert.match(html, /class="startup-line-reveal-motion"[^>]+begin="indefinite"[^>]+dur="3\.5s"/);
	assert.match(script, /lineStart: 140/);
	assert.match(script, /lineComplete: 3640/);
	assert.match(script, /fillComplete: 4840/);
	assert.match(script, /logoComplete: 6390/);
	assert.match(css, /\.startup-line-drawing \.startup-loader-word[\s\S]+startup-line-draw 3\.5s linear forwards/);
	assert.match(css, /\.startup-word-filled \.startup-loader-word[\s\S]+translateX\(-40px\)/);
	assert.match(css, /\.startup-seven-visible \.startup-loader-seven[\s\S]+startup-seven-arrive/);
	assert.equal((css.match(/format\("truetype"\)/g) || []).length, 6);
	assert.doesNotMatch(css, /format\("ttf"\)/);
});

test('Windows-App und Update-Shortcuts verwenden das neue Icon stabil', () => {
	const packageJson = JSON.parse(fs.readFileSync(path.join(projectRoot, 'package.json'), 'utf8'));
	const main = fs.readFileSync(path.join(projectRoot, 'app.js'), 'utf8');

	assert.equal(packageJson.build.win.icon, './public/icons/icon.ico');
	assert.equal(packageJson.build.nsis.installerIcon, './public/icons/icon.ico');
	assert.equal(packageJson.build.nsis.uninstallerIcon, './public/icons/icon.ico');
	assert.equal(packageJson.build.nsis.createDesktopShortcut, 'always');
	assert.match(main, /app\.setAppUserModelId\('net\.srpnt3\.repetierer'\)/);
	assert.equal((main.match(/public', 'icons', 'icon\.ico'/g) || []).length, 2);
});
