const fs = require('fs');
const path = require('path');
const test = require('node:test');
const assert = require('node:assert/strict');

const projectRoot = path.join(__dirname, '..');

test('Intro-Tester bleibt ein eigener Electron-Einstiegspunkt', () => {
	const testerPackage = JSON.parse(fs.readFileSync(path.join(projectRoot, 'scripts', 'intro-tester', 'package.json'), 'utf8'));
	const launcher = fs.readFileSync(path.join(projectRoot, 'intro-tester.cmd'), 'utf8');

	assert.equal(testerPackage.main, 'main.js');
	assert.match(launcher, /scripts\\intro-tester/);
	assert.doesNotMatch(launcher, /app\.js/);
});

test('alle Varianten verwenden denselben Intro-Player und sind neu startbar', () => {
	const html = fs.readFileSync(path.join(projectRoot, 'scripts', 'intro-tester', 'index.html'), 'utf8');
	const player = fs.readFileSync(path.join(projectRoot, 'scripts', 'intro-tester', 'intro-player.js'), 'utf8');

	assert.equal((html.match(/data-variant=/g) || []).length, 4);
	assert.match(html, /data-variant="2\.5">Variante 2 v2/);
	assert.match(html, /Nochmal abspielen/);
	assert.match(player, /function playerMarkup\(variant\)/);
	assert.match(player, /mount\.innerHTML = playerMarkup\(currentVariant\)/);
	assert.match(player, /playVariant\(button\.dataset\.variant\)/);
});

test('Tester übernimmt die produktiven Intro-Kernwerte', () => {
	const productionCss = fs.readFileSync(path.join(projectRoot, 'public', 'style.css'), 'utf8');
	const testerCss = fs.readFileSync(path.join(projectRoot, 'scripts', 'intro-tester', 'intro-tester.css'), 'utf8');
	const testerPlayer = fs.readFileSync(path.join(projectRoot, 'scripts', 'intro-tester', 'intro-player.js'), 'utf8');

	for (const sharedValue of ['stroke-width: .12rem', 'font-size: 5.5rem']) {
		assert.match(testerCss, new RegExp(sharedValue.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')));
	}
	assert.match(testerCss, /stroke-opacity: 0/);
	assert.match(testerCss, /stroke-dasharray: 420 420/);
	assert.match(testerCss, /\.line-drawing \.intro-word[\s\S]+animation: intro-line-draw 3\.5s linear forwards/);
	assert.match(testerCss, /\.variant-3 \.intro-word[\s\S]+intro-dash-array 4\.2s[^;]+infinite/);
	assert.match(testerCss, /\.intro-seven[\s\S]+font-weight: 900/);
	assert.match(productionCss, /#3b82f6|--accent: #3b82f6/);
	assert.match(testerPlayer, /stop-color="#3b82f6"/);
	assert.match(testerCss, /Roboto/);
	assert.match(testerPlayer, /beginElement\(\)/);
	assert.match(testerPlayer, /line-reveal-mask/);
	assert.match(testerPlayer, /class="line-reveal-motion"[^>]+dur="3\.5s"/);
	assert.match(testerPlayer, /variant === 1 \? \{x1: 680, x2: 840\}/);
	assert.match(testerPlayer, /currentVariant === 2\.5/);
	assert.match(testerPlayer, /function alignSevenInkCenter\(canvas\)/);
	assert.match(testerPlayer, /wordCenter - sevenCenter/);
});
