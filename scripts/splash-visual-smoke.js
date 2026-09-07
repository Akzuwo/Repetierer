const {app, BrowserWindow} = require('electron');
const fs = require('fs');
const path = require('path');

app.commandLine.appendSwitch('force-prefers-reduced-motion', 'no-preference');

const captureMoments = [
	{ name: '01-lines', at: 1800 },
	{ name: '02-word-reveal', at: 3250 },
	{ name: '03-seven-arrival', at: 3800 },
	{ name: '04-logo-hold', at: 4450 },
	{ name: '05-before-fade', at: 5000 }
];

app.whenReady().then(async () => {
	const outputDirectory = path.join(app.getPath('temp'), 'repetierer-splash-smoke');
	fs.mkdirSync(outputDirectory, {recursive: true});
	const results = [];

	const window = new BrowserWindow({
		width: 560,
		height: 260,
		show: true,
		frame: false,
		resizable: false,
		skipTaskbar: true,
		transparent: true,
		webPreferences: {
			nodeIntegration: true,
			contextIsolation: false,
			backgroundThrottling: false
		}
	});

	await window.loadFile(path.join(__dirname, '..', 'public', 'splash.html'));
	const startedAt = Date.now();

	for (const moment of captureMoments) {
		const remaining = moment.at - (Date.now() - startedAt);
		if (remaining > 0) await new Promise(resolve => setTimeout(resolve, remaining));

		const image = await window.webContents.capturePage();
		const filePath = path.join(outputDirectory, `${moment.name}.png`);
		fs.writeFileSync(filePath, image.toPNG());

		const state = await window.webContents.executeJavaScript(`(() => {
			const splash = document.getElementById('startup-splash');
			const word = document.querySelector('.startup-loader-word');
			const seven = document.querySelector('.startup-loader-seven');
			return {
				classes: splash.className,
				wordBounds: word.getBoundingClientRect().toJSON(),
				sevenBounds: seven.getBoundingClientRect().toJSON(),
				sevenOpacity: getComputedStyle(seven).opacity
			};
		})()`);
		results.push({moment: moment.name, at: moment.at, filePath, state});
		fs.writeFileSync(path.join(outputDirectory, 'results.json'), JSON.stringify(results, null, 2));
	}

	window.setSize(1100, 800);
	await window.webContents.executeJavaScript(`new Promise(resolve => {
		const stylesheet = document.createElement('link');
		stylesheet.rel = 'stylesheet';
		stylesheet.href = 'tailwind.css';
		stylesheet.onload = resolve;
		document.head.appendChild(stylesheet);
		document.body.innerHTML = '<div id="container"><h1 id="title" data-text="Repetierer">Repetierer</h1></div>';
	})`);
	await new Promise(resolve => setTimeout(resolve, 300));
	fs.writeFileSync(
		path.join(outputDirectory, '06-main-title-static.png'),
		(await window.webContents.capturePage()).toPNG()
	);

	window.webContents.debugger.attach('1.3');
	await window.webContents.debugger.sendCommand('Emulation.setEmulatedMedia', {
		features: [{name: 'prefers-reduced-motion', value: 'no-preference'}]
	});
	await window.webContents.executeJavaScript(`(() => {
		document.getElementById('container').classList.add('logo-animation-enabled');
		document.getElementById('title').classList.add('logo-animation-enabled');
	})()`);
	await new Promise(resolve => setTimeout(resolve, 1500));
	fs.writeFileSync(
		path.join(outputDirectory, '07-main-title-effect.png'),
		(await window.webContents.capturePage()).toPNG()
	);
	results.push(await window.webContents.executeJavaScript(`(() => {
		const title = document.getElementById('title');
		const titleStyle = getComputedStyle(title);
		const effectStyle = getComputedStyle(title, '::after');
		return {
			moment: 'main-title-effect',
			state: {
				color: titleStyle.color,
				filter: titleStyle.filter,
				effectAnimation: effectStyle.animationName,
				effectAnimationDuration: effectStyle.animationDuration,
				reducedMotion: matchMedia('(prefers-reduced-motion: reduce)').matches,
				effectOpacity: effectStyle.opacity
			}
		};
	})()`));
	fs.writeFileSync(path.join(outputDirectory, 'results.json'), JSON.stringify(results, null, 2));

	window.destroy();
	app.quit();
});
