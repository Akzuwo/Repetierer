const {app, BrowserWindow} = require('electron');
const fs = require('fs');
const path = require('path');

const isSmokeTest = process.argv.includes('--smoke');

function wait(milliseconds) {
	return new Promise(resolve => setTimeout(resolve, milliseconds));
}

async function captureState(window, outputDirectory, name) {
	const image = await window.webContents.capturePage();
	fs.writeFileSync(path.join(outputDirectory, `${name}.png`), image.toPNG());

	return window.webContents.executeJavaScript(`(() => {
		const canvas = document.querySelector('.intro-canvas');
		const word = document.querySelector('.intro-word');
		const frontSeven = document.querySelector('.intro-seven-front');
		const backSeven = document.querySelector('.intro-seven-behind');
		return {
			name: ${JSON.stringify(name)},
			classes: canvas.className.baseVal,
			word: word.getBoundingClientRect().toJSON(),
			wordOpacity: getComputedStyle(word).opacity,
			frontSevenOpacity: getComputedStyle(frontSeven).opacity,
			backSevenOpacity: getComputedStyle(backSeven).opacity,
			backSeven: backSeven.getBoundingClientRect().toJSON()
		};
	})()`);
}

async function runSmokeTest(window) {
	const outputDirectory = path.join(app.getPath('temp'), 'repetierer-intro-tester-smoke');
	fs.mkdirSync(outputDirectory, {recursive: true});
	const results = [];

	await wait(60);
	results.push(await captureState(window, outputDirectory, 'variant-1-initial'));
	await wait(1750);
	results.push(await captureState(window, outputDirectory, 'variant-1-drawing'));
	await wait(1900);
	results.push(await captureState(window, outputDirectory, 'variant-1-lines-complete'));
	await wait(600);
	results.push(await captureState(window, outputDirectory, 'variant-1-fill'));
	await wait(2100);
	results.push(await captureState(window, outputDirectory, 'variant-1-final'));

	await window.webContents.executeJavaScript('window.introTester.playVariant(2)');
	await wait(60);
	results.push(await captureState(window, outputDirectory, 'variant-2-initial'));
	await wait(1750);
	results.push(await captureState(window, outputDirectory, 'variant-2-drawing'));
	await wait(1900);
	results.push(await captureState(window, outputDirectory, 'variant-2-lines-complete'));
	await wait(1700);
	results.push(await captureState(window, outputDirectory, 'variant-2-final'));

	await window.webContents.executeJavaScript('window.introTester.playVariant(2.5)');
	await wait(6500);
	results.push(await captureState(window, outputDirectory, 'variant-2-5-final'));

	await window.webContents.executeJavaScript('window.introTester.playVariant(3)');
	await wait(3300);
	results.push(await captureState(window, outputDirectory, 'variant-3-fill'));
	await wait(3100);
	results.push(await captureState(window, outputDirectory, 'variant-3-final'));

	fs.writeFileSync(path.join(outputDirectory, 'results.json'), JSON.stringify(results, null, 2));
	window.destroy();
	app.quit();
}

function createWindow() {
	const window = new BrowserWindow({
		width: 780,
		height: 430,
		minWidth: 660,
		minHeight: 380,
		backgroundColor: '#070a10',
		autoHideMenuBar: true,
		skipTaskbar: isSmokeTest,
		title: 'Repetierer – Intro-Tester',
		webPreferences: {
			contextIsolation: true,
			nodeIntegration: false,
			backgroundThrottling: false
		}
	});

	if (isSmokeTest) {
		window.webContents.once('did-finish-load', () => {
			runSmokeTest(window).catch(error => {
				console.error(error);
				app.exit(1);
			});
		});
	}
	window.loadFile(path.join(__dirname, 'index.html'));
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => app.quit());
