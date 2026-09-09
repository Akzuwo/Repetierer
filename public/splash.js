const {ipcRenderer} = require('electron');

const startupSplash = document.getElementById('startup-splash');
let startupFinished = false;

const startupTiming = {
	lineStart: 140,
	lineComplete: 3640,
	fillComplete: 4840,
	logoComplete: 6390,
	fadeOut: 500
};

function finishStartupSplash() {
	if (startupFinished) return;
	startupFinished = true;

	if (startupSplash) {
		startupSplash.classList.add('startup-splash-hidden');
	}

	setTimeout(() => {
		ipcRenderer.send('startup-splash-finished');
	}, startupTiming.fadeOut + 80);
}

setTimeout(() => {
	if (!startupSplash) return;
	startupSplash.classList.add('startup-line-drawing');
	startupSplash.querySelectorAll('.startup-line-reveal-motion').forEach(animation => animation.beginElement());
}, startupTiming.lineStart);

setTimeout(() => {
	if (!startupSplash) return;
	startupSplash.classList.add('startup-lines-complete');
	startupSplash.querySelectorAll('.startup-fill-motion').forEach(animation => animation.beginElement());
}, startupTiming.lineComplete);

setTimeout(() => {
	if (startupSplash) startupSplash.classList.add('startup-word-filled', 'startup-seven-visible');
}, startupTiming.fillComplete);

setTimeout(finishStartupSplash, startupTiming.logoComplete);
