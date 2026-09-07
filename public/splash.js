const {ipcRenderer} = require('electron');

const startupSplash = document.getElementById('startup-splash');
let startupFinished = false;

const startupTiming = {
	lineAnimation: 2600,
	dimPause: 320,
	wordReveal: 720,
	sevenArrival: 520,
	logoHold: 900,
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
	if (startupSplash) startupSplash.classList.add('startup-lines-complete');
}, startupTiming.lineAnimation);

setTimeout(() => {
	if (startupSplash) startupSplash.classList.add('startup-word-visible');
}, startupTiming.lineAnimation + startupTiming.dimPause);

setTimeout(() => {
	if (startupSplash) startupSplash.classList.add('startup-seven-visible');
}, startupTiming.lineAnimation + startupTiming.dimPause + startupTiming.wordReveal);

setTimeout(
	finishStartupSplash,
	startupTiming.lineAnimation
		+ startupTiming.dimPause
		+ startupTiming.wordReveal
		+ startupTiming.sevenArrival
		+ startupTiming.logoHold
);
