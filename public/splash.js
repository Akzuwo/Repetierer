const {ipcRenderer} = require('electron');

const startupSplash = document.getElementById('startup-splash');
let startupFinished = false;

function finishStartupSplash() {
	if (startupFinished) return;
	startupFinished = true;

	if (startupSplash) {
		startupSplash.classList.add('startup-splash-hidden');
	}

	setTimeout(() => {
		ipcRenderer.send('startup-splash-finished');
	}, 650);
}

setTimeout(finishStartupSplash, 2600);
