const {normalizeNotification,MAX_VISIBLE} = require('./notificationPolicy');
const windows = new Set(), pending = [];

function publish(type, text, options = {}) {
	const payload = normalizeNotification(type,text,options,require('./storage').getAppSettings().notificationDurationSeconds);
	if (!payload) return false;
	const targets = options.target ? [options.target.webContents || options.target] : [...windows].filter(w=>!w.isDestroyed()).map(w=>w.webContents);
	const ready = targets.filter(contents=>!contents.isDestroyed() && !contents.isLoadingMainFrame());
	if (!ready.length) {
		const previous = payload.key ? pending.findIndex(item=>item.key === payload.key) : -1;
		if (previous >= 0) pending.splice(previous,1);
		pending.push(payload); if (pending.length > MAX_VISIBLE) pending.shift();
		return true;
	}
	for (const contents of ready) contents.send('ui:notification',payload);
	return true;
}
function registerWindow(window) {
	if (windows.has(window)) return;
	windows.add(window);
	window.webContents.on('did-finish-load',()=>{
		if (window.isDestroyed()) return;
		for (const payload of pending.splice(0)) window.webContents.send('ui:notification',payload);
	});
	window.once('closed',()=>windows.delete(window));
}
const api = {registerWindow};
for (const type of ['error','warning','info','success']) api[type] = (text,options)=>publish(type,text,options);
api.warn = api.warning; api.Error = api.error;
module.exports = api;
