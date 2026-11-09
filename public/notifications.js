const {gsap} = require('./vendor/ui-motion');
const {ipcRenderer} = require('electron');
const {mountMenuGlass} = require('./kristall/menuGlass');
const {normalizeNotification,normalizeDuration,MAX_VISIBLE,DEFAULT_DURATION_SECONDS} = require('../src/notificationPolicy');
let instance;
const ICONS = {
	error:'M12 8v5M12 17h.01M12 3a9 9 0 1 0 0 18 9 9 0 0 0 0-18',
	warning:'m12 3 10 18H2L12 3ZM12 9v5M12 17h.01',
	info:'M12 11v6M12 7h.01M12 3a9 9 0 1 0 0 18 9 9 0 0 0 0-18',
	success:'m7 12 3 3 7-7M12 3a9 9 0 1 0 0 18 9 9 0 0 0 0-18'
};
function initNotifications() {
	if (instance) return instance;
	const root = document.createElement('aside'); root.id = 'notification-layer'; root.setAttribute('popover','manual'); root.setAttribute('aria-label','Benachrichtigungen');
	document.body.append(root); root.showPopover();
	const glass = mountMenuGlass(root,{surfaces:'.notification-card'});
	const lifecycle = new AbortController();
	const signal = lifecycle.signal;
	const cards = new Map(), reduced = matchMedia('(prefers-reduced-motion: reduce)');
	let duration = DEFAULT_DURATION_SECONDS, serial = 0, disposed = false, clearing = false;
	function topmost() {if (root.matches(':popover-open')) root.hidePopover(); root.showPopover();}
	function reflow(previous) {
		if (disposed || clearing) return;
		for (const card of cards.values()) {
			const before = previous.get(card.id); if (before === undefined || card.state === 'enter') continue;
			const shift = before - card.element.getBoundingClientRect().top;
			if (Math.abs(shift) > 1 && !reduced.matches && !document.hidden) {
				card.layout?.kill();
				card.layout = gsap.fromTo(card.element,{y:shift},{y:0,duration:.28,ease:'power3.out',onUpdate:glass.refresh});
			}
		}
		glass.refresh();
	}
	const positions = () => new Map([...cards.values()].map(card=>[card.id,card.element.getBoundingClientRect().top]));
	function destroy(card) {
		clearTimeout(card.timer); card.animation?.kill(); card.layout?.kill();
		const before = positions(); cards.delete(card.id); card.element.remove(); reflow(before);
	}
	function dismiss(id, immediate = false) {
		const card = cards.get(id); if (!card || card.state === 'exit') return;
		clearTimeout(card.timer); card.animation?.kill(); card.layout?.kill();
		card.state = 'exit'; card.element.dataset.motion = 'exit';
		if (immediate || reduced.matches || document.hidden) {destroy(card); return;}
		card.animation = gsap.to(card.element,{y:-24,scaleX:.97,scaleY:.92,opacity:0,duration:.22,ease:'power2.in',onUpdate:glass.refresh,onComplete:()=>destroy(card)});
	}
	function render(card) {
		card.element.dataset.type = card.type;
		card.element.setAttribute('role',card.type === 'error' || card.type === 'warning' ? 'alert' : 'status');
		card.element.setAttribute('aria-atomic','true');
		card.icon.innerHTML = `<path d="${ICONS[card.type]}"/>`;
		card.heading.textContent = card.title; card.body.textContent = card.text;
		card.element.dataset.durationSeconds = String(card.durationSeconds);
	}
	function show(type, text, options = {}) {
		if (disposed) return null;
		const data = normalizeNotification(type,text,options,duration); if (!data) return null;
		topmost();
		const existing = data.key && [...cards.values()].find(card=>card.key === data.key);
		if (existing) {
			const wasExiting = existing.state === 'exit';
			clearTimeout(existing.timer); existing.animation?.kill(); existing.layout?.kill();
			Object.assign(existing,data); render(existing);
			existing.state = 'idle'; existing.element.dataset.motion = 'idle';
			gsap.set(existing.element,{clearProps:'transform,opacity'});
			existing.timer = setTimeout(()=>dismiss(existing.id),data.durationSeconds*1000);
			if (wasExiting) glass.prepare(); glass.refresh(); return existing.id;
		}
		while (cards.size >= MAX_VISIBLE) destroy(cards.values().next().value);
		const element = document.createElement('div'); element.className = 'notification-card';
		const icon = document.createElementNS('http://www.w3.org/2000/svg','svg'); icon.setAttribute('viewBox','0 0 24 24'); icon.setAttribute('aria-hidden','true');
		const copy = document.createElement('div'); copy.className = 'notification-copy';
		const heading = document.createElement('strong'), body = document.createElement('p'); copy.append(heading,body); element.append(icon,copy);
		const card = {...data,id:`notice-${++serial}`,element,icon,heading,body,state:'enter'};
		element.dataset.notificationId = card.id; element.dataset.motion = 'enter';
		cards.set(card.id,card); render(card); root.append(element); glass.prepare();
		const landed = () => {card.state = 'idle'; element.dataset.motion = 'idle'; gsap.set(element,{clearProps:'transform,opacity'});glass.refresh();};
		if (reduced.matches || document.hidden) landed();
		else card.animation = gsap.timeline({onComplete:landed})
			.fromTo(element,{opacity:0,y:-56,scaleX:.94,scaleY:.82},{opacity:1,y:2,scaleX:1.018,scaleY:.97,duration:.28,ease:'power3.out',onUpdate:glass.refresh})
			.to(element,{y:0,scaleX:1,scaleY:1,duration:.18,ease:'power2.out',onUpdate:glass.refresh});
		card.timer = setTimeout(()=>dismiss(card.id),data.durationSeconds*1000);
		return card.id;
	}
	function configure(settings) {duration = normalizeDuration(settings?.notificationDurationSeconds);}
	const receive = (_event,payload) => {if (payload) show(payload.type,payload.text,payload);};
	const settings = (_event,value) => configure(value);
	ipcRenderer.on('ui:notification',receive);
	for (const event of ['ui-settings','settings-data','settings-saved']) ipcRenderer.on(event,settings);
	const finishMotion = () => {
		for (const card of [...cards.values()]) {
			if (card.state === 'exit') destroy(card);
			else {card.animation?.kill();card.layout?.kill();gsap.set(card.element,{clearProps:'transform,opacity'});card.state='idle';card.element.dataset.motion='idle';}
		}
		glass.refresh();
	};
	reduced.addEventListener('change',()=>{if (reduced.matches) finishMotion();},{signal});
	document.addEventListener('visibilitychange',()=>{if (document.hidden) finishMotion();},{signal});
	document.addEventListener('glass-ready',()=>{if (cards.size) glass.prepare();},{signal});
	instance = {show,configure,dismiss,get durationSeconds(){return duration;},getActive:()=>[...cards.values()].map(({id,type,title,text,durationSeconds,key})=>({id,type,title,text,durationSeconds,key})),clear(){clearing=true;for (const card of [...cards.values()]) destroy(card);clearing=false;if(!disposed) glass.refresh();},glass};
	for (const type of ['error','warning','info','success']) instance[type] = (text,options)=>show(type,text,options);
	instance.warn = instance.warning; instance.Error = instance.error;
	window.notify = window.Notifications = instance;
	window.addEventListener('beforeunload',()=>{
		disposed = true; lifecycle.abort(); instance.clear(); glass.dispose(); root.remove();
		ipcRenderer.removeListener('ui:notification',receive);
		for (const event of ['ui-settings','settings-data','settings-saved']) ipcRenderer.removeListener(event,settings);
	},{once:true});
	return instance;
}
const api = {initNotifications};
for (const type of ['error','warning','info','success']) api[type] = (text,options)=>initNotifications()[type](text,options);
api.warn = api.warning; api.Error = api.error;
module.exports = api;
