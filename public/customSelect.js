const {gsap} = require('./vendor/ui-motion');
const {mountMenuGlass} = require('./kristall/menuGlass');

// Keep the native select as the data source; the visible control owns a
// top-layer listbox, so no scrolling page can clip its options.
function initCustomSelects() {
	const controls = new Map(), lifecycle = new AbortController(), signal = lifecycle.signal;
	const reduced = matchMedia('(prefers-reduced-motion: reduce)');
	let active = null, frame = 0, sequence = 0;
	function create(select) {
		const wrapper = document.createElement('span'); wrapper.className = 'liquid-select';
		select.before(wrapper); wrapper.append(select);
		select.classList.add('liquid-select-native'); select.tabIndex = -1; select.setAttribute('aria-hidden', 'true');
		const trigger = document.createElement('button'); trigger.type = 'button'; trigger.className = 'liquid-select-trigger';
		trigger.setAttribute('role', 'combobox'); trigger.setAttribute('aria-haspopup', 'listbox'); trigger.setAttribute('aria-expanded', 'false');
		const label = document.createElement('span'); label.className = 'liquid-select-value';
		const arrow = document.createElementNS('http://www.w3.org/2000/svg', 'svg'); arrow.setAttribute('viewBox', '0 0 24 24'); arrow.setAttribute('aria-hidden', 'true');
		arrow.innerHTML = '<path d="m6 9 6 6 6-6"/>'; trigger.append(label, arrow); wrapper.append(trigger);
		const menu = document.createElement('div'); menu.id = `liquid-select-menu-${++sequence}`; menu.className = 'liquid-select-menu';
		menu.setAttribute('popover', 'manual'); menu.setAttribute('role', 'listbox'); document.body.append(menu);
		const items = document.createElement('div'); items.className = 'liquid-select-options'; menu.append(items);
		const glass = mountMenuGlass(menu);
		trigger.setAttribute('aria-controls', menu.id);
		let signature = '', selected = -1, opened = false, tween, search = '', searchAt = 0;
		const name = select.getAttribute('aria-label') || [...(select.labels?.[0]?.childNodes || [])].filter(n => n.nodeType === 3).map(n => n.textContent).join('').trim() || 'Auswahl';
		menu.setAttribute('aria-label', name);
		function setActive(index, highlight = false) {
			const options = [...select.options];
			if (index < 0 || options[index]?.disabled) return;
			selected = index;
			[...items.children].forEach((item, i) => item.classList.toggle('is-active', highlight && i === selected));
			if (opened && items.children[index]) {trigger.setAttribute('aria-activedescendant', items.children[index].id); items.children[index].scrollIntoView({block:'nearest'});}
		}
		function sync() {
			const text = select.selectedOptions[0]?.textContent || 'Auswählen';
			if (label.textContent !== text) label.textContent = text;
			const disabled = select.disabled || select.classList.contains('disabled');
			if (trigger.disabled !== disabled) trigger.disabled = disabled;
			trigger.setAttribute('aria-label', `${name}: ${text}`);
			const next = JSON.stringify([...select.options].map(o => [o.textContent, o.value, o.disabled]));
			if (signature !== next) {
				signature = next; items.replaceChildren();
				[...select.options].forEach((option, index) => {
					const item = document.createElement('div'); item.className = 'liquid-select-option'; item.id = `${menu.id}-${index}`;
					item.setAttribute('role', 'option'); item.setAttribute('aria-disabled', String(option.disabled)); item.dataset.index = index;
					const text = document.createElement('span'); text.textContent = option.textContent;
					item.append(text); items.append(item);
				});
			}
			[...items.children].forEach((item, i) => item.setAttribute('aria-selected', String(i === select.selectedIndex)));
			if (opened && (trigger.disabled || !wrapper.checkVisibility())) close(true);
		}
		function position() {
			const box = trigger.getBoundingClientRect(), below = innerHeight - box.bottom - 18, above = box.top - 18;
			const up = below < 160 && above > below;
			const width = Math.min(Math.max(box.width, 180), innerWidth - 24);
			menu.style.width = `${width}px`; menu.style.maxHeight = `${Math.max(60, Math.min(320, up ? above : below))}px`;
			menu.style.left = `${Math.max(12, Math.min(box.right - width, innerWidth - width - 12))}px`;
			menu.style.top = `${up ? box.top - menu.offsetHeight - 8 : box.bottom + 8}px`;
			menu.style.transformOrigin = up ? '50% 100%' : '50% 0';
		}
		function open(keyboard = false) {
			if (trigger.disabled) return;
			active?.close(true); active = api; sync(); tween?.kill();
			const wasShown = menu.matches(':popover-open');
			opened = true; menu.showPopover(); position(); trigger.setAttribute('aria-expanded', 'true');
			trigger.focus({preventScroll:true});
			search = ''; searchAt = 0;
			setActive(select.selectedIndex >= 0 ? select.selectedIndex : [...select.options].findIndex(o => !o.disabled),keyboard);
			glass.prepare();
			if (reduced.matches) gsap.set(menu, {clearProps:'transform,opacity'});
			else {
				if (!wasShown) gsap.set(menu, {opacity:0, y:-8, scaleX:.98, scaleY:.94});
				tween = gsap.to(menu, {opacity:1, y:0, scaleX:1, scaleY:1, duration:.34, ease:'power3.out', onUpdate:glass.refresh, onComplete:() => {gsap.set(menu,{clearProps:'transform,opacity'}); glass.refresh();}});
			}
		}
		function close(immediate = false) {
			opened = false; if (active === api) active = null;
			trigger.setAttribute('aria-expanded', 'false'); trigger.removeAttribute('aria-activedescendant'); tween?.kill();
			const finish = () => {menu.hidePopover(); gsap.set(menu,{clearProps:'transform,opacity'});};
			if (immediate || reduced.matches || !menu.matches(':popover-open')) finish();
			else tween = gsap.to(menu, {opacity:0, y:-5, scaleY:.96, duration:.18, ease:'power2.in', onUpdate:glass.refresh, onComplete:finish});
		}
		function choose(index) {
			if (select.options[index]?.disabled || !select.options[index]) return;
			select.selectedIndex = index; sync(); close();
			select.dispatchEvent(new Event('input', {bubbles:true})); select.dispatchEvent(new Event('change', {bubbles:true}));
		}
		trigger.addEventListener('click', event => {event.preventDefault(); opened ? close() : open();}, {signal});
		trigger.addEventListener('keydown', event => {
			if (event.key === 'Tab') {close(); return;}
			if (event.key === 'Escape') {if (opened) {event.preventDefault(); event.stopPropagation(); close();} return;}
			if (['ArrowDown','ArrowUp','Home','End'].includes(event.key)) {
				event.preventDefault(); if (!opened) open(true);
				const valid = [...select.options].map((o,i) => o.disabled ? -1 : i).filter(i => i >= 0), current = valid.indexOf(selected);
				setActive(event.key === 'Home' ? valid[0] : event.key === 'End' ? valid.at(-1) : valid[Math.max(0, Math.min(valid.length - 1, current + (event.key === 'ArrowDown' ? 1 : -1)))],true);
			} else if (event.key === 'Enter' || event.key === ' ') {event.preventDefault(); opened ? choose(selected) : open(true);}
			else if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
				event.preventDefault(); if (!opened) open(true);
				search = performance.now() - searchAt > 650 ? event.key : search + event.key; searchAt = performance.now();
				setActive([...select.options].findIndex(o => !o.disabled && o.textContent.trim().toLocaleLowerCase().startsWith(search.toLocaleLowerCase())),true);
			}
		}, {signal});
		menu.addEventListener('pointerdown', event => event.preventDefault(), {signal});
		menu.addEventListener('pointermove', event => {const item = event.target.closest('[data-index]'); if (item) setActive(Number(item.dataset.index),true);}, {signal});
		menu.addEventListener('click', event => {const item = event.target.closest('[data-index]'); if (item) choose(Number(item.dataset.index));}, {signal});
		select.addEventListener('change', sync, {signal});
		select.addEventListener('click', event => event.preventDefault(), {signal});
		const api = {select, wrapper, menu, trigger, glass, sync, close, contains:node => wrapper.contains(node) || menu.contains(node), dispose(){if (active === api) active = null; tween?.kill(); glass.dispose(); menu.remove();}};
		controls.set(select, api); sync();
	}
	function refresh() {
		frame = 0;
		for (const [select, control] of controls) if (!select.isConnected) {control.dispose(); controls.delete(select);}
		for (const select of document.querySelectorAll('select[data-custom-select]')) if (!controls.has(select)) create(select);
		for (const control of controls.values()) control.sync();
	}
	const schedule = () => {if (!frame) frame = requestAnimationFrame(refresh);};
	const observer = new MutationObserver(schedule);
	observer.observe(document.getElementById('container'), {subtree:true, childList:true, attributes:true, attributeFilter:['disabled','class','hidden','selected']});
	document.addEventListener('pointerdown', event => {if (active && !active.contains(event.target)) active.close();}, {capture:true,signal});
	document.addEventListener('scroll', event => {if (active && !active.menu.contains(event.target)) active.close(true);}, {capture:true,passive:true,signal});
	window.addEventListener('resize', () => active?.close(true), {signal});
	reduced.addEventListener('change', () => {for (const control of controls.values()) control.close(true);}, {signal});
	refresh();
	const instance = {refresh, closeAll:() => {for (const control of controls.values()) control.close();}, controls};
	window.repetiererSelects = instance;
	window.addEventListener('beforeunload', () => {lifecycle.abort(); observer.disconnect(); cancelAnimationFrame(frame); for (const control of controls.values()) control.dispose(); controls.clear();}, {once:true});
	return instance;
}
module.exports = {initCustomSelects};
