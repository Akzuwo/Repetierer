const { gsap } = require('./vendor/ui-motion.js');

// One owner per interaction. Reopening an overlay cancels its pending exit;
// completion callbacks belong to the timeline, rather than a parallel timeout.
const reduced = window.matchMedia('(prefers-reduced-motion: reduce)');
const animations = new Map();
let messageObserver;

function cancel(element) {
	const animation = animations.get(element);
	if (animation) { animation.kill(); animations.delete(element); }
}

function animate(element, build, finish) {
	cancel(element);
	if (reduced.matches) { finish(); return; }
	const timeline = gsap.timeline({ defaults: { ease: 'expo.out' }, onComplete: () => {
		if (animations.get(element) === timeline) animations.delete(element);
		finish();
	} });
	animations.set(element, timeline);
	build(timeline);
}

function openDialog(modal) {
	if (!modal) return;
	const wasVisible = modal.style.display === 'flex';
	const content = modal.querySelector('.modal-content');
	cancel(modal);
	modal.style.display = 'flex';
	modal.classList.remove('modal-closing');
	modal.classList.add('modal-open');
	modal.setAttribute('aria-hidden', 'false');
	if (!wasVisible && !reduced.matches) {
		gsap.set(modal, { opacity: 0 });
		gsap.set(content, { opacity: .4, scale: .99 });
	}
	animate(modal, timeline => {
		timeline.to(modal, { opacity: 1, duration: .18 }, 0)
			.to(content, { opacity: 1, scale: 1, duration: .24 }, 0);
	}, () => {
		gsap.set(modal, { clearProps: 'opacity' });
		gsap.set(content, { clearProps: 'opacity,transform' });
	});
}

function closeDialog(modal, afterClose) {
	if (!modal || modal.classList.contains('modal-closing')) return;
	const content = modal.querySelector('.modal-content');
	modal.classList.remove('modal-open');
	modal.classList.add('modal-closing');
	modal.setAttribute('aria-hidden', 'true');
	animate(modal, timeline => {
		timeline.to(modal, { opacity: 0, duration: .14, ease: 'power2.in' }, 0)
			.to(content, { opacity: .5, scale: .99, duration: .14, ease: 'power2.in' }, 0);
	}, () => {
		modal.style.display = 'none';
		modal.classList.remove('modal-closing');
		gsap.set(modal, { clearProps: 'opacity' });
		gsap.set(content, { clearProps: 'opacity,transform' });
		if (typeof afterClose === 'function') afterClose();
	});
}

function setDrawerOpen(open) {
	const drawer = document.getElementById('drawer');
	const scrim = document.getElementById('drawer-scrim');
	const toggle = document.getElementById('drawer-toggle');
	const bars = toggle.querySelectorAll('span');
	cancel(drawer);
	gsap.set(drawer, { x: gsap.getProperty(drawer, 'x'), xPercent: gsap.getProperty(drawer, 'xPercent') });
	gsap.set(scrim, { opacity: getComputedStyle(scrim).opacity });
	for (const bar of bars) gsap.set(bar, { y: gsap.getProperty(bar, 'y'), rotation: gsap.getProperty(bar, 'rotation'), scaleX: gsap.getProperty(bar, 'scaleX'), opacity: getComputedStyle(bar).opacity });
	drawer.classList.toggle('drawer-open', open);
	scrim.classList.toggle('scrim-visible', open);
	toggle.classList.toggle('drawer-toggle-open', open);
	toggle.setAttribute('aria-expanded', String(open));
	drawer.setAttribute('aria-hidden', String(!open));
	drawer.inert = !open;
	animate(drawer, timeline => {
		timeline.to(drawer, { xPercent: open ? 0 : -100, x: open ? 0 : -24, duration: open ? .28 : .2 }, 0)
			.to(scrim, { opacity: open ? 1 : 0, duration: .18 }, 0)
			.to(bars[0], { y: open ? '.47rem' : 0, rotation: open ? 45 : 0, duration: .2 }, 0)
			.to(bars[1], { opacity: open ? 0 : 1, scaleX: open ? 0 : 1, duration: .12 }, 0)
			.to(bars[2], { y: open ? '-.47rem' : 0, rotation: open ? -45 : 0, duration: .2 }, 0);
	}, () => {
		gsap.set([drawer, ...bars], { clearProps: 'transform,opacity' });
		gsap.set(scrim, { clearProps: 'opacity' });
	});
}

function revealView(view) {
	if (!view || view.classList.contains('update-hidden')) return;
	animate(view, timeline => {
		timeline.fromTo(view, { opacity: .65 }, { opacity: 1, duration: .18 });
	}, () => gsap.set(view, { clearProps: 'opacity' }));
}

function exitView(view, afterExit) {
	animate(view, timeline => timeline.to(view, { opacity: 0, duration: .12, ease: 'power2.in' }), () => {
		view.classList.add('update-hidden');
		gsap.set(view, { clearProps: 'opacity' });
		afterExit();
	});
}

function revealStudent(name) {
	animate(name, timeline => {
		timeline.fromTo(name, { opacity: .5, clipPath: 'inset(0 0 18% 0)' }, { opacity: 1, clipPath: 'inset(0 0 0% 0)', duration: .28 });
	}, () => gsap.set(name, { clearProps: 'opacity,clipPath' }));
}

function minimizeWindow(container, onMinimize) {
	animate(container, timeline => timeline.to(container, { opacity: 0, scale: .985, duration: .16, ease: 'power2.in' }), () => {
		gsap.set(container, { opacity: 0, scale: .985 });
		onMinimize();
	});
}

function restoreWindow(container) {
	animate(container, timeline => timeline.to(container, { opacity: 1, scale: 1, duration: .22 }), () => {
		gsap.set(container, { clearProps: 'transform,opacity' });
	});
}

function initUiMotion() {
	document.documentElement.dataset.motionEngine = 'gsap';
	const drawer = document.getElementById('drawer');
	drawer.inert = true;
	drawer.setAttribute('aria-hidden', 'true');
	const onPreferenceChange = () => {
		if (reduced.matches) [...animations.values()].forEach(animation => animation.progress(1));
	};
	reduced.addEventListener('change', onPreferenceChange);
	// New chat messages acknowledge arrival; static lists have no entrance show.
	messageObserver = new MutationObserver(records => {
		for (const record of records) for (const node of record.addedNodes) {
			if (node.nodeType !== 1 || !node.isConnected) continue;
			revealView(node);
		}
	});
	messageObserver.observe(document.getElementById('ai-messages'), { childList: true });
	window.addEventListener('beforeunload', () => {
		for (const animation of animations.values()) animation.kill();
		animations.clear();
		messageObserver.disconnect();
		reduced.removeEventListener('change', onPreferenceChange);
	}, { once: true });
}

module.exports = { initUiMotion, openDialog, closeDialog, setDrawerOpen, revealView, exitView, revealStudent, minimizeWindow, restoreWindow, cancel };
