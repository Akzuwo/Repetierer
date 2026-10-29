const { gsap } = require('./vendor/ui-motion.js');

// One owner per interaction. Reopening an overlay cancels its pending exit;
// completion callbacks belong to the timeline, rather than a parallel timeout.
const reduced = window.matchMedia('(prefers-reduced-motion: reduce)');
const animations = new Map();
let messageObserver, scrollEffects;
let pageTransition, pendingPageUpdate, pageCompletion, updatingPage = false;
let drawerViewTransition, pendingDrawerUpdate;

function finishDrawerView() {
	drawerViewTransition?.kill(); drawerViewTransition = null;
	const update = pendingDrawerUpdate; pendingDrawerUpdate = null; update?.();
	const content = document.getElementById('drawer-content');
	gsap.set(content, {clearProps: 'transform,opacity,willChange'});
	delete content.dataset.viewMotion;
}
function changeDrawerView(update) {
	finishDrawerView();
	if (reduced.matches || document.hidden || !document.getElementById('drawer').classList.contains('drawer-open')) { update(); return; }
	const content = document.getElementById('drawer-content');
	pendingDrawerUpdate = update;
	content.dataset.viewMotion = 'exit';
	content.style.willChange = 'transform, opacity';
	window.repetiererKristall?.clock.animateFor(550);
	drawerViewTransition = gsap.timeline({onComplete: finishDrawerView})
		.to(content, {y: 7, scale: .985, opacity: 0, duration: .12, ease: 'power2.in'})
		.call(() => {const apply = pendingDrawerUpdate; pendingDrawerUpdate = null; apply?.(); content.dataset.viewMotion = 'enter';})
		.set(content, {y: -12, scale: .985})
		.to(content, {y: 0, scale: 1, opacity: 1, duration: .36, ease: 'power3.out'});
}

function applyPendingPage() {
	const update = pendingPageUpdate;
	pendingPageUpdate = null;
	if (!update) return;
	updatingPage = true;
	try { update(); } finally { updatingPage = false; }
}
function finishPageMotion() {
	pageTransition?.kill(); pageTransition = null;
	applyPendingPage();
	const content = document.getElementById('content');
	gsap.set(content, { clearProps: 'transform,opacity,willChange' });
	delete content.dataset.pageMotion;
	const complete = pageCompletion; pageCompletion = null; complete?.();
}
// Animate the live content, so the shared glass clock follows its geometry.
// Header and scene are siblings and never participate in the pop.
function pageChange(update, afterChange) {
	if (updatingPage) { update(); afterChange?.(); return; }
	pageCompletion = null;
	applyPendingPage(); // Commit an interrupted exit before the next intent.
	pageTransition?.kill();
	if (reduced.matches || document.hidden) { finishPageMotion(); update(); afterChange?.(); return; }
	const content = document.getElementById('content');
	content.dataset.pageMotion = 'exit';
	content.style.willChange = 'transform, opacity';
	pendingPageUpdate = update;
	pageCompletion = afterChange;
	window.repetiererKristall?.clock.animateFor(720);
	const timeline = gsap.timeline({ onComplete: () => {
		if (pageTransition !== timeline) return;
		pageTransition = null;
		gsap.set(content, { clearProps: 'transform,opacity,willChange' });
		delete content.dataset.pageMotion;
		pageCompletion = null;
		afterChange?.();
	} });
	pageTransition = timeline;
	timeline.to(content, { y: 18, scale: .955, opacity: 0, duration: .14, ease: 'power2.in' })
		.call(() => { applyPendingPage(); content.dataset.pageMotion = 'enter'; })
		.set(content, { y: -38, scale: .94, opacity: 0 })
		.to(content, { y: 0, scale: 1, opacity: 1, duration: .46, ease: 'back.out(1.35)' });
	return timeline;
}
function pageHandler(handler) { return (...args) => pageChange(() => handler(...args)); }
function beginPageChange(update) {
	if (updatingPage || reduced.matches || document.hidden) return false;
	pageChange(update); return true;
}

function cancel(element) {
	const animation = animations.get(element);
	if (animation) { animation.kill(); animations.delete(element); }
}

function animate(element, build, finish) {
	cancel(element);
	window.repetiererKristall?.clock.animateFor(520);
	if (reduced.matches || document.hidden) { finish(); return; }
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
		gsap.set(content, { opacity: 0, scaleX: .88, scaleY: .84, y: 22 });
	}
	animate(modal, timeline => {
		timeline.to(modal, { opacity: 1, duration: .16, ease: 'power2.out' }, 0)
			.to(content, { opacity: 1, scaleX: 1, scaleY: 1, y: 0, duration: .42, ease: 'back.out(1.25)' }, 0);
	}, () => {
		gsap.set(modal, { clearProps: 'opacity' });
		gsap.set(content, { clearProps: 'opacity,transform' });
	});
}

function closeDialog(modal, afterClose) {
	if (!modal || modal.classList.contains('modal-closing')) return;
	if (modal.style.display !== 'flex' && !modal.classList.contains('modal-open')) { afterClose?.(); return; }
	const content = modal.querySelector('.modal-content');
	modal.classList.remove('modal-open');
	modal.classList.add('modal-closing');
	modal.setAttribute('aria-hidden', 'true');
	animate(modal, timeline => {
		timeline.to(content, { scaleX: 1.025, scaleY: .97, y: 2, duration: .07, ease: 'power1.out' }, 0)
			.to(content, { opacity: 0, scaleX: .88, scaleY: .82, y: 18, duration: .17, ease: 'power3.in' }, .07)
			.to(modal, { opacity: 0, duration: .16, ease: 'power2.in' }, .08);
	}, () => {
		modal.style.display = 'none';
		modal.classList.remove('modal-closing');
		gsap.set(modal, { clearProps: 'opacity' });
		gsap.set(content, { clearProps: 'opacity,transform' });
		window.repetiererKristall?.refresh();
		if (typeof afterClose === 'function') afterClose();
	});
}

function setGradingVisible(element, visible) {
	const wasVisible = element.dataset.gradingVisible === 'true';
	if (wasVisible === visible && element.dataset.gradingVisible !== undefined) return;
	const initial = element.dataset.gradingVisible === undefined;
	element.dataset.gradingVisible = String(visible);
	cancel(element);
	const controls = [...element.children];
	gsap.set([element, ...controls], {clearProps:'transform,opacity'});
	if (initial || reduced.matches || document.hidden) { element.classList.toggle('update-hidden', !visible); return; }
	if (visible) {
		element.classList.remove('update-hidden');
		animate(element, timeline => timeline.fromTo(controls, {opacity:0,y:22,scale:.9}, {opacity:1,y:0,scale:1,duration:.38,stagger:.035,ease:'back.out(1.5)'}), () => gsap.set(controls,{clearProps:'opacity,transform'}));
	} else {
		animate(element, timeline => timeline.to(element,{opacity:0,y:10,scale:.95,duration:.16,ease:'power2.in'}), () => {element.classList.add('update-hidden');gsap.set(element,{clearProps:'opacity,transform'});});
	}
}
function confirmScore(button, afterFeedback) {
	button.classList.add('score-confirmed'); button.setAttribute('aria-pressed','true');
	const number = button.querySelector('strong');
	animate(button, timeline => timeline.fromTo(number,{scale:.78},{scale:1.18,duration:.14,ease:'back.out(2)'}).to(number,{scale:1,duration:.12,ease:'power2.out'}), () => {gsap.set(number,{clearProps:'transform'});afterFeedback();});
}

function setDrawerOpen(open) {
	const drawer = document.getElementById('drawer');
	const scrim = document.getElementById('drawer-scrim');
	const toggle = document.getElementById('drawer-toggle');
	const bars = toggle.querySelectorAll('span');
	const openingFromRest = open && !animations.has(drawer) && !drawer.classList.contains('drawer-open');
	cancel(drawer);
	gsap.set(drawer, { x: gsap.getProperty(drawer, 'x'), xPercent: gsap.getProperty(drawer, 'xPercent') });
	if (openingFromRest) gsap.set(drawer, {scaleX: .975});
	gsap.set(scrim, { opacity: getComputedStyle(scrim).opacity });
	for (const bar of bars) gsap.set(bar, { y: gsap.getProperty(bar, 'y'), rotation: gsap.getProperty(bar, 'rotation'), scaleX: gsap.getProperty(bar, 'scaleX'), opacity: getComputedStyle(bar).opacity });
	drawer.classList.toggle('drawer-open', open);
	scrim.classList.toggle('scrim-visible', open);
	toggle.classList.toggle('drawer-toggle-open', open);
	toggle.setAttribute('aria-expanded', String(open));
	drawer.setAttribute('aria-hidden', String(!open));
	drawer.inert = !open;
	drawer.dataset.motionState = open ? 'opening' : 'closing';
	animate(drawer, timeline => {
		timeline.to(drawer, { xPercent: open ? 0 : -100, x: open ? 0 : -24, scaleX: open ? 1 : .975, transformOrigin: 'left center', duration: open ? .56 : .38, ease: open ? 'power3.out' : 'power2.inOut' }, 0)
			.to(scrim, { opacity: open ? 1 : 0, duration: open ? .4 : .32, ease: 'power2.inOut' }, 0)
			.to(bars[0], { y: open ? '.47rem' : 0, rotation: open ? 45 : 0, duration: .3 }, 0)
			.to(bars[1], { opacity: open ? 0 : 1, scaleX: open ? 0 : 1, duration: .2 }, 0)
			.to(bars[2], { y: open ? '-.47rem' : 0, rotation: open ? -45 : 0, duration: .3 }, 0);
	}, () => {
		gsap.set([drawer, ...bars], { clearProps: 'transform,opacity' });
		gsap.set(scrim, { clearProps: 'opacity' });
		delete drawer.dataset.motionState;
	});
	window.repetiererKristall?.clock.animateFor(680);
}

function revealView(view) {
	if (!view || view.classList.contains('update-hidden')) return;
	if (updatingPage) return;
	animate(view, timeline => {
		timeline.fromTo(view, { opacity: .3 }, { opacity: 1, duration: .46, ease: 'power3.out' });
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

function refreshScrollEffects() { scrollEffects?.schedule(); }

function initUiMotion() {
 scrollEffects = require('./scrollEffects').initScrollEffects();
	document.documentElement.dataset.motionEngine = 'gsap';
	const drawer = document.getElementById('drawer');
	drawer.inert = true;
	drawer.setAttribute('aria-hidden', 'true');
	const onPreferenceChange = () => {
		if (reduced.matches) finishPageMotion();
		if (reduced.matches) finishDrawerView();
		if (reduced.matches) [...animations.values()].forEach(animation => animation.progress(1));
	};
	reduced.addEventListener('change', onPreferenceChange);
	const onVisibilityChange = () => { if (document.hidden) { finishPageMotion(); finishDrawerView(); } };
	document.addEventListener('visibilitychange', onVisibilityChange);
	// New chat messages acknowledge arrival; static lists have no entrance show.
	messageObserver = new MutationObserver(records => {
		for (const record of records) for (const node of record.addedNodes) {
			if (node.nodeType !== 1 || !node.isConnected || node.classList.contains('ai-user')) continue;
			revealView(node);
		}
	});
	messageObserver.observe(document.getElementById('ai-messages'), { childList: true });
	window.addEventListener('beforeunload', () => {
		pageTransition?.kill(); pendingPageUpdate = null;
		drawerViewTransition?.kill(); pendingDrawerUpdate = null;
		for (const animation of animations.values()) animation.kill();
		animations.clear();
		messageObserver.disconnect();
  scrollEffects.dispose();
		reduced.removeEventListener('change', onPreferenceChange);
		document.removeEventListener('visibilitychange', onVisibilityChange);
	}, { once: true });
}

module.exports = { initUiMotion, refreshScrollEffects, openDialog, closeDialog, setDrawerOpen, changeDrawerView, revealView, exitView, revealStudent, minimizeWindow, restoreWindow, cancel, pageChange, pageHandler, beginPageChange, setGradingVisible, confirmScore };
