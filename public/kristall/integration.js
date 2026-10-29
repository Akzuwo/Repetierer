const { GlassRenderer } = require('./glass.js');
const { LiquidMotion, LiquidMarker, LiquidToggle, LiquidRange, LiquidSurface } = require('./motion.js');

const CRYSTAL = Object.freeze({ refraction: 1, frost: 0, tint: .08, brightness: 1, contrast: 0, focus: 0, motion: false });
let instance;

function initKristall() {
  if (instance) return instance;
  const canvas = document.getElementById('glass-canvas');
  const lifecycle = new AbortController();
  const signal = lifecycle.signal;
  let renderer, refreshFrame = 0, disposed = false, completionOn = false;
  const clock = new LiquidMotion(() => renderer?.draw());
  const controls = new Map();
  const navigations = new Map();
  const dismissBindings = new WeakSet();
  const panelSelector = '.completion-receipt, #title-bar, .modal-content, #drawer, #content > .neum:not(#participation-view), .session-tabs, .update-panel';
  const controlSelector = 'button, input:not([type="checkbox"]):not([type="radio"]):not([type="range"]):not([type="hidden"]), select, textarea, .attendance-person, .assessment-round-card';
  const mark = (element, type = 'pill') => {
    if (!element.hasAttribute('data-glass')) {
      element.dataset.glass = type;
      element.classList.add('crystal-glass');
      renderer?.observer.observe(element);
    }
  };
  const pressBindings = (element, controller) => {
    element.addEventListener('pointerdown', event => { if (event.button === 0 && !element.disabled && !element.classList.contains('disabled')) controller.press(); }, { signal });
    element.addEventListener('pointerup', () => controller.release(), { signal });
    element.addEventListener('pointercancel', () => controller.release(), { signal });
  };

  function toggle(input) {
    if (controls.has(input)) return;
    const track = document.createElement('span');
    track.className = 'liquid-switch';
    input.before(track);
    input.parentElement.querySelector('.toggle-switch')?.remove();
    track.append(input);
    const base = document.createElement('span');
    base.className = 'switch-track';
    base.setAttribute('aria-hidden', 'true');
    const thumb = document.createElement('span');
    thumb.className = 'switch-thumb liquid-marker crystal-glass';
    thumb.dataset.glass = 'pill';
    thumb.setAttribute('aria-hidden', 'true');
    track.append(base, thumb);
    input.classList.add('switch-input');
    mark(track);
    const controller = new LiquidToggle(thumb, clock);
    const sync = (immediate = false) => controller.to({ x: input.checked ? 24 : 2, y: 2, width: 26, height: 26 }, immediate);
    controls.set(input, { controller, sync });
    controller.bind(input);
    input.addEventListener('change', () => sync(), { signal });
    sync(true);
  }

  function range(input) {
    if (controls.has(input)) return;
    const wrapper = document.createElement('span');
    wrapper.className = 'range-control filled-range';
    input.before(wrapper);
    const fill = document.createElement('span');
    fill.className = 'range-fill';
    const thumb = document.createElement('span');
    thumb.className = 'range-thumb liquid-marker crystal-glass';
    thumb.dataset.glass = 'pill';
    fill.setAttribute('aria-hidden', 'true');
    thumb.setAttribute('aria-hidden', 'true');
    wrapper.append(fill, input, thumb);
    mark(wrapper);
    const controller = new LiquidRange(thumb, clock, input, true);
    controls.set(input, { controller, sync: immediate => controller.update(immediate) });
    input.addEventListener('input', () => controller.update(), { signal });
    pressBindings(input, controller);
    controller.update(true);
  }

  function navigation(container) {
    let nav = navigations.get(container);
    if (!nav) {
      const marker = document.createElement('span');
      marker.className = 'liquid-marker crystal-glass nav-selection';
      marker.dataset.glass = 'pill';
      marker.setAttribute('aria-hidden', 'true');
      container.prepend(marker);
      mark(container);
      nav = { marker: new LiquidMarker(marker, clock, { container, radius: 30 }), target: null };
      navigations.set(container, nav);
    }
    container.querySelectorAll('button').forEach(button => {
      if (!controls.has(button)) { controls.set(button, { controller: nav.marker }); pressBindings(button, nav.marker); }
    });
    const target = container.querySelector('.btn-1, .drawer-tab-active, [aria-current="page"]');
    if (target !== nav.target) { nav.marker.follow(target, !nav.target); nav.target = target; }
    else if (target && !clock.active.has(nav.marker)) nav.marker.follow(target, true);
  }

  function refresh() {
    refreshFrame = 0;
    if (disposed) return;
    const completion = !!document.querySelector('.completion-effect.is-active');
    if(completion && !completionOn) clock.animateFor(1800);
    completionOn = completion;
    // The optical canvas samples the scene, not the HTML. Occlude the inactive
    // HTML layer while a modal is on top so background labels cannot bleed in.
    const modal = [...document.querySelectorAll('.modal')].some(element => element.classList.contains('modal-open') || element.classList.contains('modal-closing'));
    const overlay = modal ? 'true' : 'false';
    if (document.documentElement.dataset.crystalOverlay !== overlay) document.documentElement.dataset.crystalOverlay = overlay;
    document.documentElement.dataset.crystalDrawer = String(document.getElementById('drawer').classList.contains('drawer-open'));
    document.querySelectorAll('.modal').forEach(modal => {
      if (dismissBindings.has(modal)) return;
      dismissBindings.add(modal);
      modal.addEventListener('click', event => {
        if (event.target !== modal || !modal.classList.contains('modal-open') || modal.classList.contains('modal-closing')) return;
        const candidates=[...modal.querySelectorAll('.modal-buttons button, .ai-header button, .debug-dismiss')];
        const target=candidates.find(button=>/abbrechen|schliessen|schließen|zurück|verstanden|später/i.test(button.textContent+button.title+button.getAttribute('aria-label')));
        if (target && !target.disabled) target.click();
        else require('../uiMotion').closeDialog(modal);
      }, {signal});
    });
    document.querySelectorAll(panelSelector).forEach(element => { mark(element, 'panel'); if (!element.classList.contains('floating-panel')) element.classList.add('floating-panel'); });
    document.querySelectorAll('.toggle-input').forEach(toggle);
    document.querySelectorAll('input[type="range"]').forEach(range);
    document.querySelectorAll(controlSelector).forEach(element => {
      if (element.matches('#grade-input, #ai-input, #ai-send-btn, #participation-summary-btn, #settings-back-btn, .modal-close-action') || element.closest('#title-bar, #drawer, .session-tabs, .drawer-tabs')) return;
      mark(element, element.matches('.ico, .help-button, .history-btn') ? 'circle' : 'pill');
      if (!controls.has(element) && element.matches('button')) {
        const controller = new LiquidSurface(element, clock);
        controls.set(element, { controller });
        pressBindings(element, controller);
      }
    });
    document.querySelectorAll('.session-tabs, .drawer-tabs').forEach(navigation);
    for (const [element, control] of controls) {
      if (!element.isConnected) { clock.active.delete(control.controller); controls.delete(element); }
      else if (control.sync && !clock.active.has(control.controller)) control.sync(true);
    }
    renderer?.requestDraw();
  }
  const schedule = () => { if (!refreshFrame && !disposed) refreshFrame = requestAnimationFrame(refresh); };
  const release = () => { for (const { controller } of controls.values()) controller.release(); for (const { marker } of navigations.values()) marker.release(); };
  ['pointerup', 'pointercancel', 'blur'].forEach(name => window.addEventListener(name, release, { signal }));
  window.addEventListener('resize', schedule, { signal });
  document.addEventListener('visibilitychange', () => { if (document.hidden) release(); else { schedule(); renderer?.requestDraw(); } }, { signal });
  document.addEventListener('glass-fallback', () => { document.documentElement.dataset.renderer = 'css'; canvas.style.visibility = 'hidden'; }, { signal });
  const observer = new MutationObserver(schedule);
  observer.observe(document.getElementById('container'), { childList: true, subtree: true, attributes: true, attributeFilter: ['class', 'hidden', 'checked', 'value', 'aria-current'] });
  refresh();
  try {
    renderer = new GlassRenderer(canvas);
    renderer.update(CRYSTAL);
    renderer.loadScene('abstract').catch(() => { document.documentElement.dataset.renderer = 'css'; canvas.style.visibility = 'hidden'; });
  } catch {
    document.documentElement.dataset.renderer = 'css';
    canvas.style.visibility = 'hidden';
  }
  instance = { renderer, clock, controls, navigations, material: CRYSTAL, refresh,
    dispose() {
      disposed = true;
      lifecycle.abort(); observer.disconnect(); cancelAnimationFrame(refreshFrame);
      clock.dispose(); renderer?.dispose(); controls.clear(); navigations.clear(); instance = null;
    }
  };
  window.repetiererKristall = instance;
  window.addEventListener('beforeunload', () => instance?.dispose(), { once: true, signal });
  return instance;
}

module.exports = { initKristall };
