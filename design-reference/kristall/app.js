import { GlassRenderer } from './glass.js';
import { LiquidMotion, LiquidMarker, LiquidToggle, LiquidRange, LiquidSurface } from './motion.js';

const $ = selector => document.querySelector(selector);
const $$ = selector => [...document.querySelectorAll(selector)];
const CRYSTAL = Object.freeze({ refraction: 1, frost: 0, tint: .08, brightness: 1, contrast: 0, focus: 0, motion: true });
let renderer;
const clock = new LiquidMotion(() => renderer?.draw());
const nav = new LiquidMarker($('.nav-selection'), clock, { container: $('.top-nav'), radius: 30 });
const dock = new LiquidMarker($('.dock-selection'), clock, { container: $('.dock'), radius: 21 });
const ranges = new Map();
const toggles = new Map();
let selection = 'one';

function select(key, immediate = false) {
  selection = key;
  $$('[data-key]').forEach(button => {
    if (button.dataset.key === key) button.setAttribute('aria-current', 'page');
    else button.removeAttribute('aria-current');
  });
  nav.follow($(`.nav-tab[data-key="${key}"]`), immediate);
  dock.follow($(`.dock-button[data-key="${key}"]`), immediate);
}

function updateRange(input, immediate = false) {
  const progress = (Number(input.value) - Number(input.min)) / (Number(input.max) - Number(input.min));
  input.style.setProperty('--range-fill', `${progress * 100}%`);
  $(`#${input.id}-value`).textContent = `${input.value}%`;
  ranges.get(input.id).update(immediate);
  if (input.id === 'brightness') {
    if (renderer?.ready) renderer.update({ brightness: Number(input.value) / 100 });
    else $('.wallpaper').style.filter = `brightness(${Number(input.value) / 100})`;
  }
}

$$('input[type="range"]').forEach(input => {
  const large = ['brightness', 'volume'].includes(input.id);
  const wrapper = document.createElement('span');
  wrapper.className = large ? 'range-control filled-range glass' : 'range-control';
  if (large) wrapper.dataset.glass = 'pill';
  const fill = document.createElement('span'); fill.className = 'range-fill'; fill.setAttribute('aria-hidden', 'true');
  const thumb = document.createElement('span'); thumb.className = 'range-thumb liquid-marker glass'; thumb.dataset.glass = 'pill'; thumb.setAttribute('aria-hidden', 'true');
  input.before(wrapper); wrapper.append(fill, input, thumb);
  const marker = new LiquidRange(thumb, clock, input, large);
  ranges.set(input.id, marker);
  input.addEventListener('input', () => updateRange(input));
  input.addEventListener('pointerdown', () => marker.press());
  input.addEventListener('pointerup', () => marker.release());
  input.addEventListener('pointercancel', () => marker.release());
});

function syncToggle(input, immediate = false) {
  const track = input.parentElement;
  const size = track.classList.contains('mini-switch') ? 20 : 26;
  toggles.get(input.id).to({ x: input.checked ? track.clientWidth - 2 - size : 2, y: 2, width: size, height: size }, immediate);
}
$$('.switch-input').forEach(input => {
  const marker = new LiquidToggle(input.parentElement.querySelector('.switch-thumb'), clock);
  toggles.set(input.id, marker);
  input.addEventListener('change', () => syncToggle(input));
  marker.bind(input);
});

for (const [selector, marker] of [['.nav-tab', nav], ['.dock-button', dock]]) {
  $$(selector).forEach(button => {
    button.addEventListener('click', () => select(button.dataset.key));
    button.addEventListener('pointerdown', () => marker.press());
    button.addEventListener('pointerup', () => marker.release());
    button.addEventListener('pointercancel', () => marker.release());
  });
}
const surfaces = $$('.icon-button').map(element => {
  const surface = new LiquidSurface(element, clock);
  element.addEventListener('pointerdown', () => surface.press());
  element.addEventListener('pointerup', () => surface.release());
  element.addEventListener('pointercancel', () => surface.release());
  return surface;
});
const release = () => [nav, dock, ...ranges.values(), ...toggles.values(), ...surfaces].forEach(control => control.release());
window.addEventListener('pointerup', release);
window.addEventListener('pointercancel', release);
window.addEventListener('blur', release);

function setWater(playing) {
  document.body.classList.toggle('is-playing', playing);
  $('#water-toggle').setAttribute('aria-pressed', playing);
  $('#water-toggle').setAttribute('aria-label', playing ? 'Wasserbewegung pausieren' : 'Wasserbewegung starten');
  $('#play-path').setAttribute('d', playing ? 'M6 5h4v14H6zm8 0h4v14h-4z' : 'm6 4 12 8-12 8Z');
}
$('#water-toggle').addEventListener('click', () => setWater(!document.body.classList.contains('is-playing')));
new IntersectionObserver(entries => { $('.water-wave').dataset.visible = entries[0].isIntersecting; }).observe($('.water-wave'));
document.addEventListener('visibilitychange', () => {
  if (document.hidden) setWater(false);
  else renderer?.requestDraw();
});

function layout() {
  select(selection, true);
  $$('input[type="range"]').forEach(input => updateRange(input, true));
  $$('.switch-input').forEach(input => syncToggle(input, true));
}
function fallback() {
  document.documentElement.dataset.renderer = 'css';
  $('#glass-canvas').style.visibility = 'hidden';
  $('#renderer-status').textContent = 'Kristall-Fallback · ohne optische Brechung';
  $('.wallpaper').style.filter = `brightness(${Number($('#brightness').value) / 100})`;
}
layout();
try {
  renderer = new GlassRenderer($('#glass-canvas'));
  renderer.update(CRYSTAL);
  await renderer.loadScene('background');
  $('#renderer-status').textContent = 'Kristall · optische Brechung aktiv';
} catch { fallback(); }
document.addEventListener('glass-fallback', fallback);
document.addEventListener('glass-ready', () => {
  renderer.update({ ...CRYSTAL, brightness: Number($('#brightness').value) / 100 });
  $('.wallpaper').style.filter = '';
  $('#renderer-status').textContent = 'Kristall · optische Brechung aktiv';
});
window.addEventListener('resize', layout);
layout();
// Nur Diagnose des Prüfgerüsts; keine Abhängigkeit der Materialmodule.
window.crystalReference = { renderer, clock, material: CRYSTAL };
