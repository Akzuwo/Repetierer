const { MODE_DRAWS, resolvePreset } = require('./vendor/ui-motion.js');

// Vanilla binding of the original ThinkingOrb canvas contract. Geometry and
// tuning come unchanged from thinking-orbs/engine, without a React renderer.
function createThinkingOrb(canvas, { state = 'working', size = 64, paused = false } = {}) {
	const ctx = canvas.getContext('2d');
	const reduced = window.matchMedia('(prefers-reduced-motion: reduce)');
	let preset = resolvePreset(state, size), visible = false, frameId = 0, destroyed = false;
	const dpr = Math.min(2, window.devicePixelRatio || 1);
	canvas.width = Math.round(size * dpr);
	canvas.height = Math.round(size * dpr);
	canvas.dataset.orbState = state;
	canvas.dataset.orbMotion = 'paused';
	function paint(time) {
		if (!ctx || destroyed) return;
		ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
		ctx.clearRect(0, 0, size, size);
		MODE_DRAWS[preset.mode](ctx, size, time, true, preset.opts);
	}
	function stop() { cancelAnimationFrame(frameId); frameId = 0; canvas.dataset.orbMotion = 'paused'; }
	function loop(time) {
		if (!canvas.isConnected) { destroy(); return; }
		paint(time / 1000 * preset.speed);
		frameId = requestAnimationFrame(loop);
	}
	function update() {
		stop();
		if (destroyed) return;
		paint(reduced.matches ? .6 : performance.now() / 1000 * preset.speed);
		if (visible && !paused && !reduced.matches && document.visibilityState !== 'hidden') {
			canvas.dataset.orbMotion = 'running';
			frameId = requestAnimationFrame(loop);
		}
	}
	const observer = new IntersectionObserver(([entry]) => { visible = entry.isIntersecting; update(); });
	observer.observe(canvas);
	reduced.addEventListener('change', update);
	document.addEventListener('visibilitychange', update);
	function destroy() { destroyed = true; stop(); observer.disconnect(); reduced.removeEventListener('change', update); document.removeEventListener('visibilitychange', update); }
	update();
	return {
		setState(next) { preset = resolvePreset(next, size); canvas.dataset.orbState = next; update(); },
		setPaused(value) { paused = value; update(); },
		destroy
	};
}

module.exports = { createThinkingOrb };
