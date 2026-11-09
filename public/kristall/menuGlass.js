const {GlassRenderer} = require('./glass');

function mountMenuGlass(root, {surfaces = null} = {}) {
	const canvas = document.createElement('canvas'); canvas.className = 'menu-glass-canvas'; canvas.setAttribute('aria-hidden','true');
	root.prepend(canvas); root.dataset.layerRenderer = 'css';
	let renderer, loading, disposed = false;
	async function prepare() {
		if (disposed || !window.repetiererKristall?.renderer?.ready || document.documentElement.dataset.renderer !== 'webgl') return;
		try {
			if (!renderer) {
				renderer = new GlassRenderer(canvas,{root,surfaces});
				renderer.update(window.repetiererKristall.material);
				loading = renderer.loadScene('abstract');
			}
			await loading;
			if (!disposed) {renderer.resize(); renderer.draw();}
		} catch {root.dataset.layerRenderer = 'css'; canvas.style.visibility = 'hidden';}
	}
	return {prepare, refresh(){if (!disposed && renderer?.ready) {renderer.resize(); renderer.draw();}}, get renderer(){return renderer;}, dispose(){disposed = true; renderer?.dispose(); canvas.remove();}};
}
module.exports = {mountMenuGlass};
