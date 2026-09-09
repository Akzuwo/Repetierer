(() => {
	const mount = document.getElementById('intro-mount');
	const status = document.getElementById('tester-status');
	const replayButton = document.getElementById('replay-button');
	const variantButtons = [...document.querySelectorAll('[data-variant]')];
	const timers = [];
	let currentVariant = 1;
	let runId = 0;
	const variantLabels = {
		1: 'Variante 1',
		2: 'Variante 2',
		2.5: 'Variante 2 v2',
		3: 'Variante 3'
	};

	const timing = {
		lineStart: 140,
		lineComplete: 3640,
		fillComplete: 4840,
		variantOneComplete: 6390,
		variantTwoComplete: 5290,
		variantTwoFiveComplete: 6390,
		variantThreeFill: 2600,
		variantThreeFilled: 3800,
		variantThreeFade: 4050,
		variantThreeSeven: 4650,
		variantThreeComplete: 6200
	};

	function schedule(callback, delay, activeRun) {
		const timer = setTimeout(() => {
			if (activeRun === runId) callback();
		}, delay);
		timers.push(timer);
	}

	function clearRun() {
		runId += 1;
		timers.splice(0).forEach(timer => clearTimeout(timer));
	}

	function playerMarkup(variant) {
		const frontSevenX = variant === 3 ? 380 : 600;
		const lineMask = variant === 3 ? '' : ' mask="url(#line-reveal-mask)"';
		const lineRevealEnd = variant === 1 ? {x1: 680, x2: 840} : {x1: 780, x2: 940};
		const variantClass = `variant-${String(variant).replace('.', '-')}`;
		return `
			<svg class="intro-canvas ${variantClass}" viewBox="0 0 760 220" role="img" aria-label="Repetierer Intro, ${variantLabels[variant]}">
				<defs>
					<linearGradient gradientUnits="userSpaceOnUse" y2="0" x2="760" y1="150" x1="0" id="line-gradient">
						<stop stop-color="#0f1113"></stop>
						<stop stop-color="#3b82f6"></stop>
						<stop stop-color="#0f1113" offset="1"></stop>
						<animateTransform repeatCount="indefinite" dur="14s" values="0 380 75;-270 380 75;-270 380 75;-540 380 75;-540 380 75;-810 380 75;-810 380 75;-1080 380 75;-1080 380 75" type="rotate" attributeName="gradientTransform"></animateTransform>
					</linearGradient>
					<linearGradient gradientUnits="userSpaceOnUse" x1="-260" x2="-40" y1="0" y2="0" id="natural-fill">
						<stop offset="0" stop-color="#ffffff" stop-opacity="1"></stop>
						<stop offset="0.48" stop-color="#ffffff" stop-opacity="0.96"></stop>
						<stop offset="0.78" stop-color="#f5f7fb" stop-opacity="0.62"></stop>
						<stop offset="1" stop-color="#f5f7fb" stop-opacity="0"></stop>
						<animate class="natural-fill-motion" attributeName="x1" begin="indefinite" dur="1.2s" from="-260" to="800" fill="freeze"></animate>
						<animate class="natural-fill-motion" attributeName="x2" begin="indefinite" dur="1.2s" from="-40" to="1020" fill="freeze"></animate>
					</linearGradient>
					<linearGradient gradientUnits="userSpaceOnUse" x1="-180" x2="-20" y1="0" y2="0" id="line-reveal-gradient">
						<stop offset="0" stop-color="#ffffff"></stop>
						<stop offset="0.55" stop-color="#ffffff"></stop>
						<stop offset="0.82" stop-color="#ffffff" stop-opacity="0.5"></stop>
						<stop offset="1" stop-color="#000000"></stop>
						<animate class="line-reveal-motion" attributeName="x1" begin="indefinite" dur="3.5s" from="-180" to="${lineRevealEnd.x1}" fill="freeze"></animate>
						<animate class="line-reveal-motion" attributeName="x2" begin="indefinite" dur="3.5s" from="-20" to="${lineRevealEnd.x2}" fill="freeze"></animate>
					</linearGradient>
					<mask id="line-reveal-mask" maskUnits="userSpaceOnUse" x="0" y="0" width="760" height="220">
						<rect x="0" y="0" width="760" height="220" fill="url(#line-reveal-gradient)"></rect>
					</mask>
				</defs>

				<text class="intro-seven intro-seven-behind" x="380" y="110" dominant-baseline="middle" text-anchor="middle">7</text>
				<text class="intro-word" x="380" y="110" dominant-baseline="middle" text-anchor="middle"${lineMask}>Repetierer</text>
				<text class="intro-seven intro-seven-front" x="${frontSevenX}" y="110" dominant-baseline="middle" text-anchor="middle">7</text>
			</svg>`;
	}

	function setStatus(message) {
		status.textContent = message;
	}

	function alignSevenInkCenter(canvas) {
		if (currentVariant !== 2.5 || !canvas) return;
		const word = canvas.querySelector('.intro-word');
		const seven = canvas.querySelector('.intro-seven-behind');
		const wordBounds = word.getBBox();
		const sevenBounds = seven.getBBox();
		const wordCenter = wordBounds.x + (wordBounds.width / 2);
		const sevenCenter = sevenBounds.x + (sevenBounds.width / 2);
		const currentX = Number(seven.getAttribute('x'));
		seven.setAttribute('x', String(currentX + wordCenter - sevenCenter));
	}

	function finish(activeRun) {
		if (activeRun !== runId) return;
		mount.querySelector('.intro-canvas')?.classList.add('is-finished');
		setStatus(`${variantLabels[currentVariant]} – fertig`);
	}

	function playVariant(variant = currentVariant) {
		clearRun();
		currentVariant = Number(variant);
		const activeRun = runId;

		variantButtons.forEach(button => {
			const selected = Number(button.dataset.variant) === currentVariant;
			button.classList.toggle('is-selected', selected);
			button.setAttribute('aria-pressed', String(selected));
		});

		mount.innerHTML = playerMarkup(currentVariant);
		const canvas = mount.querySelector('.intro-canvas');
		alignSevenInkCenter(canvas);
		document.fonts.ready.then(() => {
			if (activeRun === runId) alignSevenInkCenter(canvas);
		});
		setStatus(`${variantLabels[currentVariant]} – wird abgespielt`);

		if (currentVariant === 1 || currentVariant === 2 || currentVariant === 2.5) {
			schedule(() => {
				canvas.classList.add('line-drawing');
				canvas.querySelectorAll('.line-reveal-motion').forEach(animation => animation.beginElement());
			}, timing.lineStart, activeRun);
			schedule(() => canvas.classList.add('line-frozen'), timing.lineComplete, activeRun);
		}
		if (currentVariant === 1 || currentVariant === 2.5) {
			schedule(() => {
				canvas.querySelectorAll('.natural-fill-motion').forEach(animation => animation.beginElement());
			}, timing.lineComplete, activeRun);
		} else if (currentVariant === 3) {
			schedule(() => {
				canvas.querySelectorAll('.natural-fill-motion').forEach(animation => animation.beginElement());
			}, timing.variantThreeFill, activeRun);
		}

		if (currentVariant === 1) {
			schedule(() => canvas.classList.add('word-filled', 'show-compact-seven'), timing.fillComplete, activeRun);
			schedule(() => finish(activeRun), timing.variantOneComplete, activeRun);
		} else if (currentVariant === 2) {
			schedule(() => canvas.classList.add('show-backdrop-seven'), timing.lineComplete, activeRun);
			schedule(() => finish(activeRun), timing.variantTwoComplete, activeRun);
		} else if (currentVariant === 2.5) {
			schedule(() => canvas.classList.add('word-filled', 'show-filled-backdrop-seven'), timing.fillComplete, activeRun);
			schedule(() => finish(activeRun), timing.variantTwoFiveComplete, activeRun);
		} else {
			schedule(() => canvas.classList.add('word-filled'), timing.variantThreeFilled, activeRun);
			schedule(() => canvas.classList.add('fade-word'), timing.variantThreeFade, activeRun);
			schedule(() => canvas.classList.add('show-centered-seven'), timing.variantThreeSeven, activeRun);
			schedule(() => finish(activeRun), timing.variantThreeComplete, activeRun);
		}
	}

	variantButtons.forEach(button => {
		button.addEventListener('click', () => playVariant(button.dataset.variant));
	});
	replayButton.addEventListener('click', () => playVariant());

	window.introTester = {playVariant};
	playVariant(1);
})();
