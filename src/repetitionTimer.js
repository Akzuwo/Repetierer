const {randomUUID} = require('node:crypto');
const {performance} = require('node:perf_hooks');

// The renderer only announces that the name has landed. All elapsed time is
// measured in the main process, using a monotonic clock, before Excel I/O.
function createRepetitionTimer(now = () => performance.now()) {
	let token = null, started = null, stopped = null;
	return {
		select() {
			token = randomUUID();
			started = stopped = null;
			return token;
		},
		get token() { return token; },
		ready(candidate) {
			if (!token || candidate !== token) return false;
			if (started === null) started = now();
			return true;
		},
		capture() {
			if (started === null) return undefined;
			if (stopped === null) stopped = now();
			const elapsed = stopped - started;
			return Number.isFinite(elapsed) && elapsed >= 0 ? Math.round(elapsed) : undefined;
		},
		cancel() { token = null; started = stopped = null; }
	};
}
module.exports = {createRepetitionTimer};
