// Keep the shipped executable and renderer IPC in agreement. --dev-mode is
// retained for existing shortcuts. Use --debug-mode for packaged executables:
// Electron intercepts a bare --debug as an obsolete Node option before startup.
module.exports = { isDebugMode: argv => argv.some(arg => ['--debug', '--debug-mode', '--dev-mode'].includes(arg)) };
