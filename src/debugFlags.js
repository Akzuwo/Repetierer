// Keep the shipped executable and renderer IPC in agreement. --dev-mode is
// retained for existing shortcuts; --debug is the public simulation flag.
module.exports = { isDebugMode: argv => argv.some(arg => ['--debug', '--debug-mode', '--dev-mode'].includes(arg)) };
