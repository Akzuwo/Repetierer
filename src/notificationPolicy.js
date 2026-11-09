const DEFAULT_DURATION_SECONDS = 6;
const MAX_VISIBLE = 3;
const TITLES = Object.freeze({error:'Fehler', warning:'Achtung', info:'Hinweis', success:'Erfolgreich'});

function normalizeDuration(value) {
	if (value === null || value === undefined || value === '' || !Number.isFinite(Number(value))) return DEFAULT_DURATION_SECONDS;
	return Math.max(2,Math.min(8,Math.round(Number(value))));
}
function normalizeNotification(type, text, options = {}, defaultDuration = DEFAULT_DURATION_SECONDS) {
	const severity = type === 'warn' ? 'warning' : Object.hasOwn(TITLES,type) ? type : 'info';
	const message = String(text instanceof Error ? text.message : text ?? '').trim().slice(0,1600);
	if (!message) return null;
	return {
		type:severity, text:message,
		title:String(options.title || TITLES[severity]).trim().slice(0,120),
		durationSeconds:normalizeDuration(options.durationSeconds ?? defaultDuration),
		key:typeof options.key === 'string' ? options.key.slice(0,120) : undefined
	};
}
function inferType(title = '', detail = '') {
	const text = `${title} ${detail}`;
	if (/fehler|fehlgeschlagen|ungültig|nicht (möglich|gespeichert|geladen|gestartet|beendet)|konnte.*nicht|nicht.*durchgeführt/i.test(text)) return 'error';
	if (/warnung|ausstehend|erforderlich|gesperrt|geöffnet|teilweise|noch offen|bitte.*schliess|keine (datei|klasse|person)|wähle.*(zuerst|klasse)|nicht eingerichtet/i.test(text)) return 'warning';
	if (/gespeichert|erfolgreich|exportiert|gesendet|aktualisiert|übernommen|abgeschlossen|beendet|kopiert|importiert|rückgängig gemacht|wiederhergestellt|angepasst|neu geladen/i.test(text)) return 'success';
	return 'info';
}
module.exports = {DEFAULT_DURATION_SECONDS,MAX_VISIBLE,TITLES,normalizeDuration,normalizeNotification,inferType};
