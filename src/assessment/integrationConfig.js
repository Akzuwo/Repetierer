const fs = require('fs');
const path = require('path');
const { isValidEmail } = require('./matching.js');

const ASSESSMENT_KEYS = ['ASSESSMENT_API_URL', 'ASSESSMENT_TEACHER_TOKEN'];
const BREVO_KEYS = [
	'BREVO_SMTP_HOST',
	'BREVO_SMTP_PORT',
	'BREVO_SMTP_USER',
	'BREVO_SMTP_PASSWORD',
	'BREVO_FROM_EMAIL',
	'BREVO_FROM_NAME'
];
const KNOWN_KEYS = [...ASSESSMENT_KEYS, 'ASSESSMENT_FALLBACK_API_URL', ...BREVO_KEYS];

function parseEnv(content) {
	const values = {};
	const lines = String(content || '').replace(/^\uFEFF/, '').split(/\r?\n/);
	for (let index = 0; index < lines.length; index++) {
		const trimmed = lines[index].trim();
		if (!trimmed || trimmed.startsWith('#')) continue;
		const match = trimmed.match(/^(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/);
		if (!match) throw configError(`Ungültige Zeile ${index + 1} in der .env-Datei.`);
		let value = match[2].trim();
		if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
			const quote = value[0];
			value = value.slice(1, -1);
			if (quote === '"') value = value.replace(/\\n/g, '\n').replace(/\\"/g, '"').replace(/\\\\/g, '\\');
		} else {
			value = value.replace(/\s+#.*$/, '').trim();
		}
		values[match[1]] = value;
	}
	return values;
}

function validateAssessmentConfig(values) {
	assertCompleteGroup(values, ASSESSMENT_KEYS, 'Mitmachnote');
	const apiUrl = validateHttpsUrl(values.ASSESSMENT_API_URL, 'ASSESSMENT_API_URL');
	const url = new URL(apiUrl);
	if (url.username || url.password || url.search || url.hash || !url.pathname.endsWith('/api')) throw configError('ASSESSMENT_API_URL muss auf /api enden und darf keine Zugangsdaten enthalten.');
	const teacherToken = String(values.ASSESSMENT_TEACHER_TOKEN).trim();
	if (!/^rt_[A-Za-z0-9_-]{43}$/.test(teacherToken) && !/^encrypted:[A-Za-z0-9+/=]+$/.test(teacherToken)) throw configError('ASSESSMENT_TEACHER_TOKEN ist syntaktisch ungültig.');
	const fallbackApiUrl = values.ASSESSMENT_FALLBACK_API_URL
		? validateHttpsUrl(values.ASSESSMENT_FALLBACK_API_URL, 'ASSESSMENT_FALLBACK_API_URL')
		: apiUrl === 'https://timolu.ch/app/services/repetierer/api' ? 'https://akzuwo.ch/app/services/repetierer/api' : null;
	if (fallbackApiUrl) {
		const fallback = new URL(fallbackApiUrl);
		if (fallback.username || fallback.password || fallback.search || fallback.hash || !fallback.pathname.endsWith('/api')) throw configError('Ungültige Fallback-API-Adresse.');
	}
	return { formUrl: apiUrl.slice(0, -3), apiUrl, fallbackApiUrl, teacherToken };
}

function validateBrevoConfig(values) {
	assertCompleteGroup(values, BREVO_KEYS, 'Brevo SMTP');
	const port = Number(values.BREVO_SMTP_PORT);
	if (!Number.isInteger(port) || port < 1 || port > 65535) throw configError('BREVO_SMTP_PORT ist ungültig.');
	if (!isValidEmail(values.BREVO_FROM_EMAIL)) throw configError('BREVO_FROM_EMAIL ist ungültig.');
	if (!/^[A-Za-z0-9.-]+$/.test(values.BREVO_SMTP_HOST) || !values.BREVO_SMTP_HOST.includes('.')) {
		throw configError('BREVO_SMTP_HOST ist syntaktisch ungültig.');
	}
	if (/[\r\n]/.test(values.BREVO_SMTP_USER) || /[\r\n]/.test(values.BREVO_FROM_NAME)) {
		throw configError('SMTP-Benutzer und Absendername dürfen keine Zeilenumbrüche enthalten.');
	}
	return {
		host: values.BREVO_SMTP_HOST,
		port,
		secure: port === 465,
		user: values.BREVO_SMTP_USER,
		password: values.BREVO_SMTP_PASSWORD,
		fromEmail: values.BREVO_FROM_EMAIL,
		fromName: values.BREVO_FROM_NAME
	};
}

function readIntegrationValues(integrationPath, legacyBrevoPath) {
	const values = {};
	if (legacyBrevoPath && fs.existsSync(legacyBrevoPath)) Object.assign(values, parseEnv(fs.readFileSync(legacyBrevoPath, 'utf8')));
	if (integrationPath && fs.existsSync(integrationPath)) Object.assign(values, parseEnv(fs.readFileSync(integrationPath, 'utf8')));
	return values;
}

function cleanIntegrationConfig(integrationPath, legacyBrevoPath) {
	if (!integrationPath || !fs.existsSync(integrationPath)) return;
	const previous = readIntegrationValues(integrationPath, legacyBrevoPath);
	const values = pickKnown(previous);
	if (!values.ASSESSMENT_TEACHER_TOKEN) { delete values.ASSESSMENT_API_URL; delete values.ASSESSMENT_FALLBACK_API_URL; }
	if (JSON.stringify(previous) !== JSON.stringify(values)) atomicWrite(integrationPath, serializeValues(values));
}

function readAssessmentConfig(integrationPath, legacyBrevoPath) {
	const values = readIntegrationValues(integrationPath, legacyBrevoPath);
	return hasAny(values, ASSESSMENT_KEYS) ? validateAssessmentConfig(values) : null;
}

function readBrevoConfig(integrationPath, legacyBrevoPath) {
	const values = readIntegrationValues(integrationPath, legacyBrevoPath);
	return hasAny(values, BREVO_KEYS) ? validateBrevoConfig(values) : null;
}

async function importIntegrationEnv(sourcePath, destinationPath, options = {}) {
	const source = path.resolve(sourcePath);
	const destination = path.resolve(destinationPath);
	if (source === destination) throw configError('Bitte wähle eine .env-Datei außerhalb des internen Datenordners.');
	let original;
	try {
		original = fs.readFileSync(source, 'utf8');
	} catch (error) {
		throw configError('Die ausgewählte .env-Datei konnte nicht gelesen werden.');
	}
	if (Buffer.byteLength(original, 'utf8') > 262144) throw configError('Die Konfigurationsdatei ist zu gross.');
	const portable = path.extname(source).toLowerCase() === '.repetierer';
	const imported = portable ? parsePortableConfig(original) : parseEnv(original);
	const hasAssessment = hasAny(imported, ASSESSMENT_KEYS);
	const hasBrevo = hasAny(imported, BREVO_KEYS);
	if (!hasAssessment && !hasBrevo) throw configError('Die .env-Datei enthält keine unterstützte Mitmachnoten- oder Brevo-Konfiguration.');
	if (hasAssessment) validateAssessmentConfig(imported);
	if (hasBrevo) validateBrevoConfig(imported);

	const current = options.currentValues || readIntegrationValues(destination, options.legacyBrevoPath);
	const merged = Object.assign({}, pickKnown(current), pickKnown(imported));
	validateConfiguredGroups(merged);
	if (options.verify) await options.verify({
		assessment: hasAny(merged, ASSESSMENT_KEYS) ? validateAssessmentConfig(merged) : null,
		brevo: hasAny(merged, BREVO_KEYS) ? validateBrevoConfig(merged) : null
	});
	if (hasAssessment && options.protectTeacherToken) merged.ASSESSMENT_TEACHER_TOKEN = options.protectTeacherToken(imported.ASSESSMENT_TEACHER_TOKEN);

	fs.mkdirSync(path.dirname(destination), { recursive: true });
	const temporary = `${destination}.${process.pid}.${Date.now()}.tmp`;
	try {
		fs.writeFileSync(temporary, serializeValues(merged), { encoding: 'utf8', mode: 0o600, flag: 'wx' });
		assertEquivalent(merged, readIntegrationValues(temporary));
		fs.renameSync(temporary, destination);
		assertEquivalent(merged, readIntegrationValues(destination));
	} catch (error) {
		if (fs.existsSync(temporary)) fs.unlinkSync(temporary);
		if (error && error.code === 'INTEGRATION_CONFIG_ERROR') throw error;
		throw configError('Die Konfiguration konnte nicht sicher im Repetierer-Datenordner gespeichert werden.');
	}

	let originalDeleted = false;
	let deleteError = null;
	try {
		if (portable || options.preserveSource) return { configured: publicIntegrationStatus(merged), originalDeleted: false, deleteError: null };
		fs.unlinkSync(source);
		originalDeleted = true;
	} catch (error) {
		deleteError = 'Die interne Kopie wurde gespeichert, aber die ursprüngliche Datei konnte nicht gelöscht werden.';
	}
	return { configured: publicIntegrationStatus(merged), originalDeleted, deleteError };
}

function parsePortableConfig(content) {
	let bundle;
	try { bundle = JSON.parse(String(content).replace(/^\uFEFF/, '')); } catch { throw configError('Die .repetierer-Datei ist kein gültiges Konfigurationspaket.'); }
	if (!bundle || bundle.format !== 'repetierer-settings' || bundle.version !== 1 || !bundle.values || Array.isArray(bundle.values) || typeof bundle.values !== 'object') throw configError('Unbekanntes .repetierer-Format oder nicht unterstützte Version.');
	if (Object.keys(bundle.values).some(key => !KNOWN_KEYS.includes(key) || typeof bundle.values[key] !== 'string')) throw configError('Das Paket enthält unbekannte oder ungültige Einstellungen.');
	const values = pickKnown(bundle.values);
	validateAssessmentConfig(values); validateBrevoConfig(values);
	if (values.ASSESSMENT_TEACHER_TOKEN.startsWith('encrypted:')) throw configError('Dieses Paket enthält einen gerätegebundenen Zugang. Bitte über «Paket exportieren» neu erstellen.');
	return values;
}

function exportIntegrationFile(destinationPath, values, options = {}) {
	const exported = pickKnown(values);
	if (exported.ASSESSMENT_TEACHER_TOKEN?.startsWith('encrypted:')) {
		if (!options.revealTeacherToken) throw configError('Der lokale Lehrerzugang konnte nicht für das Paket geöffnet werden.');
		exported.ASSESSMENT_TEACHER_TOKEN = options.revealTeacherToken(exported.ASSESSMENT_TEACHER_TOKEN);
	}
	// A package always contains both connections and only portable credentials.
	const content = JSON.stringify({format:'repetierer-settings', version:1, values:exported}, null, 2) + '\n';
	parsePortableConfig(content);
	atomicWrite(destinationPath, content);
	return { exported: true, filePath: destinationPath };
}

function removeConfigGroup(integrationPath, legacyBrevoPath, group) {
	const removeKeys = group === 'brevo' ? BREVO_KEYS : [...ASSESSMENT_KEYS, 'ASSESSMENT_FALLBACK_API_URL'];
	const values = readIntegrationValues(integrationPath, legacyBrevoPath);
	for (const key of removeKeys) delete values[key];
	if (group === 'brevo' && legacyBrevoPath && fs.existsSync(legacyBrevoPath)) fs.unlinkSync(legacyBrevoPath);
	if (!Object.keys(pickKnown(values)).length) {
		if (fs.existsSync(integrationPath)) fs.unlinkSync(integrationPath);
		return;
	}
	atomicWrite(integrationPath, serializeValues(values));
}

function publicAssessmentConfig(config) {
	return config ? { configured: true, formUrl: config.formUrl, apiUrl: config.apiUrl, fallbackApiUrl: config.fallbackApiUrl } : { configured: false };
}

function publicBrevoConfig(config) {
	return config ? { configured: true, host: config.host, port: config.port, ...require('./senderIdentity') } : { configured: false };
}

function publicIntegrationStatus(values) {
	return {
		assessment: publicAssessmentConfig(hasAny(values, ASSESSMENT_KEYS) ? validateAssessmentConfig(values) : null),
		brevo: publicBrevoConfig(hasAny(values, BREVO_KEYS) ? validateBrevoConfig(values) : null)
	};
}

function buildRoundFormUrl(config, externalRoundId) {
	const url = new URL(config.formUrl);
	url.searchParams.set('round', externalRoundId);
	return url.toString();
}

function validateConfiguredGroups(values) {
	if (hasAny(values, ASSESSMENT_KEYS)) validateAssessmentConfig(values);
	if (hasAny(values, BREVO_KEYS)) validateBrevoConfig(values);
}

function assertCompleteGroup(values, keys, label) {
	const missing = keys.filter(key => !String(values && values[key] || '').trim());
	if (missing.length) throw configError(`${label}: Fehlende oder leere Werte: ${missing.join(', ')}`);
}

function validateHttpsUrl(value, key) {
	let url;
	try { url = new URL(String(value || '').trim()); } catch (error) { throw configError(`${key} ist ungültig.`); }
	if (url.protocol !== 'https:') throw configError(`${key} muss eine HTTPS-Adresse sein.`);
	return url.toString();
}

function hasAny(values, keys) {
	return keys.some(key => Object.hasOwn(values || {}, key));
}

function pickKnown(values) {
	const result = {};
	for (const key of KNOWN_KEYS) if (Object.hasOwn(values || {}, key)) result[key] = String(values[key]);
	return result;
}

function serializeValues(values) {
	return KNOWN_KEYS.filter(key => Object.hasOwn(values, key)).map(key => `${key}=${quoteEnv(values[key])}`).concat('').join('\n');
}

function quoteEnv(value) {
	return `"${String(value).replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\r?\n/g, '\\n')}"`;
}

function assertEquivalent(expected, actual) {
	for (const [key, value] of Object.entries(pickKnown(expected))) {
		if (actual[key] !== value) throw configError('Die gespeicherte Konfiguration konnte nicht verifiziert werden.');
	}
}

function atomicWrite(filePath, content) {
	fs.mkdirSync(path.dirname(filePath), { recursive: true });
	const temporary = `${filePath}.${process.pid}.${Date.now()}.tmp`;
	try {
		fs.writeFileSync(temporary, content, { encoding: 'utf8', mode: 0o600, flag: 'wx' });
		fs.renameSync(temporary, filePath);
	} finally {
		if (fs.existsSync(temporary)) fs.unlinkSync(temporary);
	}
}

function configError(message) {
	const error = new Error(message);
	error.code = 'INTEGRATION_CONFIG_ERROR';
	return error;
}

module.exports = {
	ASSESSMENT_KEYS,
	BREVO_KEYS,
	cleanIntegrationConfig,
	buildRoundFormUrl,
	importIntegrationEnv,
	exportIntegrationFile,
	parsePortableConfig,
	parseEnv,
	publicAssessmentConfig,
	publicBrevoConfig,
	readAssessmentConfig,
	readBrevoConfig,
	readIntegrationValues,
	removeConfigGroup,
	validateAssessmentConfig,
	validateBrevoConfig
};
