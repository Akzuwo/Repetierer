export interface Env {
	REPORT_BUCKET: R2Bucket;
	REPORT_API_KEY?: string;
}

type ReportPayload = {
	appVersion?: unknown;
	windowsUser?: unknown;
	comment?: unknown;
	log?: unknown;
	clientTimestamp?: unknown;
};

const MAX_BODY_BYTES = 260_000;
const MAX_APP_VERSION_LENGTH = 50;
const MAX_WINDOWS_USER_LENGTH = 100;
const MAX_COMMENT_LENGTH = 4_000;
const MAX_LOG_LENGTH = 200_000;
const ALLOWED_ORIGINS = new Set([
	'http://localhost:8787',
	'http://127.0.0.1:8787'
]);

export default {
	async fetch(request: Request, env: Env): Promise<Response> {
		try {
			return await handleRequest(request, env);
		} catch (error) {
			console.error('Report handling failed:', getErrorMessage(error));
			return jsonResponse(request, { ok: false, error: 'internal_error' }, 500);
		}
	}
} satisfies ExportedHandler<Env>;

async function handleRequest(request: Request, env: Env): Promise<Response> {
	const url = new URL(request.url);

	if (request.method === 'OPTIONS') {
		return handleOptions(request, url);
	}

	if (url.pathname !== '/report') {
		return jsonResponse(request, { ok: false, error: 'not_found' }, 404);
	}

	if (request.method !== 'POST') {
		return jsonResponse(request, { ok: false, error: 'method_not_allowed' }, 405, {
			Allow: 'POST'
		});
	}

	if (!isAuthorized(request, env)) {
		return jsonResponse(request, { ok: false, error: 'unauthorized' }, 401);
	}

	if (!isJsonContentType(request.headers.get('content-type'))) {
		return jsonResponse(request, { ok: false, error: 'content_type_must_be_application_json' }, 415);
	}

	const contentLength = request.headers.get('content-length');
	if (contentLength && Number(contentLength) > MAX_BODY_BYTES) {
		return jsonResponse(request, { ok: false, error: 'body_too_large' }, 413);
	}

	const bodyText = await readLimitedBody(request, MAX_BODY_BYTES);
	if (bodyText.tooLarge) {
		return jsonResponse(request, { ok: false, error: 'body_too_large' }, 413);
	}

	let payload: ReportPayload;
	try {
		payload = JSON.parse(bodyText.text) as ReportPayload;
	} catch (error) {
		return jsonResponse(request, { ok: false, error: 'invalid_json' }, 400);
	}

	if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
		return jsonResponse(request, { ok: false, error: 'json_object_required' }, 400);
	}

	const validation = normalizePayload(payload);
	if (!validation.ok) {
		return jsonResponse(request, { ok: false, error: validation.error }, 400);
	}

	const id = createReportId();
	const storedAs = `reports/${id}.json`;
	const report = {
		id,
		createdAt: new Date().toISOString(),
		...validation.value,
		meta: {
			userAgent: request.headers.get('user-agent') || '',
			country: request.cf?.country || ''
		}
	};

	await env.REPORT_BUCKET.put(storedAs, JSON.stringify(report, null, 2), {
		httpMetadata: {
			contentType: 'application/json; charset=utf-8'
		}
	});

	return jsonResponse(request, { ok: true, id, storedAs }, 201);
}

function normalizePayload(payload: ReportPayload):
	| { ok: true; value: { appVersion: string; windowsUser: string; comment: string; log: string; clientTimestamp: string } }
	| { ok: false; error: string } {
	const appVersion = normalizeString(payload.appVersion);
	const windowsUser = normalizeString(payload.windowsUser);
	const comment = normalizeString(payload.comment);
	const log = normalizeString(payload.log);
	const clientTimestamp = normalizeString(payload.clientTimestamp);

	if (appVersion.length > MAX_APP_VERSION_LENGTH) return { ok: false, error: 'app_version_too_long' };
	if (windowsUser.length > MAX_WINDOWS_USER_LENGTH) return { ok: false, error: 'windows_user_too_long' };
	if (comment.length > MAX_COMMENT_LENGTH) return { ok: false, error: 'comment_too_long' };
	if (log.length > MAX_LOG_LENGTH) return { ok: false, error: 'log_too_long' };

	return {
		ok: true,
		value: {
			appVersion,
			windowsUser,
			comment,
			log,
			clientTimestamp
		}
	};
}

function normalizeString(value: unknown): string {
	return typeof value === 'string' ? value : '';
}

function isAuthorized(request: Request, env: Env): boolean {
	if (!env.REPORT_API_KEY) return true;
	return request.headers.get('authorization') === `Bearer ${env.REPORT_API_KEY}`;
}

function isJsonContentType(contentType: string | null): boolean {
	if (!contentType) return false;
	return contentType.toLowerCase().split(';')[0].trim() === 'application/json';
}

async function readLimitedBody(request: Request, maxBytes: number): Promise<{ text: string; tooLarge: boolean }> {
	if (!request.body) return { text: '', tooLarge: false };

	const reader = request.body.getReader();
	const chunks: Uint8Array[] = [];
	let receivedBytes = 0;

	while (true) {
		const { done, value } = await reader.read();
		if (done) break;
		if (!value) continue;

		receivedBytes += value.byteLength;
		if (receivedBytes > maxBytes) {
			await reader.cancel();
			return { text: '', tooLarge: true };
		}
		chunks.push(value);
	}

	const body = new Uint8Array(receivedBytes);
	let offset = 0;
	for (const chunk of chunks) {
		body.set(chunk, offset);
		offset += chunk.byteLength;
	}

	return { text: new TextDecoder().decode(body), tooLarge: false };
}

function createReportId(): string {
	const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
	const randomBytes = new Uint8Array(8);
	crypto.getRandomValues(randomBytes);
	const random = Array.from(randomBytes, byte => byte.toString(16).padStart(2, '0')).join('');
	return `${timestamp}-${random}`;
}

function handleOptions(request: Request, url: URL): Response {
	if (url.pathname !== '/report') {
		return new Response(null, { status: 404 });
	}

	return new Response(null, {
		status: 204,
		headers: corsHeaders(request, {
			'Access-Control-Allow-Methods': 'POST, OPTIONS',
			'Access-Control-Allow-Headers': 'Content-Type, Authorization',
			'Access-Control-Max-Age': '86400'
		})
	});
}

function jsonResponse(request: Request, payload: unknown, status = 200, headers: HeadersInit = {}): Response {
	return new Response(JSON.stringify(payload), {
		status,
		headers: corsHeaders(request, {
			'Content-Type': 'application/json; charset=utf-8',
			...headers
		})
	});
}

function corsHeaders(request: Request, headers: HeadersInit = {}): Headers {
	const responseHeaders = new Headers(headers);
	const origin = request.headers.get('origin');

	if (origin && isAllowedOrigin(origin)) {
		responseHeaders.set('Access-Control-Allow-Origin', origin);
		responseHeaders.set('Vary', 'Origin');
	}

	return responseHeaders;
}

function isAllowedOrigin(origin: string): boolean {
	if (origin === 'null') return true;
	if (ALLOWED_ORIGINS.has(origin)) return true;

	try {
		const url = new URL(origin);
		return url.protocol === 'app:' || url.protocol === 'capacitor:' || url.protocol === 'file:';
	} catch (error) {
		return false;
	}
}

function getErrorMessage(error: unknown): string {
	if (error instanceof Error) return error.message;
	return String(error);
}
