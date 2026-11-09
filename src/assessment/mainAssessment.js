const { app, BrowserWindow, dialog, ipcMain, net, safeStorage } = require('electron');
const { AssessmentApiClient } = require('./assessmentApiClient.js');
const { protectTeacherToken, revealTeacherToken } = require('./teacherCredentials.js');
const { AssessmentStore, hashPath } = require('./assessmentStore.js');
const {validateDeadline} = require('./deadlinePolicy');
const {createAutoSync} = require('./autoSync');
let previousAutoSync;
const { deliverRound, sanitizeMailError, validateRoundForDelivery } = require('./assessmentDelivery.js');
const {
	buildRoundFormUrl,
	importIntegrationEnv,
	exportIntegrationFile,
	readIntegrationValues,
	publicAssessmentConfig,
	publicBrevoConfig,
	readAssessmentConfig,
	readBrevoConfig,
	removeConfigGroup,
	cleanIntegrationConfig
} = require('./integrationConfig.js');
const { buildInvitation, buildResultMessage } = require('./mailTemplates.js');
const { createSmtpMailer } = require('./smtpMailer.js');

function registerAssessmentHandlers({ getProgram, getPaths }) {
	const paths = getPaths();
	cleanIntegrationConfig(paths.integrationEnvPath, paths.brevoEnvPath);
	let storeInstance;
	const store = new Proxy({}, {
		get(target, property) {
			if (!storeInstance) storeInstance = new AssessmentStore(paths.assessmentPath);
			const value = storeInstance[property];
			return typeof value === 'function' ? value.bind(storeInstance) : value;
		}
	});
	previousAutoSync?.stop();
	const broadcast = (channel,payload) => {for(const win of BrowserWindow.getAllWindows()) if(!win.isDestroyed()) win.webContents.send(channel,payload);};
	const autoSync = createAutoSync({
		getTargets: () => {
			if (!getAssessmentStatus(paths).configured) return [];
			const filePath = getProgram().getCurrentFilePath();
			if (!filePath) return [];
			return Object.values(store.data.workbooks[hashPath(filePath)]?.classes || {}).flatMap(record => record.rounds.filter(round=>['open','closed'].includes(round.cloudStatus)).map(round=>({filePath,className:record.name,roundId:round.id})));
		},
		sync: target => {
			const context=getContext(getProgram,target);
			if(context.filePath!==target.filePath)return null;
			return syncRound(context,target.roundId);
		},
		onResult: (target,view) => {if(view)broadcast('assessment:auto-synced',{...target,view});},
		onError: (target,error) => {if(target)broadcast('assessment:auto-sync-failed',{...target,message:error.message});}
	});
	previousAutoSync=autoSync;
	app.once('before-quit',()=>autoSync.stop());
	const locked = (context,roundId,action) => autoSync.run({...context,roundId},action);
	async function syncRound(context,roundId) {
		const {round}=store.getRound(context.filePath,context.className,context.students,roundId);
		if (!round.cloudStatus) throw contextError('Diese archivierte Runde bleibt lokal verfügbar. Erstelle für neue Abgaben eine neue Runde.');
		try {
			await publishRound(context,store,round,paths);
			const client=createAssessmentApiClient(await requireAssessmentConfig(paths));
			const result=await client.fetchResponses({externalRoundId:round.externalRoundId,criteria:round.criteria,knownResponseIds:round.externalResponses.map(r=>r.externalResponseId)});
			store.applyResponses(context.filePath,context.className,context.students,round.id,result.responses,{invalid:result.invalidResponses.length,ignoredRound:result.ignoredRoundCount});
			store.setCloudStatus(context.filePath,context.className,context.students,round.id,result.closed ? 'closed' : 'open',client.getFormUrl());
		} catch(error) {store.markSyncFailed(context.filePath,context.className,context.students,round.id);throw error;}
		return presentRound(context,store,round.id,paths);
	}

	register('assessment:get-overview', async (event, args) => {
		const context = getContext(getProgram, args);
		return presentClass(store.getClassContext(context.filePath, context.className, context.students));
	});
	register('assessment:create-round', async (event, args) => {
		const context = getContext(getProgram, args);
		validateDeadline(args.deadline);
		await requireAssessmentConfig(paths);
		const round = store.createRound(context.filePath, context.className, context.students, args);
		await publishRound(context, store, round, paths);
		return presentClass(store.getClassContext(context.filePath, context.className, context.students));
	});
	register('assessment:get-round', async (event, args) => {
		const context = getContext(getProgram, args);
		return presentRound(context, store, args.roundId, paths);
	});
	register('assessment:sync-responses', async (event, args) => {
		const context = getContext(getProgram, args);
		return locked(context,args.roundId,()=>syncRound(context,args.roundId));
	});
	register('assessment:assign-response', async (event, args) => {
		const context = getContext(getProgram, args);
		store.assignResponse(context.filePath, context.className, context.students, args.roundId, args.responseId, args.studentId);
		return presentRound(context, store, args.roundId, paths);
	});
	register('assessment:select-response', async (event, args) => {
		const context = getContext(getProgram, args);
		store.selectResponse(context.filePath, context.className, context.students, args.roundId, args.responseId);
		return presentRound(context, store, args.roundId, paths);
	});
	register('assessment:save-teacher', async (event, args) => {
		const context = getContext(getProgram, args);
		store.saveTeacherAssessment(context.filePath, context.className, context.students, args.roundId, args.studentId, args.answers, args.comment, args.criterionComments);
		return presentRound(context, store, args.roundId, paths);
	});
	register('assessment:delete-round', async (event, args) => {
		const context = getContext(getProgram, args);
		return locked(context,args.roundId,async()=>{
			const { round } = store.getRound(context.filePath, context.className, context.students, args.roundId);
			if (round.cloudStatus) await createAssessmentApiClient(await requireAssessmentConfig(paths)).deleteRound(round.externalRoundId);
			store.deleteRound(context.filePath, context.className, context.students, args.roundId);
			return presentClass(store.getClassContext(context.filePath, context.className, context.students));
		});
	});
	register('assessment:close-round', async (event, args) => {
		const context = getContext(getProgram, args);
		return locked(context,args.roundId,async()=>{
			const { round } = store.getRound(context.filePath, context.className, context.students, args.roundId);
			if (!round.cloudStatus) throw contextError('Diese Runde ist bereits lokal archiviert.');
			await publishRound(context, store, round, paths);
			await createAssessmentApiClient(await requireAssessmentConfig(paths)).closeRound(round.externalRoundId);
			store.setCloudStatus(context.filePath, context.className, context.students, round.id, 'closed');
			return presentRound(context, store, round.id, paths);
		});
	});
	register('assessment:reopen-round', async (event,args) => {
		const context=getContext(getProgram,args);
		const deadline=validateDeadline(args.deadline,Date.now(),true);
		return locked(context,args.roundId,async()=>{
			const {round}=store.getRound(context.filePath,context.className,context.students,args.roundId);
			if(!round.cloudStatus)throw contextError('Archivierte Runden können nicht wieder geöffnet werden.');
			const client=createAssessmentApiClient(await requireAssessmentConfig(paths));
			if(round.cloudStatus==='pending') {
				// A failed creation remains recoverable under the same local title.
				await client.createRound({...round,deadline},round.cloudMaxResponses);
			} else await client.reopenRound(round.externalRoundId,deadline);
			store.setDeadline(context.filePath,context.className,context.students,round.id,deadline);
			store.setCloudStatus(context.filePath,context.className,context.students,round.id,'open',client.getFormUrl());
			return presentRound(context,store,round.id,paths);
		});
	});

	register('assessment:config-status', async () => getAssessmentStatus(paths));
	register('assessment:config-import', async () => {
		const selected = dialog.showOpenDialogSync(BrowserWindow.getFocusedWindow(), {
			properties: ['openFile'],
			filters: [{ name: 'Repetierer-Paket', extensions: ['repetierer'] }, { name: 'Bisherige Konfiguration', extensions: ['env'] }]
		});
		if (!selected) return { cancelled: true };
		const result = await importIntegrationEnv(selected[0], paths.integrationEnvPath, {
			legacyBrevoPath: paths.brevoEnvPath,
			protectTeacherToken: token => protectTeacherToken(token, safeStorage),
			verify: async ({ assessment }) => {
				if (assessment) {
					if (assessment.teacherToken.startsWith('encrypted:')) assessment.teacherToken = revealTeacherToken(assessment.teacherToken, safeStorage);
					await createAssessmentApiClient(assessment).checkStatus();
				}
			}
		});
		return result;
	});
	register('assessment:config-export', async () => {
		const destination = dialog.showSaveDialogSync(BrowserWindow.getFocusedWindow(), {
			title: 'SMTP und Cloudflare gemeinsam exportieren',
			defaultPath: 'Repetierer-Konfiguration.repetierer',
			filters: [{name:'Repetierer-Paket',extensions:['repetierer']}]
		});
		if (!destination) return {cancelled:true};
		const filePath = destination.toLowerCase().endsWith('.repetierer') ? destination : destination + '.repetierer';
		return exportIntegrationFile(filePath, readIntegrationValues(paths.integrationEnvPath, paths.brevoEnvPath), {
			revealTeacherToken: token => revealTeacherToken(token, safeStorage)
		});
	});
	register('assessment:api-test', async () => {
		const config = await requireAssessmentConfig(paths);
		return createAssessmentApiClient(config).checkStatus();
	});
	register('assessment:config-remove', async () => {
		removeConfigGroup(paths.integrationEnvPath, paths.brevoEnvPath, 'assessment');
		return { configured: false };
	});

	register('assessment:brevo-status', async () => getBrevoStatus(paths));
	register('assessment:brevo-test', async () => {
		const config = requireBrevo(paths);
		try {
			await createSmtpMailer(config).verify();
			return { ok: true };
		} catch (error) {
			return { ok: false, error: sanitizeMailError(error) };
		}
	});
	register('assessment:brevo-remove', async () => {
		removeConfigGroup(paths.integrationEnvPath, paths.brevoEnvPath, 'brevo');
		return { configured: false };
	});

	register('assessment:send-round', async (event, args) => {
		const config = requireBrevo(paths);
		const context = getContext(getProgram, args);
		const current = store.getRound(context.filePath, context.className, context.students, args.roundId);
		const mailer = createSmtpMailer(config);
		const result = await deliverRound({
			classRecord: current.classRecord,
			round: current.round,
			mode: args.mode === 'failed_only' ? 'failed_only' : 'pending',
			buildMessage: (student, assessment) => buildResultMessage({ student, assessment, round: current.round }),
			send: message => mailer.send(message),
			updateStatus: (studentId, status, details) => store.setMailStatus(context.filePath, context.className, context.students, current.round.id, studentId, status, details)
		});
		return { result, round: await presentRound(context, store, current.round.id, paths) };
	});
}

function register(channel, handler) {
	ipcMain.removeHandler(channel);
	ipcMain.handle(channel, async (event, args) => {
		try {
			return { ok: true, data: await handler(event, args || {}) };
		} catch (error) {
			return { ok: false, error: {
				code: error && error.code || 'ASSESSMENT_ERROR',
				message: error && error.message || 'Die Aktion konnte nicht abgeschlossen werden.',
				incomplete: error && error.incomplete || undefined
			} };
		}
	});
}

function getContext(getProgram, args) {
	const program = getProgram();
	const filePath = program.getCurrentFilePath();
	const className = String(args && args.className || '').trim();
	if (!filePath) throw contextError('Bitte zuerst eine Repetierer-Datei laden.');
	if (!className) throw contextError('Bitte eine Klasse auswählen.');
	const students = program.getClassPersons(className);
	if (!students.length) throw contextError('In dieser Klasse wurden keine Schüler gefunden.');
	return { filePath, className, students };
}

function presentClass(classRecord) {
	return {
		id: classRecord.id,
		name: classRecord.name,
		students: classRecord.students.filter(student => student.active !== false),
		rounds: classRecord.rounds.map(round => summarizeRound(round, classRecord.students))
	};
}

async function presentRound(context, store, roundId, paths) {
	const { classRecord, round } = store.getRound(context.filePath, context.className, context.students, roundId);
	let invitation = null;
	let assessmentConfigured = false;
	let invitationError = '';
	try {
		if (round.cloudStatus === 'pending') await publishRound(context, store, round, paths);
		if (!round.cloudStatus) throw contextError('Archivierte Runde: Bewertungen und synchronisierte Antworten bleiben lokal erhalten.');
		const config = await requireAssessmentConfig(paths);
		assessmentConfigured = !!config;
		if (config) invitation = buildInvitation(round, buildRoundFormUrl({...config, formUrl: round.cloudFormUrl || createAssessmentApiClient(config).getFormUrl()}, round.externalRoundId));
	} catch (error) {
		try { assessmentConfigured = !!readAssessmentConfig(paths.integrationEnvPath, paths.brevoEnvPath); }
		catch (configError) { assessmentConfigured = false; }
		invitationError = error.message;
	}
	return {
		filePath: context.filePath,
		classRecord: presentClass(classRecord),
		round,
		students: classRecord.students.filter(student => student.active !== false),
		summary: summarizeRound(round, classRecord.students),
		invitation,
		invitationError,
		assessmentConfigured,
		deliveryValidation: validateRoundForDelivery(classRecord, round)
	};
}

function summarizeRound(round, students) {
	const active = students.filter(student => student.active !== false);
	const assessments = active.map(student => round.assessments[student.id]).filter(Boolean);
	return {
		id: round.id,
		cloudStatus: round.cloudStatus || 'archived',
		externalRoundId: round.externalRoundId,
		title: round.title,
		deadline: round.deadline,
		createdAt: round.createdAt,
		studentCount: active.length,
		selfCount: assessments.filter(item => item.studentAnswers).length,
		teacherCount: assessments.filter(item => item.teacherAnswers).length,
		sentCount: assessments.filter(item => item.mailStatus === 'sent').length,
		failedCount: assessments.filter(item => item.mailStatus === 'failed').length,
		unmatchedCount: (round.unmatchedResponses || []).length,
		duplicateCount: (round.duplicates || []).reduce((sum, duplicate) => sum + duplicate.count - 1, 0),
		lastSyncAt: round.lastSyncAt,
		lastSyncStatus: round.lastSyncStatus || 'never',
		lastSyncResult: round.lastSyncResult || null
	};
}

function getAssessmentStatus(paths) {
	try { return publicAssessmentConfig(readAssessmentConfig(paths.integrationEnvPath, paths.brevoEnvPath)); }
	catch (error) { return { configured: false, invalid: true, error: error.message }; }
}

function getBrevoStatus(paths) {
	try { return publicBrevoConfig(readBrevoConfig(paths.integrationEnvPath, paths.brevoEnvPath)); }
	catch (error) { return { configured: false, invalid: true, error: error.message }; }
}

async function requireAssessmentConfig(paths) {
	const config = readAssessmentConfig(paths.integrationEnvPath, paths.brevoEnvPath);
	if (!config) throw contextError('Die Mitmachnoten-Konfiguration ist noch nicht eingerichtet.');
	config.teacherToken = revealTeacherToken(config.teacherToken, safeStorage);
	return config;
}

async function publishRound(context, store, round, paths) {
	if (round.cloudStatus !== 'pending') return;
	validateDeadline(round.deadline);
	const client = createAssessmentApiClient(await requireAssessmentConfig(paths));
	await client.createRound(round, round.cloudMaxResponses);
	store.setCloudStatus(context.filePath, context.className, context.students, round.id, 'open', client.getFormUrl());
	round.cloudStatus = 'open';
	round.cloudFormUrl = client.getFormUrl();
}

function createAssessmentApiClient(config) {
	return new AssessmentApiClient(config, { fetch: net.fetch });
}

function requireBrevo(paths) {
	const config = readBrevoConfig(paths.integrationEnvPath, paths.brevoEnvPath);
	if (!config) {
		const error = new Error('Brevo SMTP ist noch nicht eingerichtet.');
		error.code = 'BREVO_NOT_CONFIGURED';
		throw error;
	}
	return config;
}

function contextError(message) {
	const error = new Error(message);
	error.code = 'ASSESSMENT_CONTEXT_MISSING';
	return error;
}

module.exports = { presentClass, registerAssessmentHandlers, summarizeRound };
