function createSmtpMailer(config, nodemailerModule) {
	const nodemailer = nodemailerModule || require('nodemailer');
	const transporter = nodemailer.createTransport({
		host: config.host,
		port: config.port,
		secure: config.secure,
		auth: { user: config.user, pass: config.password },
		connectionTimeout: 15000,
		greetingTimeout: 15000,
		socketTimeout: 30000,
		disableFileAccess: true,
		disableUrlAccess: true,
		tls: { minVersion: 'TLSv1.2' }
	});
	return {
		verify: () => transporter.verify(),
		send: message => transporter.sendMail(message)
	};
}

module.exports = { createSmtpMailer };
