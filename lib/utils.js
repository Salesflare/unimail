const simpleParser = require('mailparser').simpleParser;

exports.parseRawMail = (rawMessage, callback) => {

	return simpleParser(rawMessage.trim(), (err, mail)=>{

		if (err) {
			return callback(err);
		}

		const parsedMail = { headers: {} };

		if (mail && mail.headers) {

			const keys = [...mail.headers.keys()];

			keys.forEach(key => {

				parsedMail.headers[key] = mail.headers.get(key);
			})

		}

		parsedMail.textBody = {
			content: mail.text,
			type: 'text/plain'
		};
		parsedMail.messageId = mail.messageId;

		return callback(null, parsedMail);
	});
};