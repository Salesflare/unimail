const simpleParser = require('mailparser').simpleParser;

exports.parseRawMail = (rawMessage, callback) => {

    return simpleParser(rawMessage.trim(), (err, mail) => {

        if (err) {
            return callback(err);
        }

        const parsedMail = {
            headers: {},
            date: mail.date,
            to: mail.to,
            from: mail.from,
            cc: mail.cc,
            bcc: mail.bcc
        };

        if (mail && mail.headers) {
            const keys = [...mail.headers.keys()];

            keys.forEach((key) => {

                const headerMapValue = mail.headers.get(key);
                const value = Array.isArray(headerMapValue) ? headerMapValue : [headerMapValue];

                parsedMail.headers[key] = value.map((v) => {

                    if (typeof v === 'object' && v.text) {
                        return v.text;
                    }

                    return v;
                });
            });
        }

        parsedMail.textBody = {
            content: mail.text,
            type: 'text/plain'
        };
        parsedMail.messageId = mail.messageId;

        return callback(null, parsedMail);
    });
};