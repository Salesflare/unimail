'use strict';

const MailComposer = require('nodemailer/lib/mail-composer');
const SimpleParser = require('mailparser').simpleParser;

const internals = {};

/**
 * @param {Object} mailOptions - for detailed properties see https://nodemailer.com/extras/mailcomposer/.
 * @param {{base64Encoded: Boolean}} generateOptions
 * @param {function(Error, String)} callback
 */
exports.generateMessage = (mailOptions, generateOptions, callback) => {

    const mail = new MailComposer({ keepBcc: true, ...mailOptions });

    return mail.compile().build((err, message) => {

        if (err) {
            return callback(err);
        }

        if (generateOptions.base64Encoded) {
            return callback(null, internals.encodeBase64UrlSafe(message));
        }

        return callback(null, message.toString());
    });
};

exports.parseRawMail = (rawMessage, callback) => {

    return SimpleParser(rawMessage.trim(), (err, mail) => {

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

/**
 * return an encoded Buffer as URL Safe Base64
 *
 * Note: This function encodes to the RFC 4648 Spec where '+' is encoded
 *       as '-' and '/' is encoded as '_'. The padding character '=' is
 *       removed.
 *
 * @param {Buffer} buffer
 * @returns {String}
 */
internals.encodeBase64UrlSafe = (buffer) => {

    return buffer.toString('base64')
        .replace(/\+/g, '-') // Convert '+' to '-'
        .replace(/\//g, '_') // Convert '/' to '_'
        .replace(/=+$/, ''); // Remove ending '='
};