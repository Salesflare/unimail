'use strict';

const _ = require('lodash');
const MailComposer = require('nodemailer/lib/mail-composer');
const SimpleParser = require('mailparser').simpleParser;

const internals = {};

/**
 * @param {Object} mailOptions - for detailed properties see https://nodemailer.com/extras/mailcomposer/.
 * @param {{base64Encoded: Boolean}} generateOptions
 * @param {function(Error, String):void} callback
 * @returns {void}
 */
exports.generateMessage = (mailOptions, generateOptions, callback) => {

    // Otherwise allows email local files
    mailOptions.disableFileAccess = true;

    const mail = new MailComposer(mailOptions).compile();

    mail.keepBcc = true;
    return mail.build((err, message) => {

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

    // We pass our own Buffer here for encoding reasons
    // SimpleParser used to transform strings to Buffers with encoding `binary` which is `latin1` but we want `utf8`
    // This breaks for Greek characters for example
    // See https://github.com/nodemailer/mailparser/issues/241
    // This issue is fixed in SimpleParser but to make sure it won't break for us again I'm keeping our own parse step
    return SimpleParser(Buffer.from(rawMessage.trim(), 'utf8'), { maxHeadSize: 10 * 1024 * 1024 }, (err, mail) => {

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

                // We want to make sure the references headers only contains strings, other stuff can lead to problems downstream
                // See https://github.com/Salesflare/Server/issues/7122
                if (key === 'references') {
                    parsedMail.headers[key] = parsedMail.headers[key].filter((reference) => {

                        return _.isString(reference);
                    });
                }
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
 * Return an encoded Buffer as URL Safe Base64
 *
 * Note:
 * This function encodes to the RFC 4648 Spec where '+' is encoded
 * as '-' and '/' is encoded as '_'. The padding character '=' is removed.
 *
 * @param {Buffer} buffer
 * @returns {String}
 */
internals.encodeBase64UrlSafe = (buffer) => {

    return buffer.toString('base64')
        .replace(/\+/g, '-') // Convert '+' to '-'
        .replace(/\//g, '_') // Convert '/' to '_'
        .replace(/\u003D+$/, ''); // Remove ending '='
};
