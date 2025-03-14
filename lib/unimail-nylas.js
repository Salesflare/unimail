'use strict';

const EventEmitter = require('events');
const Async = require('async');

const Nylas = require('@salesflare/nylas');
const Boom = require('@hapi/boom');


const Utils = require('../lib/utils');

const internals = {};

/**
 * @typedef {import('./index').MessageResource} MessageResource
 * @typedef {import('./index').MessageListResource} MessageListResource
 * @typedef {import('./index').FileResource} FileResource
 * @typedef {import('./index').FileListResource} FileListResource
 * @typedef {import('./index').MessageRecipient} MessageRecipient
 */

class NylasConnector extends EventEmitter {

    /**
     * @class
     *
     * @param {Object} config - Configuration object
     * @param {String} config.clientId
     * @param {String} config.clientSecret
     */
    constructor(config) {

        super();

        if (!config || !config.clientId || !config.clientSecret) {
            throw new Error('Invalid configuration. Please refer to the documentation to get the required fields.');
        }

        this.clientId = config.clientId;
        this.clientSecret = config.clientSecret;

        this.name = 'nylas';

        Nylas.config({
            appId: this.clientId,
            appSecret: this.clientSecret
        });
    }

    /* MESSAGES */

    /**
     *
     * @throws
     *
     * @param {Object} auth - Authentication object
     * @param {String} auth.access_token - Access token
     *
     * @param {Object} params
     * @param {String} params.id - Nylas message id
     * @param {String} [params.rfc2822Format=false] - Return the email in rfc2822 format https://www.ietf.org/rfc/rfc2822.txt
     *
     * @param {Object} options
     * @param {Boolean} [options.raw] - If true the response will not be transformed to the unified object
     *
     * @param {function(Error?, (MessageResource | Object | String)?):void} callback Returns a unified message resource when options.raw is falsy or the raw response of the API when truthy
     * @returns {void}
     */
    getMessage(auth, params, options, callback) {

        if (!params || !params.id) {
            throw new Error('Invalid configuration. Please refer to the documentation to get the required fields.');
        }

        if (typeof options === 'function') {
            callback = options;
            options = {};
        }

        options = options || {};

        const nylas = Nylas.with(auth.access_token);

        return nylas.messages.find(params.id).then((message) => {

            if (options.raw) {
                return callback(null, message);
            }

            return message.getRaw().then((raw) => {

                if ((/^Message .* has not been synced yet$/).test(raw)) {
                    return callback(new Error(raw));
                }

                if (params.rfc2822Format) {
                    return callback(null, raw);
                }

                return Utils.parseRawMail(raw, (err, parsedMail) => {

                    if (err) {
                        return callback(err);
                    }

                    message.headers = parsedMail.headers;
                    message.textBody = parsedMail.textBody;
                    message.email_message_id = parsedMail.messageId;

                    if (message.headers['in-reply-to'] && message.headers['in-reply-to'].length > 0) {
                        message.in_reply_to = message.headers['in-reply-to'][0];
                    }
                    else {
                        message.in_reply_to = null;
                    }

                    if (message.files && message.files.length > 0) {
                        return Async.map(message.files, (file, callback) => {

                            const fileObject = {
                                message,
                                metadata: file
                            };

                            if (options.raw) {
                                return callback(null, fileObject);
                            }

                            return callback(null, this._transformFiles(fileObject)[0]);
                        }, (err, files) => {

                            if (err) {
                                return callback(err);
                            }

                            message.files = files.filter((f) => !!f);

                            return callback(null, this._transformMessages(message)[0]);
                        });
                    }

                    return callback(null, this._transformMessages(message)[0]);
                });
            })
                .catch((err) => {

                    return callback(err);
                });
        })
            .catch((err) => {

                let statusCode = 500;

                if (err.message) {
                    if (err.message.includes('Couldn\'t find')) {
                        statusCode = 404;
                    }

                    if (err.message.includes('Too many concurrent query requests')) {
                        statusCode = 429;
                    }
                }

                return callback(Boom.boomify(err, { statusCode }));
            });
    }

    /**
     * @param {Object} auth - Authentication object
     * @param {String} auth.access_token - Access token
     * @param {Object} params
     * @param {Number} params.limit - Maximum amount of messages in response, max = 100
     * @param {Boolean} params.hasAttachment - If true, only return messages with attachments (false doesn't work yet in Nylas)
     * @param {Date} params.before - Only return messages before this date
     * @param {Date} params.after - Only return messages after this date
     * @param {String | Number} params.pageToken - Token used to retrieve a certain page in the list
     * @param {String} params.from - Only return messages sent from this address
     * @param {String} params.to - Only return messages sent to this address
     * @param {String[]} params.participants - Array of email addresses: only return messages with at least one of these participants are involved.
     * Due to Nylas api limitation the participants filter will only be applied when an 'after' filter is applied and limit and offset will be ignored
     * @param {String} params.folder - Only return messages in a specific folder
     * @param {Boolean} params.includeDrafts - Whether to include drafts or not, defaults to false
     * @param {String} params.subject
     *
     * @param {Object} options
     * @param {Boolean} [options.raw] - If true the response will not be transformed to the unified object
     * @param {Boolean} [options.idsOnly] - If true the response will only contain the ids of the messages
     *
     * @param {function(Error?, MessageListResource | { messages: Array.<Object | String>, next_page_token: String? }?):void} callback Returns an array of unified message resources when options.raw is falsy or the raw response of the API when truthy
     * @returns {void}
     */
    //TODO: rewrite ugly files logic
    listMessages(auth, params, options, callback) {

        if (typeof options === 'function') {
            callback = options;
            options = {};
        }

        options = options || {};
        const clonedParams = { ...params };

        const nylas = Nylas.with(auth.access_token);

        let participant = null;

        if (clonedParams.after && clonedParams.participants && clonedParams.participants.length > 0) {
            if (clonedParams.participants.length > 1) {
                return Async.mapLimit(clonedParams.participants, 5, (clonedParticipant, callback) => {

                    clonedParams.participants = [clonedParticipant];

                    return this.listMessages(auth, clonedParams, options, callback);
                }, (err, results) => {

                    if (err) {
                        return callback(err);
                    }

                    if (options.idsOnly) {
                        return callback(null, { messages: results.flatMap((result) => result.messages) });
                    }

                    // eslint-disable-next-line unicorn/no-array-reduce
                    const responseObject = results.reduce((accumulator, currentValue) => {

                        return {
                            messages: [...accumulator.messages, ...(currentValue.messages.filter((message) => {

                                return !accumulator.messages.some((m) => {

                                    return m.service_message_id === message.service_message_id;
                                });
                            }))],
                            next_page_token: 0
                        };
                    }, { messages: [], next_page_token: 0 });

                    responseObject.messages = responseObject.messages.sort((a, b) => b.date - a.date).slice(0, clonedParams.limit);

                    return callback(null, responseObject);
                });
            }

            participant = clonedParams.participants[0];
        }

        const nylasParams = {
            limit: clonedParams.limit,
            received_before: clonedParams.before,
            received_after: clonedParams.after,
            to: clonedParams.to,
            from: clonedParams.from,
            subject: clonedParams.subject,
            in: clonedParams.folder,
            has_attachment: clonedParams.hasAttachment,
            view: options.idsOnly ? 'ids' : 'expanded' // `expanded` so we get headers (message id) back
        };

        if (!clonedParams.includeDrafts) {
            nylasParams.not_in = 'drafts';
        }

        if (participant) {
            nylasParams.any_email = participant;
        }

        if (clonedParams.pageToken) {
            nylasParams.offset = Number.parseInt(clonedParams.pageToken);
        }

        return nylas.messages.list(nylasParams).then((response) => {

            // eslint-disable-next-line unicorn/explicit-length-check
            const limit = clonedParams.limit || response.length;

            const responseObject = {
                messages: response,
                next_page_token: (limit + (nylasParams.offset || 0)).toString()
            };

            if (options.raw || options.idsOnly) {
                return callback(null, responseObject);
            }

            responseObject.messages = responseObject.messages.map((message) => {

                message.email_message_id = message.headers['Message-Id'];

                if (message.headers['In-Reply-To'] && message.headers['In-Reply-To'].length > 0) {
                    message.in_reply_to = message.headers['In-Reply-To'];
                }
                else {
                    message.in_reply_to = null;
                }

                if (message.files && message.files.length > 0) {
                    message.files = message.files.map((file) => {

                        const fileObject = {
                            message,
                            metadata: file
                        };

                        return this._transformFiles(fileObject)[0];
                    }).filter((x) => !!x);
                }

                return this._transformMessages(message)[0];
            });

            return callback(null, responseObject);
        })
            .catch((err) => {

                let statusCode = 500;

                if (err.message) {
                    if (err.message.includes('Couldn\'t find')) {
                        statusCode = 404;
                    }

                    if (err.message.includes('Too many concurrent query requests')) {
                        statusCode = 429;
                    }
                }

                return callback(Boom.boomify(err, { statusCode }));
            });
    }

    sendMessage() {

        throw new Error('Not yet implemented!');
    }

    /* FILES */

    /**
     * @param {Object} auth - Authentication object
     * @param {String} auth.access_token - Access token
     * @param {String} auth.refresh_token - Refresh token
     * @param {Object} params - same as listMessages
     * @param {Object} options
     * @param {Boolean} [options.raw] - If true the response will not be transformed to the unified object
     * @param {function(Error?, FileListResource | { files: Array.<Object>, next_page_token: String? }):void} callback Returns an array of unified file resources when options.raw is falsy or the raw response of the API when truthy
     *
     * @returns {void}
     */
    //TODO: implement decent solution
    listFiles(auth, params, options, callback) {

        if (typeof options === 'function') {
            callback = options;
            options = {};
        }

        options = options || {};

        params.hasAttachment = true;

        return this.listMessages(auth, params, options, (err, response) => {

            if (err) {
                return callback(err);
            }

            return Async.reduce(response.messages, [], (files, message, callback) => {

                return callback(null, [...files, ...message.files]);
            }, (err, files) => {

                if (err) {
                    return callback(err);
                }

                return callback(null, { files, next_page_token: response.next_page_token });
            });
        });
    }

    /**
     *
     * @throws
     *
     * @param {Object} auth - Authentication object
     * @param {String} auth.access_token - Access token
     *
     * @param {Object} params
     * @param {String} params.id - Nylas attachment id
     *
     * @param {Object} options
     * @param {Boolean} [options.raw] If true the response will not be transformed to the unified object
     *
     * @param {function(Error, FileResource | Object):void} callback  Returns a unified file resource when options.raw is falsy or the raw response of the API when truthy
     * @returns {void}
     */
    getFile(auth, params, options, callback) {

        if (!params || !params.id) {
            throw new Error('Invalid configuration. Please refer to the documentation to get the required fields.');
        }

        if (typeof options === 'function') {
            callback = options;
            options = {};
        }

        options = options || {};

        const nylas = Nylas.with(auth.access_token);

        const file = nylas.files.build({ id: params.id });

        const fileObject = {};

        return file.metadata().then((metadata) => {

            fileObject.metadata = metadata;

            return file.download().then((download) => {

                fileObject.download = download;

                fileObject.metadata.contentDisposition = fileObject.metadata.contentDisposition || download['content-disposition'];

                return this.getMessage(auth, { id: fileObject.metadata.message_ids[0] }, options, (err, message) => {

                    if (err) {
                        return callback(Boom.boomify(err, 500));
                    }

                    fileObject.message = message;

                    if (options.raw) {
                        return callback(null, fileObject);
                    }

                    return callback(null, this._transformFiles(fileObject)[0]);
                });
            })
                .catch((err) => {

                    return callback(Boom.boomify(err, { statusCode: err.message.includes('Couldn\'t find') ? 404 : 500 }));
                });
        })
            .catch((err) => {

                return callback(Boom.boomify(err, { statusCode: err.message.includes('Couldn\'t find') ? 404 : 500 }));
            });
    }

    /**
     * Transform raw service messages to unified messages
     *
     * @param {Object} auth
     * @param {Object} messages
     * @returns {Array.<Object>}
     */
    transformMessages(auth, messages) {

        return this._transformMessages(messages);
    }

    /* TRANSFORMERS */

    /**
     * Transforms a raw Nylas API messages response to a unified message resource
     *
     * @param {Object[]} messagesArray - Array of messages in the format returned by the Nylas API
     *
     * @returns {Array.<MessageResource>} - Array of unified message resources
     */
    _transformMessages(messagesArray) {

        if (!messagesArray || messagesArray.length === 0) {
            return [];
        }

        const messages = Array.isArray(messagesArray) ? messagesArray : [messagesArray];

        return messages.map((message) => {

            const formattedMessage = {
                email_message_id: message.email_message_id,
                service_message_id: message.id,
                service_thread_id: message.threadId,
                date: Number(message.date),
                subject: message.subject,
                folders: [],
                attachments: message.files.length > 0,
                body: [{ content: message.body, type: 'text/html' }],
                addresses: internals.getAddressesObject(message),
                in_reply_to: message.in_reply_to,
                service_type: this.name,
                headers: message.headers,
                files: message.files
            };

            if (message.folder) {
                formattedMessage.folders = [message.folder.display_name];
            }

            if (message.textBody) {
                formattedMessage.body.push(message.textBody);
            }

            return formattedMessage;
        });
    }

    /**
     * Transforms a raw Nylas API response to a unified file resource
     *
     * @param {Array.<Object> | Object} filesArray - Array of messages in the format returned by the Nylas API
     *
     * @returns {Array.<FileResource>} - Array of unified file resources
     */
    _transformFiles(filesArray) {

        const files = Array.isArray(filesArray) ? filesArray : [filesArray];

        return files.map((file) => {

            /*
             * The fallbacks in mapping are for when the metadata comes from the actual `.metadata()` call
             * which returns different keys ¯\_(ツ)_/¯
             * /files/{id} also doesn't return the content disposition while /files does ¯\_(ツ)_/¯
             * but that is getting monkey patched in `this.getFile` so `metadata.contentDisposition` is always there
             */
            const formattedFile = {
                type: file.metadata.contentType || file.metadata.content_type,
                size: file.metadata.size,
                service_message_id: file.message.id || file.message.service_message_id,
                service_thread_id: file.message.threadId || file.message.service_thread_id,
                email_message_id: file.message.email_message_id,
                file_name: file.metadata.filename,
                service_file_id: file.metadata.id,
                content_disposition: file.metadata.contentDisposition,
                content_id: file.metadata.contentId || file.metadata.content_id,
                service_type: this.name,
                is_embedded: file.metadata.contentDisposition ? file.metadata.contentDisposition.startsWith('inline;') : false,
                addresses: internals.getAddressesObject(file.message),
                date: file.message.date
            };

            if (file.download && file.download.body) {
                formattedFile.data = file.download.body.toString('base64');
            }

            if (file.metadata.contentId) {
                formattedFile.content_id = file.metadata.contentId.replace(/[<>]/g, '');
            }

            if (file.message.body.includes(`cid:${formattedFile.content_id}`)) {
                formattedFile.is_embedded = true;
                formattedFile.content_disposition = formattedFile.content_disposition.replace('attachment;', 'inline;');
            }

            return formattedFile;
        });
    }

    // Dummy implementation since Nylas access tokens are not short-lived
    refreshAuthCredentials(auth, callback) {

        return callback(null, auth);
    }
}

/* Internal utility functions */

/**
 * @param {({ addresses: any , from: any[], to: any[], cc: any[], bcc: any[] })} message
 * @returns {{ from: any[], to: any[], cc: any[], bcc: any[] }}
 */
internals.getAddressesObject = (message) => {

    if (message.addresses) {
        return message.addresses;
    }

    return {
        from: message.from.map((contact) => {

            contact.email = contact.email && contact.email.toLowerCase();
            delete contact.connection;

            return contact;
        })[0],
        to: message.to.map((contact) => {

            contact.email = contact.email && contact.email.toLowerCase();
            delete contact.connection;

            return contact;
        }),
        cc: message.cc.map((contact) => {

            contact.email = contact.email && contact.email.toLowerCase();
            delete contact.connection;

            return contact;
        }),
        bcc: message.bcc.map((contact) => {

            contact.email = contact.email && contact.email.toLowerCase();
            delete contact.connection;

            return contact;
        })
    };
};

module.exports = NylasConnector;
