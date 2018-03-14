'use strict';

const EventEmitter = require('events');

const Async = require('async');
const Google = require('googleapis');
const Gmail = Google.gmail('v1');
const OAuth2 = Google.auth.OAuth2;
const Batchelor = require('batchelor');
const Boom = require('boom');
const Utils = require('./utils');

const internals = {};

class GmailConnector extends EventEmitter {

    /**
     * @constructor
     *
     * @param {Object} config - Configuration object
     * @param {string} config.clientId
     * @param {string} config.clientSecret
     */
    constructor(config) {

        super();

        if (!config || !config.clientId || !config.clientSecret) {
            throw new Error('Invalid configuration. Please refer to the documentation to get the required fields.');
        }

        this.clientId = config.clientId;
        this.clientSecret = config.clientSecret;

        this.name = 'gmail';
    }

    /* MESSAGES */

    /**
     *
     * @param {Object} auth - Authentication object
     * @param {string} auth.access_token - Access token
     * @param {string} auth.refresh_token - Refresh token
     *
     * @param {Object} params
     * @param {string} params.id - Gmail message id
     *
     * @param {Object} options
     * @param {boolean} options.raw - If true the response will not be transformed to the unified object
     *
     * @returns {MessageResource | Object} Returns a unified message resource when options.raw is falsy or the raw response of the API when truthy
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

        const gmailParams = {
            auth,
            userId: 'me',
            id: params.id
        };

        return this._callAPI(Gmail.users.messages.get, gmailParams, (err, response) => {

            if (err) {
                return callback(err);
            }

            if (options.raw) {
                return callback(null, response);
            }

            return this._transformMessages(response, (err, transformedMessages) => {

                if (err) {
                    return callback(err);
                }

                return callback(null, transformedMessages[0]);
            })
        });
    }

    /**
     *
     * @param {Object} auth - Authentication object
     * @param {string} auth.access_token - Access token
     * @param {string} auth.refresh_token - Refresh token
     *
     * @param {Object} params
     * @param {number} params.limit - Maximum amount of messages in response, max = 100
     * @param {boolean} params.hasAttachment - If true, only return messages with attachments
     * @param {Date} params.before - Only return messages before this date
     * @param {Date} params.after - Only return messages after this date
     * @param {string} params.pageToken - Token used to retrieve a certain page in the list
     * @param {string} params.from - Only return messages sent from this address
     * @param {string} params.to - Only return messages sent to this address
     * @param {string[]} params.participants - Array of email addresses: only return messages with at least one of these participants are involved
     * @param {string[]} params.folder - Only return messages in these folders
     * @param {boolean} params.includeDrafts - Whether to include drafts or not, defaults to false
     *
     * @param {Object} options
     * @param {boolean} options.raw - If true the response will not be transformed to the unified object
     *
     * @returns {MessageListResource | Object[]} Returns an array of unified message resources when options.raw is falsy or the raw response of the API when truthy
     * @returns
     */
    listMessages(auth, params, options, callback) {

        if (typeof options === 'function') {
            callback = options;
            options = {};
        }
        options = options || {};

        const gmailParams = {
            auth,
            userId: 'me'
        };
        let q = '-(in:chats) ';

        if (params.limit || params.limit === 0) {
            gmailParams.maxResults = params.limit > 100 ? 100 : params.limit;
        }
        else {
            gmailParams.maxResults = 100;
        }

        if (params.pageToken) {
            gmailParams.pageToken = params.pageToken;
        }

        if (!params.includeDrafts) {
            q += '-(in:draft) '
        }

        if (params.hasAttachment) {
            q += 'has:attachment ';
        }

        if (params.before) {
            q += `before:${Math.ceil(params.before.getTime() / 1000)} `;
        }

        if (params.after) {
            q += `after:${Math.ceil(params.after.getTime() / 1000)} `;
        }

        if (params.participants) {
            q += `{${params.participants.map((participant) => {

                return `from:${participant} to:${participant} cc:${participant} `;
            }).join('')}}`;
        }

        if (params.from) {
            q += `from:${params.from} `;
        }

        if (params.to) {
            q += `to:${params.to} `;
        }

        if (params.subject) {
            // does not match literally, e.g. params.subject = 'test' would also match email subject 'this is a test'
            q += `subject:"${params.subject}" `;
        }

        if (params.folder) {
            q += `in:${params.folder}`;
        }

        if (q) {
            gmailParams.q = q;
        }

        return this._callAPI(Gmail.users.messages.list, gmailParams, (err, listResponse) => {

            if (err) {
                return callback(err);
            }

            if (!listResponse.messages || listResponse.messages.length === 0) {
                return callback(null, { messages: [] });
            }

            const batch = new Batchelor({
                uri: 'https://www.googleapis.com/batch',
                method:'POST',
                auth: {
                    'bearer': gmailParams.auth.credentials.access_token
                },
                headers: {
                    'Content-Type': 'multipart/mixed'
                }
            });

            listResponse.messages.forEach((message) => {

                batch.add({
                    method: 'GET',
                    path: `/gmail/v1/users/me/messages/${message.id}`
                })
            });

            return batch.run((err, batchResponse) => {

                if (err) {
                    return callback(err);
                }

                const erroredParts = batchResponse.parts.filter((part) => {

                    return part.statusCode > 399;
                });

                if (erroredParts.length > 0) {
                    return callback(Boom.boomify(new Error(erroredParts[0].body.error.message), { statusCode: erroredParts[0].statusCode ? Number(erroredParts[0].statusCode) : 500 }));
                }

                return Async.map(batchResponse.parts, (messagePart, callback) => {

                    if (options.raw) {
                        return callback(null, messagePart.body);
                    }

                    return this._transformMessages(messagePart.body, (err, transformedMessages) => {

                        if (err) {
                            return callback(err);
                        }

                        return callback(null, transformedMessages[0])
                    });
                }, (err, messages) => {
                
                    if (err) {
                        return callback(err);
                    }

                    const responseObject = {
                        messages
                    };

                    if (listResponse.nextPageToken) {
                        responseObject.next_page_token = listResponse.nextPageToken;
                    }

                    return callback(null, responseObject);
                });
            });
        });
    }

    /* FILES */

    /**
     *
     * @param {Object} auth - Authentication object
     * @param {string} auth.access_token - Access token
     * @param {string} auth.refresh_token - Refresh token
     *
     * @param {Object} params - same as listMessages
     *
     * @param {Object} options
     * @param {boolean} options.raw - If true the response will not be transformed to the unified object
     *
     * @returns {FileListResource | Object[]} Returns an array of unified file resources when options.raw is falsy or the raw response of the API when truthy
     */
    listFiles(auth, params, options, callback) {

        if (typeof options === 'function') {
            callback = options;
            options = {};
        }
        options = options || {};

        const gmailParams = {
            ...params,
            auth,
            userId: 'me',
            hasAttachment: true
        };

        return this.listMessages(auth, gmailParams, { raw: true }, (err, messagesList) => {

            if (err) {
                return callback(err);
            }

            if (messagesList.messages.length === 0) {
                return callback(null, { files: [] });
            }

            if (options.raw) {
                return callback(null, messagesList.messages);
            }

            // Lowercase the header name to make sure we don't get derpes due to Message-Id vs message-id
            return this._transformMessages(messagesList.messages, (err, messages) => {

                const fileListObject = {
                    files: messages.map(message => message.files).reduce((allFiles, files) => allFiles.concat(files))
                };

                if (messagesList.nextPageToken) {
                    fileListObject.next_page_token = messagesList.nextPageToken;
                }

                return callback(null, fileListObject);
            });
        });
    }

    /**
     *
     * @param {Object} auth - Authentication object
     * @param {string} auth.access_token - Access token
     * @param {string} auth.refresh_token - Refresh token
     *
     * @param {Object} params
     * @param {string} params.id - The id of the attachment
     * @param {string} params.messageId - The id of the message containing the attachment
     *
     * @returns {Object} - Content of the attachment
     */
    getFile(auth, params, callback) {

        if (!params || !params.id || !params.messageId) {
            throw new Error('Invalid configuration. Please refer to the documentation to get the required fields.');
        }

        const messageParams = {
            auth,
            userId: 'me',
            id: params.messageId
        };

        return this._callAPI(Gmail.users.messages.get, messageParams, (err, messageResponse) => {

            if (err) {
                return callback(err);
            }

            const nonContainerParts = this._extractNonContainerParts(messageResponse.payload);

            const fileParts = nonContainerParts.filter((part) => {

                return part.partId === params.id;
            });

            const fileParams = {
                auth,
                userId: 'me',
                id: fileParts[0].body.attachmentId,
                messageId: params.messageId
            };

            //return callback(null, this.transformMessages(response)[0]);
            return this._callAPI(Gmail.users.messages.attachments.get, fileParams, (err, fileResponse) => {

                if (err) {
                    return callback(err);
                }

                fileResponse.service_file_id = fileParts[0].partId;
                fileResponse.id = fileParams.id;
                fileResponse.type = fileParts[0].mimeType;
                fileResponse.file_name = fileParts[0].filename;

                return callback(null, fileResponse);
            });
        });
    }

    /* TRANSFORMERS */

    /**
     * Transforms a raw Gmail API response to a unified file resource
     *
     * @param {Object[]} messagesArray - Array of messages in the format returned by the Gmail API
     *
     * @returns {FileResource[]} - Array of unified file resources
     */
    _transformFiles(messagesArray) {

        const messages = Array.isArray(messagesArray) ? messagesArray : [messagesArray];
        const files = [];

        messages.forEach((message) => {

            const nonContainerParts = this._extractNonContainerParts(message.payload);

            nonContainerParts.forEach((part) => {

                if (!part.filename || !part.headers) {
                    return;
                }

                const contentDisposition = internals.getHeaderValue(part.headers, 'content-disposition'); // Raw header since parsed splits into value and params
                let contentId = internals.getHeaderValue(part.headers, 'content-id');

                // Remove angle brackets around the content ID
                if (contentId) {
                    contentId = contentId.replace(/[<>]/g, "");
                }

                files.push({
                    service_type: this.name,
                    type: part.mimeType,
                    size: Number(part.body.size),
                    service_message_id: message.formattedMessage.service_message_id,
                    service_thread_id: message.formattedMessage.service_thread_id,
                    email_message_id: message.headers['message-id'][0] || null,
                    date: message.date,
                    addresses: internals.getAddressesObject(message),
                    file_name: part.filename,
                    content_id: contentId,
                    content_disposition: contentDisposition,
                    service_file_id: part.partId,
                    is_embedded: contentDisposition ? contentDisposition.startsWith('inline;') : false
                });
            });
        });

        return files;
    }

    /**
     * Transforms a raw Gmail API messages response to a unified message resource
     *
     * @param {Object[]} messagesArray - Array of messages in the format returned by the Gmail API
     *
     * @returns {MessageResource[]} - Array of unified message resources
     */
    _transformMessages(messagesArray, callback) {

        const messages = Array.isArray(messagesArray) ? messagesArray : [messagesArray];

        return Async.map(messages, (message, callback) => {

            if (!message.payload || !message.payload.headers) {
                const error = new Error('message part has no headers');
                error.message = message;

                return callback(error);
            }

            // Lowercase the header name to make sure we don't get derps due to Message-Id vs message-id
            return Utils.parseRawMail(message.payload.headers.map(h => `${h.name.toLowerCase()}: ${h.value}`).join('\n'), (err, mail) => {

                if (err) {
                    return callback(err);
                }

                const formattedMessage = {
                    service_type: this.name,
                    email_message_id: (mail.headers['message-id'] && mail.headers['message-id'].length && mail.headers['message-id'][0]) || null,
                    service_message_id: message.id,
                    service_thread_id: message.threadId,
                    date: mail.date,
                    subject: (mail.headers['subject'] && mail.headers['subject'].length && mail.headers['subject'][0]) || null,
                    folders: message.labelIds,
                    files: [],
                    body: [],
                    addresses: internals.getAddressesObject(mail),
                    in_reply_to: (mail.headers['in-reply-to'] && mail.headers['in-reply-to'].length && mail.headers['in-reply-to'][0]) || null,
                    headers: mail.headers
                };

                // Files
                mail.payload = message.payload;
                mail.formattedMessage = formattedMessage;

                const files = this._transformFiles(mail);
                if (files.length > 0) {
                    formattedMessage.attachments = true;
                    formattedMessage.files = formattedMessage.files.concat(files);
                }

                // Extract body
                const nonContainerParts = this._extractNonContainerParts(message.payload);
                const allowedBodyMimeTypes = ['text/html', 'text/plain'];

                nonContainerParts.forEach((nonContainerPart) => {

                    if (allowedBodyMimeTypes.includes(nonContainerPart.mimeType.toLowerCase()) && nonContainerPart.body.data) {
                        formattedMessage.body.push({
                            type: nonContainerPart.mimeType,
                            content: Buffer.from(nonContainerPart.body.data, 'base64').toString()
                        })
                    }
                });

                return callback(null, formattedMessage);
            });
        }, callback);
    }

    refreshAuthCredentials(auth, callback) {

        if (!auth.expiration_date || new Date(auth.expiration_date) > new Date()) {
            return callback(null, auth);
        }

        const oauth2Client = new OAuth2(
            this.clientId,
            this.clientSecret
        );

        oauth2Client.credentials = {
            refresh_token: auth.refresh_token
        };

        return oauth2Client.refreshAccessToken((err, token) => {

            if (err) {
                return callback(err);
            }

            this.emit('newAccessToken', token);

            return callback(null, token);
        });
    }

    _extractNonContainerParts(part) {

        if (part.parts && part.parts.length > 0) {

            // Flatten the array
            return [].concat.apply([], part.parts.map((multipart) => {

                return this._extractNonContainerParts(multipart);
            }));
        }

        return [part];
    }

    _callAPI(method, params, callback) {

        if (!(params.auth instanceof OAuth2)) {
            const oauth2Client = new OAuth2(
                this.clientId,
                this.clientSecret
            );

            oauth2Client.credentials = {
                access_token: params.auth.access_token,
                refresh_token: params.auth.refresh_token
            };

            params.auth = oauth2Client;
        }

        const oldAccessToken = params.auth.credentials.access_token;

        return method(params, (err, results) => {

            if (params.auth.credentials.access_token !== oldAccessToken) {
                this.emit('newAccessToken', { ...params.auth.credentials });
            }

            if (err) {
                return callback(Boom.boomify(err, { statusCode: err.code ? Number(err.code) : 500 }));
            }

            return callback(null, results);
        });
    }
}

/* Internal utility functions */

internals.getAddressesObject = (mail) => {

    return {
        to: mail.to && internals.getAddressesObjectFromValue(mail.to.value),
        from: mail.from && internals.getAddressesObjectFromValue(mail.from.value)[0] || {},
        cc: mail.cc && internals.getAddressesObjectFromValue(mail.cc.value),
        bcc: mail.bcc && internals.getAddressesObjectFromValue(mail.bcc.value)
    };
};

internals.getAddressesObjectFromValue = (value) => {

    return value && value.map(vObject => {

        if (!vObject.address) {
            return undefined;
        }

        return {
            name: vObject.name,
            email: vObject.address.toLowerCase()
        }
    }).filter(address => !!address);
}

internals.getHeaderValue = (headers, name) => {

    const headerObject = headers.find((header) => {

        return header.name.toLowerCase() === name.toLowerCase();
    });

    if (headerObject && headerObject.value) {
        return headerObject.value;
    }

    return null;
};

module.exports = GmailConnector;
