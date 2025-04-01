/* eslint-disable no-underscore-dangle */
'use strict';

const EventEmitter = require('events');

const _ = require('lodash');
const Async = require('async');
const Batchelor = require('@salesflare/batchelor');
const Boom = require('@hapi/boom');

const Gmail = require('@googleapis/gmail').gmail('v1');
const OAuth2 = require('@googleapis/gmail').auth.OAuth2;

const Utils = require('./utils');

const internals = {};

/**
 * @typedef {import('./index').MessageResource} MessageResource
 * @typedef {import('./index').MessageListResource} MessageListResource
 * @typedef {import('./index').FileResource} FileResource
 * @typedef {import('./index').FileListResource} FileListResource
 * @typedef {import('./index').MessageRecipient} MessageRecipient
 */

class GmailConnector extends EventEmitter {

    /**
     * @class
     * @throws
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

        this.name = 'gmail';
    }

    /**
     * @typedef {Object} Auth - Authentication object
     * @property {String} access_token
     * @property {String} refresh_token
     * @property {Date} expiration_date
     * @property {*} [id] - will be passed back when emitting `newAccessToken`
     */

    /* MESSAGES */

    /**
     * @throws
     *
     * @param {Auth} auth
     *
     * @param {Object} params
     * @param {String} params.id - Gmail message id
     * @param {String} [params.rfc2822Format=false] - Return the email in rfc2822 format https://www.ietf.org/rfc/rfc2822.txt
     *
     * @param {Object} options
     * @param {Boolean} [options.raw] - If true the response will not be transformed to the unified object
     *
     * @param {function(Error?, ( MessageResource | Object | String)?):void} callback Returns a unified message resource when options.raw is falsy or the raw response of the API when truthy
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

        const gmailParams = {
            auth,
            userId: 'me',
            id: params.id
        };

        if (params.rfc2822Format) {
            gmailParams.format = 'raw';
        }

        return this._callAPI(Gmail.users.messages.get.bind(Gmail.users.messages), gmailParams, (err, response) => {

            if (err) {
                return callback(err);
            }

            if (params.rfc2822Format) {
                return callback(null, Buffer.from(response.raw, 'base64').toString());
            }

            if (options.raw) {
                return callback(null, response);
            }

            return this._transformMessages(response, (err, transformedMessages) => {

                if (err) {
                    return callback(err);
                }

                return callback(null, transformedMessages[0]);
            });
        });
    }

    /**
     *
     * @param {Auth} auth
     *
     * @param {Object} params
     * @param {Number} params.limit - Maximum amount of messages in response, max = 100
     * @param {Boolean} params.hasAttachment - If true, only return messages with attachments
     * @param {Date} params.before - Only return messages before this date
     * @param {Date} params.after - Only return messages after this date
     * @param {String} params.pageToken - Token used to retrieve a certain page in the list
     * @param {String} params.from - Only return messages sent from this address
     * @param {String} params.to - Only return messages sent to this address
     * @param {String[]} params.participants - Array of email addresses: only return messages with at least one of these participants are involved
     * @param {String} params.folder - Only return messages in a specific folder
     * @param {Boolean} params.includeDrafts - Whether to include drafts or not, defaults to false
     * @param {String} params.subject
     * @param {String} [params.q] - Allows you to pass a custom search to the api, mostly used for debugging
     *
     * @param {Object} options
     * @param {Boolean} [options.raw] - If true the response will not be transformed to the unified object
     * @param {Boolean} [options.idsOnly] - If true the response will only contain the ids of the messages
     *
     * @param {function(Error?, (MessageListResource | { messages: Array.<Object>, next_page_token: String? })?):void} callback Returns an array of unified message resources when options.raw is falsy or the raw response of the API when truthy
     *
     * @returns {void}
     */
    listMessages(auth, params, options, callback) {

        if (typeof options === 'function') {
            callback = options;
            options = {};
        }

        options = options || {};

        const paramsArray = [];
        const gmailParams = {
            auth,
            userId: 'me',
            q: '-(in:chats) '
        };

        if (params.limit || params.limit === 0) {
            gmailParams.maxResults = params.limit > 500 ? 500 : params.limit;
        }
        else {
            gmailParams.maxResults = 100;
        }

        if (params.pageToken) {
            gmailParams.pageToken = params.pageToken;
        }

        if (!params.includeDrafts) {
            gmailParams.q += '-(in:draft) ';
        }

        if (params.hasAttachment) {
            gmailParams.q += 'has:attachment ';
        }

        if (params.before) {
            gmailParams.q += `before:${Math.ceil(params.before.getTime() / 1000)} `;
        }

        if (params.after) {
            gmailParams.q += `after:${Math.ceil(params.after.getTime() / 1000)} `;
        }

        if (params.from) {
            gmailParams.q += `from:${params.from} `;
        }

        if (params.to) {
            gmailParams.q += `to:${params.to} `;
        }

        if (params.subject) {
            // Does not match literally, e.g. params.subject = 'test' would also match email subject 'this is a test'
            gmailParams.q += `subject:"${params.subject}" `;
        }

        if (params.folder) {
            gmailParams.q += `in:${params.folder} `;
        }

        if (params.q) {
            gmailParams.q += params.q;
        }

        // We split participants up in chunks of 50 since gmail can't handle more at the same time
        if (params.participants) {
            const participantsChunks = _.chunk(params.participants, 50);
            participantsChunks.forEach((participants) => {

                const tempGmailParams = { ...gmailParams };

                tempGmailParams.q += `{${participants.map((participant) => `from:${participant} to:${participant} cc:${participant} `).join('')}} `;

                paramsArray.push(tempGmailParams);
            });
        }
        else {
            paramsArray.push(gmailParams);
        }

        // Refresh manually so that when we do multiple calls we don't refresh multiple times
        return this.refreshAuthCredentials(gmailParams.auth, (err, token) => {

            if (err) {
                return callback(err);
            }

            paramsArray.forEach((param) => {

                param.auth.access_token = token.access_token;
            });

            let nextPageToken = null;

            return Async.map(paramsArray, (gmailParam, callback) => {

                return this._callAPI(Gmail.users.messages.list.bind(Gmail.users.messages), gmailParam, (err, listResponse) => {

                    if (err) {
                        return callback(err);
                    }

                    if (!listResponse.messages || listResponse.messages.length === 0) {
                        return callback(null, []);
                    }

                    nextPageToken = listResponse.nextPageToken;

                    return callback(null, listResponse.messages);
                });
            }, (err, messages) => {

                if (err) {
                    return callback(err);
                }

                return this._processMessages(messages, options, nextPageToken, paramsArray, token, callback);
            });
        });
    }

    _processMessages(messages, options, nextPageToken, paramsArray, token, callback) {
        // Flatten and filter out any falsy values
        messages = messages.flat().filter(Boolean);
        if (messages.length === 0) {
            return callback(null, { messages: [] });
        }

        if (options.idsOnly) {
            const idsOnlyResponseObject = {
                messages: messages.map((message) => message.id)
            };
            if (nextPageToken && paramsArray.length === 1) {
                idsOnlyResponseObject.next_page_token = nextPageToken;
            }

            return callback(null, idsOnlyResponseObject);
        }

        // Common batchelor parameters for batch requests
        const batchelorParams = {
            uri: 'https://www.googleapis.com/batch/gmail/v1',
            method: 'POST',
            auth: { bearer: token.access_token },
            headers: { 'Content-Type': 'multipart/mixed' }
        };

        // Create an errorBatch for failed parts to retry later
        const errorBatch = new Batchelor(batchelorParams);

        let allTransformedMessages = [];

        // Helper function to process one chunk of messages
        const processChunk = (chunk, cb) => {

            const batch = new Batchelor(batchelorParams);
            const successfulParts = [];

            // Add each message request to the batch
            chunk.forEach((message) => {

                const request = {
                    method: 'GET',
                    path: `/gmail/v1/users/me/messages/${message.id}`
                };

                batch.add({
                    ...request,
                    extend: { request },
                    callback: (response, extend) => {

                        if (Number.parseInt(response.statusCode) > 399) {
                            // Add the failed request to the errorBatch for a retry
                            errorBatch.add(extend.request);
                        }
                        else {
                            successfulParts.push(response);
                        }
                    }
                });
            });

            // Run the batch for this chunk
            batch.run((err) => {

                if (err) {
                    return cb(err);
                }

                // Process successful parts
                Async.map(successfulParts, (part, mapCb) => {

                    if (options.raw) {
                        return mapCb(null, part.body);
                    }

                    this._transformMessages(part.body, (err, transformed) => {

                        if (err) {
                            return mapCb(err);
                        }

                        // Transformed is an array; we take the first element
                        return mapCb(null, transformed[0]);
                    });
                }, (err, results) => {

                    if (err) {
                        return cb(err);
                    }

                    allTransformedMessages = [...allTransformedMessages, ...results];
                    return cb(null);
                });
            });
        };

        // Process each chunk sequentially
        Async.eachSeries(_.chunk(messages, 20), (chunk, seriesCb) => {

            processChunk(chunk, seriesCb);
        }, (err) => {

            if (err) {
                return callback(err);
            }

            // Check if there are any failed requests to retry
            if (errorBatch._requests && errorBatch._requests.length > 0) {
                if (errorBatch._requests.length > 5) {
                    const error = new Error('Too many parts errored');
                    error.statusCode = 429;
                    return callback(error);
                }

                // Retry errorBatch once
                errorBatch.run((err, retryResponses) => {

                    if (err) {
                        return callback(err);
                    }

                    if (retryResponses.errors && retryResponses.errors > 0) {
                        const error = new Error('Too many parts errored');
                        error.statusCode = 429;
                        return callback(error);
                    }

                    // Process retry responses similarly
                    Async.map(retryResponses.parts, (response, mapCb) => {

                        if (options.raw) {
                            return mapCb(null, response.body);
                        }

                        this._transformMessages(response.body, (err, transformed) => {

                            if (err) {
                                return mapCb(err);
                            }

                            return mapCb(null, transformed[0]);
                        });
                    }, (err, retryResults) => {

                        if (err) {
                            return callback(err);
                        }

                        allTransformedMessages = [...allTransformedMessages, ...retryResults];
                        // Sort the final messages by date in descending order
                        allTransformedMessages = _.orderBy(allTransformedMessages, 'date', 'desc');
                        const responseObject = { messages: allTransformedMessages };
                        if (nextPageToken && paramsArray.length === 1) {
                            responseObject.next_page_token = nextPageToken;
                        }

                        return callback(null, responseObject);
                    });
                });
            }
            else {
                // No failed parts; simply return the results
                allTransformedMessages = _.orderBy(allTransformedMessages, 'date', 'desc');
                const responseObject = { messages: allTransformedMessages };
                if (nextPageToken && paramsArray.length === 1) {
                    responseObject.next_page_token = nextPageToken;
                }

                return callback(null, responseObject);
            }
        });
    }

    /**
     * Sends a message
     *
     * @param {Auth} auth
     *
     * @param {Object} params
     * @param {String} params.text - Plain text content of message
     * @param {String} params.html - Html content of message
     * @param {String} params.subject - Subject of message,
     * @param {String} params.inReplyTo - The message id this message is replying
     * @param {String} params.threadId - Id of the thread this message should be created in
     * @param {MessageRecipient} params.from
     * @param {MessageRecipient[]} params.to
     * @param {MessageRecipient[]} params.cc
     * @param {MessageRecipient[]} params.bcc
     * @param {{ name: String, url: String }[]} params.attachments
     *
     * @param {Object} options
     *
     * @param {function(Error?, String?):void} callback
     *
     * @returns {void}
     */
    sendMessage(auth, params, options, callback) {

        if (typeof options === 'function') {
            callback = options;
        }

        const mailOptions = {
            from: internals.convertMessageRecipientsToCsv(params.from),
            to: internals.convertMessageRecipientsToCsv(params.to),
            cc: internals.convertMessageRecipientsToCsv(params.cc),
            bcc: internals.convertMessageRecipientsToCsv(params.bcc),
            text: params.text,
            html: params.html,
            subject: params.subject,
            inReplyTo: params.inReplyTo
        };

        if (params.attachments && params.attachments.length > 0) {
            mailOptions.attachments = params.attachments.map((attachment) => {

                return {
                    filename: attachment.name,
                    path: attachment.url
                };
            });
        }

        return Utils.generateMessage(mailOptions, { base64Encoded: false }, (err, rawMessage) => {

            if (err) {
                return callback(err);
            }

            // To allow bigger (> 6-7MB) attachments we use the media property
            // This will use the upload endpoint of the api https://developers.google.com/gmail/api/v1/reference/users/messages/send
            // Found in https://github.com/googleapis/google-api-nodejs-client/issues/1491#issuecomment-446005442
            const gmailSendParams = {
                auth,
                userId: 'me',
                uploadType: 'multipart',
                media: {
                    mimeType: 'message/rfc822',
                    body: rawMessage
                }
            };

            if (params.threadId) {
                gmailSendParams.requestBody = { threadId: params.threadId };
            }

            return this._callAPI(Gmail.users.messages.send.bind(Gmail.users.messages), gmailSendParams, (err, sendResponse) => {

                if (err) {
                    // Could mean that we are replying to an email not in our mailbox, if so we omit the thread id
                    if (err.code === 404 && gmailSendParams.requestBody && gmailSendParams.requestBody.threadId) {
                        delete gmailSendParams.requestBody.threadId;

                        return this._callAPI(Gmail.users.messages.send.bind(Gmail.users.messages), gmailSendParams, (err, sendResponseWithoutThread) => {

                            if (err) {
                                return callback(err);
                            }

                            // Sending a message through the gmail API only returns the gmail message id so we fetch the message afterwards to return the email message id.
                            return this.getMessage(auth, { id: sendResponseWithoutThread.id }, (err, getMessageResponse) => {

                                if (err) {
                                    return callback();
                                }

                                return callback(null, getMessageResponse.email_message_id);
                            });
                        });
                    }

                    return callback(err);
                }

                // Sending a message through the gmail API only returns the gmail message id so we fetch the message afterwards to return the email message id.
                return this.getMessage(auth, { id: sendResponse.id }, (err, getMessageResponse) => {

                    if (err) {
                        return callback();
                    }

                    return callback(null, getMessageResponse.email_message_id);
                });
            });
        });
    }

    /* FILES */

    /**
     *
     * @param {Auth} auth
     *
     * @param {Object} params - same as listMessages
     *
     * @param {Object} options
     * @param {Boolean} [options.raw] - If true the response will not be transformed to the unified object
     *
     * @param {function(Error?, FileListResource | { files: Array.<Object>, next_page_token: String? } | Array.<Object>):void} callback Returns an array of unified file resources when options.raw is falsy or the raw response of the API when truthy
     *
     * @returns {void}
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

                if (err) {
                    return callback(err);
                }

                const fileListObject = {
                    // eslint-disable-next-line unicorn/no-array-reduce
                    files: messages.map((message) => message.files).reduce((allFiles, files) => [...allFiles, ...files])
                };

                if (messagesList.next_page_token) {
                    fileListObject.next_page_token = messagesList.next_page_token;
                }

                return callback(null, fileListObject);
            });
        });
    }

    /**
     * @throws
     *
     * @param {Auth} auth
     *
     * @param {Object} params
     * @param {String} params.id - The id of the attachment
     * @param {String} params.messageId - The id of the message containing the attachment
     *
     * @param {Object} options
     *
     * @param {function(Error?, FileResource?):void} callback - Content of the attachment
     *
     * @returns {void}
     */
    getFile(auth, params, options, callback) {

        if (!params || !params.id || !params.messageId) {
            throw new Error('Invalid configuration. Please refer to the documentation to get the required fields.');
        }

        if (typeof options === 'function') {
            callback = options;
        }

        const messageParams = {
            auth,
            userId: 'me',
            id: params.messageId
        };

        return this._callAPI(Gmail.users.messages.get.bind(Gmail.users.messages), messageParams, (err, messageResponse) => {

            if (err) {
                return callback(err);
            }

            const nonContainerParts = this._extractNonContainerParts(messageResponse.payload);

            const fileParts = nonContainerParts.filter((part) => {

                return part.partId === params.id;
            });

            if (fileParts.length === 0) {
                return callback(new Error(`No matching file parts found for message ${params.messageId} and file ${params.id}`));
            }

            const fileParams = {
                auth,
                userId: 'me',
                id: fileParts[0].body.attachmentId,
                messageId: params.messageId
            };

            return this._callAPI(Gmail.users.messages.attachments.get.bind(Gmail.users.messages.attachments), fileParams, (err, fileResponse) => {

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
     * Transforms a raw Gmail API response to a unified file resource
     *
     * @param {Object[]} messagesArray - Array of messages in the format returned by the Gmail API
     * @param {String?} htmlBody - Option html body of the message,used to check for inline attachments
     *
     * @returns {Array.<FileResource>} - Array of unified file resources
     */
    _transformFiles(messagesArray, htmlBody) {

        const messages = Array.isArray(messagesArray) ? messagesArray : [messagesArray];
        const files = [];

        messages.forEach((message) => {

            const nonContainerParts = this._extractNonContainerParts(message.payload);

            nonContainerParts.forEach((part) => {

                if (!part.filename || !part.headers) {
                    return;
                }

                let contentDisposition = internals.getHeaderValue(part.headers, 'content-disposition'); // Raw header since parsed splits into value and params
                let contentId = internals.getHeaderValue(part.headers, 'content-id');

                // Remove angle brackets around the content ID
                if (contentId) {
                    contentId = contentId.replace(/[<>]/g, '');
                }

                let isEmbedded;

                // Some files are marked as attachment while they are actually embedded
                // We consider a file embedded when there's a content ID and the HTML body contains a reference to it
                // If it has a content disposition that contains inline we also consider it embedded
                // If there's only a text body the attachment will never be inline
                if (htmlBody && contentId && htmlBody.includes(`cid:${contentId}`)) {
                    isEmbedded = true;
                    contentDisposition = contentDisposition ? contentDisposition.replace('attachment;', 'inline;') : 'inline';
                }
                else {
                    isEmbedded = !!(contentDisposition && contentDisposition.startsWith('inline;'));
                }

                files.push({
                    service_type: this.name,
                    type: part.mimeType,
                    size: Number(part.body.size),
                    service_message_id: message.formattedMessage.service_message_id,
                    service_thread_id: message.formattedMessage.service_thread_id,
                    email_message_id: (message.headers['message-id'] && message.headers['message-id'].length > 0 && message.headers['message-id'][0]) || null,
                    date: message.date,
                    addresses: internals.getAddressesObject(message),
                    file_name: part.filename,
                    content_id: contentId,
                    content_disposition: contentDisposition,
                    service_file_id: part.partId,
                    is_embedded: isEmbedded
                });
            });
        });

        return files;
    }

    /**
     * Transforms a raw Gmail API messages response to a unified message resource
     *
     * @param {Object[]} messagesArray - Array of messages in the format returned by the Gmail API
     * @param {function(Error?, Array.<MessageResource>):void} callback
     *
     * @returns {void} - Array of unified message resources
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
            return Utils.parseRawMail(message.payload.headers.map((h) => `${h.name.toLowerCase()}: ${h.value}`).join('\n'), (err, mail) => {

                if (err) {
                    return callback(err);
                }

                if (!mail.date) {
                    const dateHeader = message.payload.headers.find((header) => {

                        return header.name === 'Date';
                    });

                    if (dateHeader) {
                        mail.date = new Date(dateHeader.value);
                    }
                    else if (message.internalDate) {
                        mail.date = new Date(Number.parseInt(message.internalDate));
                    }
                }

                const formattedMessage = {
                    service_type: this.name,
                    email_message_id: (mail.headers['message-id'] && mail.headers['message-id'].length > 0 && mail.headers['message-id'][0]) || null,
                    service_message_id: message.id,
                    service_thread_id: message.threadId,
                    date: mail.date || (mail.headers['delivery-date'] && mail.headers['delivery-date'].length > 0 && new Date(mail.headers['delivery-date'][0])),
                    subject: (mail.headers.subject && mail.headers.subject.length > 0 && mail.headers.subject[0]) || null,
                    folders: message.labelIds,
                    files: [],
                    body: [],
                    addresses: internals.getAddressesObject(mail),
                    in_reply_to: (mail.headers['in-reply-to'] && mail.headers['in-reply-to'].length > 0 && mail.headers['in-reply-to'][0]) || null,
                    headers: mail.headers
                };

                // Extract body
                const nonContainerParts = this._extractNonContainerParts(message.payload);
                const allowedBodyMimeTypes = new Set(['text/html', 'text/plain']);

                nonContainerParts.forEach((nonContainerPart) => {

                    if (allowedBodyMimeTypes.has(nonContainerPart.mimeType.toLowerCase()) && nonContainerPart.body.data) {
                        formattedMessage.body.push({
                            type: nonContainerPart.mimeType,
                            content: Buffer.from(nonContainerPart.body.data, 'base64').toString()
                        });
                    }
                });

                // Files
                mail.payload = message.payload;
                mail.formattedMessage = formattedMessage;

                const htmlBody = formattedMessage.body.find((body) => body.type === 'text/html')?.content;
                const files = this._transformFiles(mail, htmlBody);
                if (files.length > 0) {
                    formattedMessage.attachments = true;
                    formattedMessage.files = [...formattedMessage.files, ...files];
                }

                return callback(null, formattedMessage);
            });
        }, callback);
    }

    /**
     * Emits `newAccessToken` when a new access token for the refresh token was generated
     *
     * @param {Auth} auth
     * @param {function(Error?, Auth?):void} callback
     * @returns {void}
     */
    refreshAuthCredentials(auth, callback) {

        if (auth.access_token && (!auth.expiration_date || new Date(auth.expiration_date) > new Date())) {
            return callback(null, auth);
        }

        const oauth2Client = new OAuth2(this.clientId, this.clientSecret);

        oauth2Client.setCredentials({
            refresh_token: auth.refresh_token
        });

        return oauth2Client.refreshAccessToken((err, token) => {

            if (err) {
                return callback(err);
            }

            // Make sure we emit the id for reference and also pass it for chaining abilities
            token.id = auth.id;

            this.emit('newAccessToken', token);

            return callback(null, token);
        });
    }

    _extractNonContainerParts(part) {

        if (part.parts && part.parts.length > 0) {

            // Flatten the array
            return part.parts.flatMap((multipart) => {

                return this._extractNonContainerParts(multipart);
            });
        }

        return [part];
    }

    /**
     * @throws
     *
     * @param {function(Object, Function):void} method
     *
     * @param {Object} params
     * @param {Auth} params.auth
     *
     * @param {function(Error?, Object?):void} callback
     * @returns {void}
     */
    _callAPI(method, params, callback) {

        if (!method.name.startsWith('bound ') || Object.prototype.hasOwnProperty.call(method, 'prototype')) {
            throw new Error('Gmail functions need to be bound using `.bind`. We wrap Gmail function for auth and error handling and this causes them to lose their `this` so you need to explicitly bind the this. For example: `Gmail.users.messages.get.bind(Gmail.users.messages)`');
        }

        const id = params.auth.id;

        if (!(params.auth instanceof OAuth2)) {
            const oauth2Client = new OAuth2(
                this.clientId,
                this.clientSecret
            );

            oauth2Client.setCredentials({
                access_token: params.auth.access_token,
                refresh_token: params.auth.refresh_token
            });

            params.auth = oauth2Client;
        }

        const oldAccessToken = params.auth.credentials.access_token;

        return method(params, (err, res) => {

            if (params.auth.credentials.access_token !== oldAccessToken) {
                this.emit('newAccessToken', { ...params.auth.credentials, id });
            }

            if (err) {
                return callback(Boom.boomify(err, { statusCode: err.code ? Number(err.code) : 500 }));
            }

            return callback(null, res.data);
        });
    }
}

/* Internal utility functions */

internals.getAddressesObject = (mail) => {

    return {
        to: mail.to && internals.getAddressesObjectFromValue(mail.to.value),
        from: (mail.from && internals.getAddressesObjectFromValue(mail.from.value)[0]) || {},
        cc: mail.cc && internals.getAddressesObjectFromValue(mail.cc.value),
        bcc: mail.bcc && internals.getAddressesObjectFromValue(mail.bcc.value)
    };
};

/**
 * @param {{ address: String, name: String }[]} value
 * @returns {Array.<{ name: String, email: String }> | undefined}
 */
internals.getAddressesObjectFromValue = (value) => {

    return value && value.map((vObject) => {

        if (!vObject.address) {
            return;
        }

        return {
            name: vObject.name,
            email: vObject.address.toLowerCase()
        };
    }).filter((address) => !!address);
};

/**
 * @param {{ name: String, value: String }[]} headers
 * @param {String} name
 * @returns {null | String}
 */
internals.getHeaderValue = (headers, name) => {

    const headerObject = headers.find((header) => {

        return header.name.toLowerCase() === name.toLowerCase();
    });

    if (headerObject && headerObject.value) {
        return headerObject.value;
    }

    return null;
};

/**
 * @param {MessageRecipient[] | MessageRecipient} recipients
 * @returns {'' | String} When recipients.length === 0 returns '' otherwise csv String of recipients
 */
internals.convertMessageRecipientsToCsv = (recipients) => {

    if (!Array.isArray(recipients)) {
        recipients = [recipients];
    }

    recipients = recipients.filter((x) => !!x);

    if (recipients.length === 0) {
        return '';
    }

    return recipients.map((recipient) => {

        if (recipient.name) {
            return `"${recipient.name.replace(/"/g, '""')}" <${recipient.email}>`;
        }

        return recipient.email;
    }).join(', ');
};

module.exports = GmailConnector;
