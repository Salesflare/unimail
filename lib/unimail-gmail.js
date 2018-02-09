const EventEmitter = require('events');

const Google = require('googleapis');
const Gmail = Google.gmail('v1');
const OAuth2 = Google.auth.OAuth2;
const Batchelor = require('batchelor');
const DateFns = require('date-fns');
const Boom = require('boom');

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

            return callback(null, this.transformMessages(response)[0]);
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
        let q = '';

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
            q += `before:${DateFns.format(params.before, 'YYYY-MM-DD')} `;
        }

        if (params.after) {
            q += `after:${DateFns.format(params.after, 'YYYY-MM-DD')} `;
        }

        if (params.participants) {
            q += `{${params.participants.map((participant) => {

                return `from:${participant} to:${participant} `;
            }).join('')}}`;
        }

        if (params.from) {
            q += `from:${params.from} `;
        }

        if (params.to) {
            q += `to:${params.to} `;
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
                return callback(null, []);
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

                const responseObject = {
                    messages: batchResponse.parts.map((messagePart) => {

                        if (options.raw) {
                            return messagePart.body;
                        }

                        return this.transformMessages(messagePart.body)[0];
                    })
                };

                if (listResponse.nextPageToken) {
                    responseObject.nextPageToken = listResponse.nextPageToken;
                }

                return callback(null, responseObject);
            });
        });
    }

    /**
     * Receive notifications of a user's inbox. Requires an active PubSub topic. See https://developers.google.com/gmail/api/v1/reference/users/watch
     *
     * @param {Object} auth - Authentication object
     * @param {string} auth.access_token - Access token
     * @param {string} auth.refresh_token - Refresh token
     *
     * @param {Object} params
     * @param {string} params.topicName - Topic name format should follow 'projects/my-project-id/topics/my-topic-id'
     * @param {string[]} params.labelIds - List of label_ids to restrict notifications about
     * @param {string} params.labelFilterAction - Filtering behavior of labelIds list specified can be 'include' or 'exclude'
     */
    watchMessages(auth, params, callback) {

        if (!params || !params.topicName) {
            throw new Error('Invalid configuration. Please refer to the documentation to get the required fields.');
        }

        const gmailParams = {
            auth,
            userId: 'me',
            resource: {
                topicName: params.topicName
            }
        };

        if (params.labelIds) {
            gmailParams.labelIds = params.labelIds;
        }

        if (params.labelFilterAction) {
            gmailParams.labelFilterAction = params.labelFilterAction;
        }

        return this._callAPI(Gmail.users.watch, gmailParams, (err, response) => {

            return callback(err, response);
        });
    };

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
                return callback(null, []);
            }

            if (options.raw) {
                return callback(null, messagesList.messages);
            }

            const fileListObject = {
                files: this.transformFiles(messagesList.messages)
            };

            if (messagesList.nextPageToken) {
                fileListObject.next_page_token = messagesList.nextPageToken;
            }

            return callback(null, fileListObject);
        })
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

        const fileParams = {
            auth,
            userId: 'me',
            id: params.id,
            messageId: params.messageId
        };

        return this._callAPI(Gmail.users.messages.get, messageParams, (err, messageResponse) => {

            if (err) {
                return callback(err);
            }

            const nonContainerParts = this._extractNonContainerParts(messageResponse.payload);

            const part = nonContainerParts.filter((part) => {

                return part.body.attachmentId === fileParams.id;
            });

            //return callback(null, this.transformMessages(response)[0]);
            return this._callAPI(Gmail.users.messages.attachments.get, fileParams, (err, fileResponse) => {

                if (err) {
                    return callback(err);
                }

                fileRespone.file_id = fileResponse.attachmentId;
                fileResponse.type = part.mimeType;
                fileResponse.file_name = part.filename;

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
    transformFiles(messagesArray) {

        const messages = Array.isArray(messagesArray) ? messagesArray : [messagesArray];
        const files = [];

        messages.forEach((message) => {

            const nonContainerParts = this._extractNonContainerParts(message.payload);

            nonContainerParts.forEach((part) => {

                if (!part.filename) {
                    return;
                }

                const contentDisposition = internals.getHeaderValue(part.headers, 'content-disposition');

                files.push({
                    service_type: this.name,
                    type: part.mimeType,
                    size: Number(part.body.size),
                    service_message_id: message.id,
                    service_thread_id: message.threadId,
                    email_message_id: internals.getHeaderValue(message.payload.headers, 'message-id'),
                    subject: internals.getHeaderValue(message.payload.headers, 'subject'),
                    date: Number(message.internalDate),
                    addresses: internals.getAddressesObject(message.payload.headers),
                    file_name: part.filename,
                    content_id: internals.getHeaderValue(part.headers, 'content-id'),
                    content_disposition: contentDisposition,
                    file_id: part.body.attachmentId,
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
    transformMessages(messagesArray) {

        const messages = Array.isArray(messagesArray) ? messagesArray : [messagesArray];

        return messages.map((message) => {

            const formattedMessage = {
                service_name: this.name,
                email_message_id: internals.getHeaderValue(message.payload.headers, 'message-id'),
                service_message_id: message.id,
                service_thread_id: message.threadId,
                date: Number(message.internalDate),
                subject: internals.getHeaderValue(message.payload.headers, 'subject'),
                folders: message.labelIds,
                files: [],
                body: [],
                addresses: internals.getAddressesObject(message.payload.headers),
                in_reply_to: internals.getHeaderValue(message.payload.headers, 'in-reply-to')
            };

            if (message.payload && message.payload.headers) {
                const headerObject = {};

                message.payload.headers.forEach((header) => {

                    if (!headerObject[header.name]) {
                        headerObject[header.name] = [header.value];
                    }
                    else {
                        headerObject[header.name].push(header.value);
                    }
                });

                formattedMessage.headers = headerObject;
            }

            const nonContainerParts = this._extractNonContainerParts(message.payload);
            const allowedBodyMimeTypes = ['text/html', 'text/plain'];

            nonContainerParts.forEach((nonContainerPart) => {

                if (allowedBodyMimeTypes.includes(nonContainerPart.mimeType.toLowerCase()) && nonContainerPart.body.data) {
                    formattedMessage.body.push({
                        type: nonContainerPart.mimeType,
                        content: Buffer.from(nonContainerPart.body.data, 'base64').toString()
                    })
                }

                if (nonContainerPart.filename && nonContainerPart.headers) {
                    const contentDisposition = internals.getHeaderValue(nonContainerPart.headers, 'content-disposition');

                    formattedMessage.files.push({
                        type: nonContainerPart.mimeType,
                        size: Number(nonContainerPart.body.size),
                        service_message_id: message.id,
                        service_thread_id: message.threadId,
                        email_message_id: internals.getHeaderValue(message.payload.headers, 'message-id'),
                        subject: internals.getHeaderValue(message.payload.headers, 'subject'),
                        date: message.internalDate,
                        addresses: internals.getAddressesObject(message.payload.headers),
                        file_name: nonContainerPart.filename,
                        content_id: internals.getHeaderValue(nonContainerPart.headers, 'content-id'),
                        content_disposition: contentDisposition,
                        file_id: nonContainerPart.body.attachmentId,
                        is_embedded: contentDisposition ? contentDisposition.startsWith('inline;') : false
                    });
                }
            });

            return formattedMessage;
        });
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

            this.emit('newAccessToken', { ...token });

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
    };
}

/* Internal utility functions */

internals.getAddressesObject = (headers) => {

    const fromHeaderValue = internals.getHeaderValue(headers, 'from');
    const toHeaderValue = internals.getHeaderValue(headers, 'to');
    const ccHeaderValue = internals.getHeaderValue(headers, 'cc');
    const bccHeaderValue = internals.getHeaderValue(headers, 'bcc');
    let parsedFromHeader;
    let parsedToHeader;
    let parsedCcHeader;
    let parsedBccHeader;

    if (fromHeaderValue) {
        parsedFromHeader = internals.parseEmailToAndFromHeaders(fromHeaderValue);
    }

    if (toHeaderValue) {
        parsedToHeader = internals.parseEmailToAndFromHeaders(toHeaderValue);
    }

    if (ccHeaderValue) {
        parsedCcHeader = internals.parseEmailToAndFromHeaders(ccHeaderValue);
    }

    if (bccHeaderValue) {
        parsedBccHeader = internals.parseEmailToAndFromHeaders(bccHeaderValue);
    }

    const addressesObject = {
        from: parsedFromHeader && parsedFromHeader.length > 0 ? parsedFromHeader[0] : {},
        to: parsedToHeader && parsedToHeader.length > 0 ? parsedToHeader : []
    };

    if (parsedCcHeader && parsedCcHeader.length > 0) {
        addressesObject.cc = parsedCcHeader;
    }

    if (parsedBccHeader && parsedBccHeader.length > 0) {
        addressesObject.bcc = parsedBccHeader;
    }

    return addressesObject;
};

internals.getHeaderValue = (headers, name) => {

    const headerObject = headers.find((header) => {

        return header.name.toLowerCase() === name.toLowerCase();
    });

    if (headerObject && headerObject.value) {
        return headerObject.value;
    }

    return null;
};

internals.parseEmailToAndFromHeaders = (recipientString) => {

    const regex = /(([\w,\"\s]+)\s)?<?([^@<\s]+@[^@\s>]+)>?,?/g;
    const recipientsArray = [];
    let m;

    // Traverse all matches in the global regex
    while ((m = regex.exec(recipientString)) !== null) {
        // This is necessary to avoid infinite loops with zero-width matches
        if (m.index === regex.lastIndex) {
            regex.lastIndex++;
        }

        let name ;
        let email;

        if (m[2]) {
            name = m[2].replace(/,$/, '').replace(/"/g, "").trim(); // Strip whitespaces and commas, and remove quotation marks
        }

        if (m[3]) {
            email = m[3].replace(/,$/, '').trim(); // Strip whitespaces and commas from end of string
        }

        const recipientObject = {
            email
        };

        if (name) {
            recipientObject.name = name;
        }

        recipientsArray.push(recipientObject)
    }

    return recipientsArray;
};

module.exports = GmailConnector;
