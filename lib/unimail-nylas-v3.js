'use strict';

const EventEmitter = require('events');
const Async = require('async');

const Nylas = require('nylas').default;
const Boom = require('@hapi/boom');

const internals = {
    folderMap: {
        archive: '\\Archive',
        drafts: '\\Drafts',
        inbox: '\\Inbox',
        junk: '\\Junk',
        sent: '\\Sent',
        trash: '\\Trash'
    }
};

/**
 * @typedef {import('./index').MessageResource} MessageResource
 * @typedef {import('./index').MessageListResource} MessageListResource
 * @typedef {import('./index').FileResource} FileResource
 * @typedef {import('./index').FileListResource} FileListResource
 * @typedef {import('./index').MessageRecipient} MessageRecipient
 */

class NylasV3Connector extends EventEmitter {

    /**
     * @class
     *
     * @param {Object} config - Configuration object
     * @param {String} config.clientId
     * @param {String} config.clientSecret
     */
    constructor(config) {

        super();

        this.apiKey = config.apiKey;
        this.apiUri = config.apiUri;
        this.clientId = config.clientId;

        this.name = 'nylas-v3';

        const nylasConfig = {
            apiKey: this.apiKey,
            apiUri: this.apiUri,
            clientId: this.clientId
        };

        internals.nylas = new Nylas(nylasConfig);

        internals.nylas.applications.getDetails({
            redirectUris: [config.callbackUri]
        });
        // Response ?
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
     * @param {
         String
     } [params.rfc2822Format = false] - Return the email in rfc2822 format https: //www.ietf.org/rfc/rfc2822.txt
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

        return internals.nylas.messages.find({
            identifier: auth.access_token,
            messageId: encodeURIComponent(params.id),
            queryParams: { fields: 'include_headers' }
        }).then((response) => {

            if (options.raw) {
                return callback(null, response.data);
            }

            if (params.rfc2822Format) {
                return callback(null, internals.generateMIMEMessage(response.data));
            }

            const message = response.data;

            const messageIdHeader = message.headers.find((header) => header.name === 'Message-Id');
            message.email_message_id = messageIdHeader ? messageIdHeader.value : null;

            const inReplyToHeader = message.headers.find((header) => header.name === 'In-Reply-To');

            if (inReplyToHeader && inReplyToHeader.value?.length > 0) {
                message.in_reply_to = inReplyToHeader.value;
            }
            else {
                message.in_reply_to = null;
            }

            if (message.attachments && message.attachments.length > 0) {
                message.files = message.attachments.map((file) => {

                    const fileObject = {
                        message,
                        metadata: file
                    };

                    return this._transformFiles(fileObject)[0];
                }).filter((x) => !!x);
            }
            else {
                message.files = [];
            }

            return internals.getFolders(auth, (err, folders) => {

                if (err) {
                    return callback(err);
                }

                return callback(null, this._transformMessages(message, folders)[0]);
            });
        }).catch((err) => {

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
     * @param {String[]} params.folder - Only return messages in these folders
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
    // TODO: test any_email without after and with limit and offset
    listMessages(auth, params, options, callback) {

        if (typeof options === 'function') {
            callback = options;
            options = {};
        }

        options = options || {};
        const clonedParams = { ...params };

        if (clonedParams.folder) {
            try {
                return internals.getFolders(auth, (err, folders) => {

                    if (err) {
                        return callback(err);
                    }

                    // Filter out the folder with the correct attribute
                    const nylasFolder = folders.find((folder) => folder.attributes.includes(internals.folderMap[clonedParams.folder]));

                    if (!nylasFolder) {
                        const err = new Error(`Folder ${clonedParams.folder} not found`);
                        err.statusCode = 404;
                        err.folders = folders;
                        return callback(err);
                    }

                    clonedParams.folder = nylasFolder.id;

                    return this.listMessagesCall(auth, clonedParams, options, callback);
                });
            }
            catch (err) {
                return callback(Boom.boomify(err, { statusCode: 500 }));
            }
        }

        return this.listMessagesCall(auth, clonedParams, options, callback);
    }

    listMessagesCall(auth, clonedParams, options, callback) {

        let participant;

        if (Array.isArray(clonedParams.participants)) {
            if (clonedParams.participants.length > 25) {
                const batchSize = 25;
                const batches = Math.ceil(clonedParams.participants.length / batchSize);
                const batchOptions = [];
                for (let i = 0; i < batches; ++i) {
                    const batchParticipants = clonedParams.participants.slice(i * batchSize, (i + 1) * batchSize);
                    const batchParams = { ...clonedParams };
                    batchParams.participants = batchParticipants;

                    batchOptions.push(batchParams);
                }

                return Async.mapLimit(batchOptions, 5, (batchParams, callback) => {

                    return this.listMessages(auth, batchParams, options, callback);
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

            participant = clonedParams.participants.join(',');
        }

        // The official Nylas Node SDK docs say the parameters should be camelcase, but snakecase works as well. During development, camelcase was used, so we're leaving it like this.
        const nylasParams = {
            limit: clonedParams.limit,
            received_before: clonedParams.before ? Math.floor(new Date(clonedParams.before) / 1000) : undefined,
            received_after: clonedParams.after ? Math.floor(new Date(clonedParams.after) / 1000) : undefined,
            to: clonedParams.to,
            from: clonedParams.from,
            subject: clonedParams.subject,
            in: clonedParams.folder,
            has_attachment: clonedParams.hasAttachment ? true : (clonedParams.hasAttachment === false ? false : undefined),
            any_email: participant,
            page_token: clonedParams.pageToken
        };

        if (clonedParams.idsOnly) {
            nylasParams.select = 'id';
        }
        else {
            nylasParams.fields = 'include_headers';
        }

        // TODO: implement includeDrafts

        // Remove undefined values since Nylas API doesn't like them
        const definedParams = Object.fromEntries(
            Object.entries(nylasParams).filter(([, value]) => value !== undefined)
        );

        try {
            return internals.nylas.messages.list({
                identifier: auth.access_token,
                queryParams: definedParams
            }).then((response) => {

                const responseObject = {
                    messages: response.data
                };

                if (response.nextCursor) {
                    responseObject.next_page_token = response.nextCursor;
                }

                if (options.raw || options.idsOnly) {
                    return callback(null, responseObject);
                }

                return internals.getFolders(auth, (err, folders) => {

                    if (err) {
                        return callback(err);
                    }

                    responseObject.messages = responseObject.messages.map((message) => {

                        const messageIdHeader = message.headers.find((header) => header.name === 'Message-Id');
                        message.email_message_id = messageIdHeader ? messageIdHeader.value : null;

                        const inReplyToHeader = message.headers.find((header) => header.name === 'In-Reply-To');

                        if (inReplyToHeader && inReplyToHeader.value?.length > 0) {
                            message.in_reply_to = inReplyToHeader.value;
                        }
                        else {
                            message.in_reply_to = null;
                        }

                        if (message.attachments && message.attachments.length > 0) {
                            message.files = message.attachments.map((file) => {

                                const fileObject = {
                                    message,
                                    metadata: file
                                };

                                return this._transformFiles(fileObject)[0];
                            }).filter((x) => !!x);
                        }
                        else {
                            message.files = [];
                        }

                        return this._transformMessages(message, folders)[0];
                    });
                    return callback(null, responseObject);
                });
            }).catch((err) => {

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
        catch (err) {

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
        }
    }

    sendMessage(auth, params, options, callback) {

        const err = new Error('Not implemented');
        return callback(err);
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

                return callback(null, { files, next_page_token: response.nextCursor });
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

        if (!params || !params.id || !params.messageId) {
            throw new Error('Invalid configuration. Please refer to the documentation to get the required fields.');
        }

        if (typeof options === 'function') {
            callback = options;
            options = {};
        }

        options = options || {};

        const fileObject = {};

        return internals.nylas.attachments.find({
            identifier: auth.access_token,
            attachmentId: params.id,
            queryParams: {
                messageId: encodeURIComponent(params.messageId)
            }
        }).then((response) => {

            fileObject.metadata = response.data;

            return internals.nylas.attachments.download({
                identifier: auth.access_token,
                attachmentId: params.id,
                queryParams: {
                    messageId: encodeURIComponent(params.messageId)
                }
            }).then((attachmentResponse) => {

                const chunks = [];

                attachmentResponse.on('data', (chunk) => chunks.push(chunk));
                attachmentResponse.on('end', () => {

                    fileObject.download = Buffer.concat(chunks);

                    return internals.nylas.messages.find({
                        identifier: auth.access_token,
                        messageId: encodeURIComponent(params.messageId)
                    }).then((messageResponse) => {

                        fileObject.message = messageResponse.data;

                        if (options.raw) {
                            return callback(null, fileObject);
                        }

                        return callback(null, this._transformFiles(fileObject)[0]);
                    }).catch((err) => {

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
                });
            }).catch((err) => {

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
        }).catch((err) => {

            return callback(Boom.boomify(err, { statusCode: err.message.includes('Couldn\'t find') ? 404 : 500 }));
        });
    }


    /* TRANSFORMERS */

    /**
     * Transforms a raw Nylas API messages response to a unified message resource
     *
     * @param {Object[]} messagesArray - Array of messages in the format returned by the Nylas API
     * @param {Object[]} folders - Array of folders in the format returned by the Nylas API
     * @returns {Array.<MessageResource>} - Array of unified message resources
     */
    _transformMessages(messagesArray, folders) {

        if (!messagesArray || messagesArray.length === 0) {
            return [];
        }

        const messages = Array.isArray(messagesArray) ? messagesArray : [messagesArray];

        return messages.map((message) => {

            const formattedMessage = {
                email_message_id: message.email_message_id,
                service_message_id: message.id,
                service_thread_id: message.threadId,
                date: Number(message.date * 1000),
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

            if (message.folders) {
                formattedMessage.folders = message.folders.map((folder) => {

                    const messageFolder = folders.find((f) => f.id === folder);
                    const commonName = Object.values(internals.folderMap).find((value) => messageFolder.attributes.includes(value));
                    if (commonName) {
                        const folderName = commonName.replace('\\', '');
                        return folderName;
                    }

                    return messageFolder.name;
                });
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
     * TODO: test inline attachments
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
                email_message_id: file.message.id,
                file_name: file.metadata.filename,
                service_file_id: file.metadata.id,
                content_disposition: file.metadata.contentDisposition,
                content_id: file.metadata.contentId || file.metadata.content_id,
                service_type: this.name,
                is_embedded: file.metadata.isInline,
                addresses: internals.getAddressesObject(file.message),
                date: file.message.date * 1000
            };

            if (file.download) {
                formattedFile.data = file.download.toString('base64');
            }

            if (file.metadata.contentId) {
                formattedFile.content_id = file.metadata.contentId.replace(/[<>]/g, '');
            }

            if (file.message.body.includes(`cid:${  formattedFile.content_id}`)) {
                formattedFile.is_embedded = true;
                if (formattedFile.content_disposition) {
                    formattedFile.content_disposition = formattedFile.content_disposition.replace('attachment;', 'inline;');
                }
                else {
                    formattedFile.content_disposition = 'inline';
                }
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

/**
 *
 * @param {Object} message
 * @param {Array.<{ email: String, name: String }>} message.from
 * @param {Array.<{ email: String, name: String }>} message.to
 * @param {Array.<{ email: String, name: String }>} message.cc
 * @param {Array.<{ email: String, name: String }>} message.bcc
 * @param {String} message.subject
 * @param {String} message.date
 * @param {String} message.messageId
 * @param {Object} message.headers
 * @param {String} message.body
 * @param {Array.<{ contentType: String, filename: String, data: String }>} message.attachments
 *
 * @returns {String}
 */
internals.generateMIMEMessage = (message) => {

    let mimeMessage = '';

    // Add headers
    mimeMessage += `From: ${message.from[0].name ? `${message.from[0].name} <${message.from[0].email}>` : message.from[0].email}\n`;
    mimeMessage += `To: ${message.to.map((recipient) => (recipient.name ? `${recipient.name} <${recipient.email}>` : recipient.email)).join(', ')}\n`;

    if (message.cc?.length > 0) {
        mimeMessage += `Cc: ${message.cc.map((recipient) => (recipient.name ? `${recipient.name} <${recipient.email}>` : recipient.email)).join(', ')}\n`;
    }

    if (message.bcc?.length > 0) {
        mimeMessage += `Bcc: ${message.bcc.map((recipient) => (recipient.name ? `${recipient.name} <${recipient.email}>` : recipient.email)).join(', ')}\n`;
    }

    mimeMessage += `Subject: ${message.subject}\n`;
    mimeMessage += `Date: ${new Date(message.date * 1000).toISOString()}\n`;
    mimeMessage += `Message-ID: ${message.id}\n`;

    for (const header of message.headers) {
        if (!['From', 'To', 'Cc', 'Bcc', 'Subject', 'Date', 'Message-ID'].includes(header.name)) {
            mimeMessage += `${header.name}: ${header.value}\n`;
        }
    }

    mimeMessage += 'MIME-Version: 1.0\n';
    mimeMessage += 'Content-Type: multipart/mixed; boundary="boundary-example"\n\n';

    // Add body
    mimeMessage += '--boundary-example\n';
    mimeMessage += 'Content-Type: text/plain; charset="UTF-8"\n';
    mimeMessage += 'Content-Transfer-Encoding: quoted-printable\n\n';
    mimeMessage += `${message.body}\n\n`;

    // Add attachments
    for (const attachment of message.attachments) {
        mimeMessage += '--boundary-example\n';
        mimeMessage += `Content-Type: ${attachment.contentType}\n`;
        mimeMessage += 'Content-Transfer-Encoding: base64\n';
        mimeMessage += `Content-Disposition: attachment; filename="${attachment.filename}"\n\n`;
        mimeMessage += `${attachment.data}\n\n`;
    }

    mimeMessage += '--boundary-example--';

    return mimeMessage;
};

internals.getFolders = (auth, callback) => {

    return internals.nylas.folders.list({
        identifier: auth.access_token,
        queryParams: { limit: 200 }
    }).then((folderResponse) => {

        const folders = folderResponse.data;
        const inboxFolder = folders.find((folder) => folder.attributes.find((attribute) => attribute.toLowerCase().includes('inbox')));
        const sentFolder = folders.find((folder) => folder.attributes.find((attribute) => attribute.toLowerCase().includes('sent')));

        if (inboxFolder && sentFolder) {
            return callback(null, folders);
        }

        if (!inboxFolder) {
            const inboxes = folders.filter((folder) => folder.name.replaceAll('INBOX.', '').toLowerCase().includes('inbox'));
            if (inboxes.length === 0) {
                return callback(Boom.notFound('Inbox folder not found'));
            }

            let inbox;
            if (inboxes.length > 1) {
                // Get the folder in the results with the highest totalCount
                for ( const folder of inboxes) {
                    if (!inbox || folder.totalCount > inbox.totalCount) {
                        inbox = folder;
                    }
                }
            }
            else {
                inbox = inboxes[0];
            }

            inbox.attributes.push('\\Inbox');
        }

        if (!sentFolder) {
            const sentFolders = folders.filter((folder) => folder.name.replaceAll('INBOX.', '').toLowerCase().includes('sent'));
            if (sentFolders.length === 0) {
                return callback(Boom.notFound('Sent folder not found'));
            }

            let sent;
            if (sentFolders.length > 1) {
                // Get the folder in the results with the highest totalCount
                for ( const folder of sentFolders) {
                    if (!sent || folder.totalCount > sent.totalCount) {
                        sent = folder;
                    }
                }
            }
            else {
                sent = sentFolders[0];
            }

            sent.attributes.push('\\Sent');
        }

        return callback(null, folders);
    }).catch((err) => {

        return callback(Boom.boomify(err, { statusCode: 500 }));
    });

};


module.exports = NylasV3Connector;
