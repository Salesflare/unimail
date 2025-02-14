'use strict';

const EventEmitter = require('events');
const Async = require('async');

const { UnipileClient } = require('unipile-node-sdk');
const Boom = require('@hapi/boom');

const internals = {};

/**
 * @typedef {import('./index').MessageResource} MessageResource
 * @typedef {import('./index').MessageListResource} MessageListResource
 * @typedef {import('./index').FileResource} FileResource
 * @typedef {import('./index').FileListResource} FileListResource
 * @typedef {import('./index').MessageRecipient} MessageRecipient
 */

class UnipileConnector extends EventEmitter {

    /**
     * @class
     *
     * @param {Object} config - Configuration object
     * @param {String} config.baseUrl - Base URL of the Unipile API
     * @param {String} config.accessToken - Access token for the Unipile API
     */
    constructor(config) {

        super();

        this.baseUrl = config.baseUrl;
        this.accessToken = config.accessToken;
        this.name = 'unipile';

        internals.client = new UnipileClient(this.baseUrl, this.accessToken);
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
     * @param {String} params.id - unipile message id
     * @param {String} [params.rfc2822Format = false] - Return the email in rfc2822 format https: //www.ietf.org/rfc/rfc2822.txt
     *
     * @param {Object} options
     * @param {Boolean} [options.raw] - If true the response will not be transformed to the unified object
     *
     * @param {function(Error?, (MessageResource | Object | String)?):void} callback Returns a unified message resource when options.raw is falsy or the raw response of the API when truthy
     * @returns {void}
     *
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

        return internals.client.email.getOne(params.id, { extra_params: { include_headers: true } }).then((response) => {

            if (options.raw) {
                return callback(null, response);
            }

            if (params.rfc2822Format) {
                return callback(null, internals.generateMIMEMessage(response));
            }

            const message = response;

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

            return callback(null, this._transformMessages(message)[0]);
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
     * @param {Boolean} params.hasAttachment - If true, only return messages with attachments (! TODO: NOT SUPPORTED BY UNIPILE, a workaround is implemented, but it's not very performant)
     * @param {Date} params.before - Only return messages before this date
     * @param {Date} params.after - Only return messages after this date
     * @param {String | Number} params.pageToken - Token used to retrieve a certain page in the list
     * @param {String} params.from - Only return messages sent from this address
     * @param {String} params.to - Only return messages sent to this address
     * @param {String[]} params.participants - Array of email addresses: only return messages with at least one of these participants are involved.
     * Due to unipile api limitation the participants filter will only be applied when an 'after' filter is applied and limit and offset will be ignored
     * @param {String[]} params.folder - Only return messages in these folders
     * @param {Boolean} params.includeDrafts - Whether to include drafts or not, defaults to false
     * @param {String} params.subject - ! TODO: SUBJECT NOT SUPPORTED this is currently not an issue, since it's only used for getting messages by email message id
     * in practice the date param included in that call will be enough to get the correct message
     *
     * @param {Object} options
     * @param {Boolean} [options.raw] - If true the response will not be transformed to the unified object
     * @param {Boolean} [options.idsOnly] - If true the response will only contain the ids of the messages
     * @param {Boolean} [options.includeBody] - Defaults to true, note that this may slow down calls for some IMAP servers
     *
     * @param {function(Error?, MessageListResource | { messages: Array.<Object | String>, next_page_token: String? }?):void} callback Returns an array of unified message resources when options.raw is falsy or the raw response of the API when truthy
     * @returns {void}
     */
    listMessages(auth, params, options, callback) {

        if (typeof options === 'function') {
            callback = options;
            options = {};
        }

        options = options || {};
        options = { includeBody: true, ...options };
        const clonedParams = { ...params };

        // Check if the pageToken is a date object or a string
        // If there's already an after filter, use that instead if that date is newer
        // We do this to replace the page token, since fetching more messages breaks the Unipile one
        if (
            internals.isDateObjectOrString(clonedParams.pageToken) &&
            (!clonedParams.after || new Date(clonedParams.pageToken) > new Date(clonedParams.after))
        ) {
            clonedParams.after = clonedParams.pageToken;
            delete clonedParams.pageToken;
        }

        if (!clonedParams.limit) {
            clonedParams.limit = 20;
        }

        // When looking for files, take a larger minimum limit
        // But keep track of the original limit to apply it again later
        // This will cause page tokens to be incorrect, but it's the best we can do right now
        if (clonedParams.hasAttachment && clonedParams.limit < 100) {
            options.originalAttachmentLimit = clonedParams.limit;
            clonedParams.limit = 100;
        }

        // If a folder is specified, get the folder id and call the listMessages method for each folder separately
        // Multiple folders at the same time are not supported by the unipile API
        if (clonedParams.folder) {
            try {
                return internals.getFolders(auth, (err, folders) => {

                    if (err) {
                        return callback(err);
                    }

                    // Filter out the folder with the correct attribute
                    const unipileFolders = folders.filter((folder) => folder.role.includes(clonedParams.folder));

                    if (!unipileFolders) {
                        const err = new Error(`Folder ${clonedParams.folder} not found`);
                        err.statusCode = 404;
                        err.folders = folders;
                        return callback(err);
                    }

                    return Async.map(unipileFolders, (unipileFolder, asyncCallback) => {

                        const folderCallParams = { ...clonedParams };
                        folderCallParams.folder = unipileFolder.provider_id;

                        return this.listMessagesCall(auth, folderCallParams, options, asyncCallback);
                    }, (err, results) => {

                        if (err) {
                            return callback(err);
                        }

                        return callback(null, { messages: results.flatMap((result) => result.messages) });
                    });
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
            participant = clonedParams.participants.join(',');
        }
        else {
            participant = clonedParams.participants;
        }

        // TODO: SUBJECT NOT SUPPORTED this is currently not an issue, since it's only used for getting messages by email message id
        // in practice the date param included in that call will be enough to get the correct message

        const input = {
            account_id: auth.access_token,
            ...(clonedParams.limit && { limit: clonedParams.limit }),
            ...(clonedParams.before && { before: new Date(clonedParams.before).toISOString() }),
            ...(clonedParams.after && { after: new Date(clonedParams.after).toISOString() }),
            ...(clonedParams.to && { to: clonedParams.to }),
            ...(clonedParams.from && { from: clonedParams.from }),
            ...(clonedParams.folder && { folder: clonedParams.folder }),
            ...(participant && { any_email: participant }),
            ...(clonedParams.pageToken && { cursor: clonedParams.pageToken })
        };

        const unipileOptions = {};

        if (!options.idsOnly && options.includeBody) {
            unipileOptions.extra_params = { include_headers: true  };
        }
        else {
            unipileOptions.extra_params = { meta_only: true };
        }

        try {
            return internals.client.email.getAll(input, unipileOptions).then((response) => {

                // Remove trash, spam, draft folder results
                let items = internals.removeUnwantedFolderResults(clonedParams.includeDrafts, response.items);

                const limit = options.nested ? options.nested_limit : options.originalAttachmentLimit || clonedParams.limit;

                // Workaround for the hasAttachment filter
                // Get all messages and filter out the ones without attachments
                // TODO: Remove once Unipile implements the hasAttachment parameter
                if (clonedParams.hasAttachment) {
                    items = items.filter((message) => message.attachments.length > 0);

                    if (items.length < limit) {
                        clonedParams.pageToken = response.cursor;
                        const clonedOptions = { ...options };
                        clonedOptions.nested = true;
                        clonedOptions.nested_limit = limit - items.length;

                        return this.listMessagesCall(auth, clonedParams, clonedOptions, (err, nestedResponse) => {

                            if (err) {
                                return callback(err);
                            }

                            items = [...items, ...nestedResponse.items.filter((message) => message.attachments.length > 0)];

                            if (options.nested) {
                                return callback(null, { items });
                            }

                            let lastEmailDate;
                            if (items.length > options.originalAttachmentLimit || clonedParams.limit) {
                                lastEmailDate = items[(options.originalAttachmentLimit || clonedParams.limit) - 1].date;
                            }

                            // Apply original limit again
                            items = items.slice(0, options.originalAttachmentLimit || clonedParams.limit);

                            const responseObject = {
                                items
                            };

                            if (lastEmailDate) {
                                responseObject.cursor = lastEmailDate;
                            }

                            return this._listMessageResponseHandler(responseObject, options, callback);
                        });

                    }
                    else if (options.nested) {
                        return callback(null, { items });
                    }
                }

                if (response.cursor && items.length < limit) {
                    clonedParams.pageToken = response.cursor;
                    const clonedOptions = { ...options };
                    clonedOptions.nested = true;
                    clonedOptions.nested_limit = limit - items.length;

                    return this.listMessagesCall(auth, clonedParams, clonedOptions, (err, nestedResponse) => {

                        if (err) {
                            return callback(err);
                        }

                        items = [...items, ...nestedResponse.items];

                        if (options.nested) {
                            return callback(null, { items });
                        }

                        let lastEmailDate;
                        if (items.length > options.originalAttachmentLimit || clonedParams.limit) {
                            lastEmailDate = items[(options.originalAttachmentLimit || clonedParams.limit) - 1].date;
                        }

                        // Apply original limit again
                        items = items.slice(0, options.originalAttachmentLimit || clonedParams.limit);

                        const responseObject = {
                            items
                        };

                        if (lastEmailDate) {
                            responseObject.cursor = lastEmailDate;
                        }

                        return this._listMessageResponseHandler(responseObject, options, callback);
                    });
                }
                else if (options.nested) {
                    return callback(null, { items });
                }

                return this._listMessageResponseHandler(response, options, callback);

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
     * @param {Object} params - same as listMessages
     * @param {Object} options
     * @param {Boolean} [options.raw] - If true the response will not be transformed to the unified object
     * @param {function(Error?, FileListResource | { files: Array.<Object>, next_page_token: String? }):void} callback Returns an array of unified file resources when options.raw is falsy or the raw response of the API when truthy
     *
     * @returns {void}
     */
    listFiles(auth, params, options, callback) {

        if (typeof options === 'function') {
            callback = options;
            options = {};
        }

        const clonedParams = { ...params };

        options = options || {};

        clonedParams.hasAttachment = true;

        return this.listMessages(auth, clonedParams, options, (err, response) => {

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
     * Method used for tests vs Nylas results
     *
     * @param {Object} auth
     * @param {Object} params
     * @param {Object} options
     * @param {function(Error?, Array.<Object>) : void} callback
     * @returns {void}
     */
    listFolders(auth, params, options, callback) {

        return internals.client.email.getAllFolders({ account_id: auth.access_token }).then((folderResponse) => {

            const folders = folderResponse.items;
            // Find send folder(s) and return them
            return callback(null, folders);
        }).catch((err) => {

            return callback(Boom.boomify(err, { statusCode: 500 }));
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
     * @param {String} params.id - unipile attachment id
     * @param {String} params.messageId - email message id
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

        return this.getMessage(auth, { id: params.messageId }, {}, (err, message) => {

            if (err) {
                return callback(err);
            }

            const fileObject = message.files.find((attachment) => attachment.service_file_id === params.id);

            return internals.client.email.getEmailAttachment.byProviderId({
                account_id: auth.access_token,
                attachment_id: params.id,
                email_provider_id: params.messageId
            }).then((response) => {

                return response.arrayBuffer().then((buffer) => {

                    fileObject.data = Buffer.from(buffer);

                    return callback(null, fileObject);
                }).catch((err) => {

                    return callback(Boom.boomify(err, { statusCode: 500 }));
                });
            }).catch((err) => {

                return callback(Boom.boomify(err, { statusCode: 500 }));
            });
        });
    }

    /* HELPERS */

    _listMessageResponseHandler(response, options, callback) {

        const responseObject = {
            messages: response.items
        };

        if (response.cursor) {
            responseObject.next_page_token = response.cursor;
        }

        if (options.raw) {
            return callback(null, responseObject);
        }

        if (options.idsOnly) {
            return callback(null, { ...responseObject, messages: responseObject.messages.map((message) => message.id) });
        }

        const messages = responseObject.messages.map((message) => {

            const inReplyToHeader = message.headers?.find((header) => header.name === 'In-Reply-To');

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

            return message;
        });

        responseObject.messages = this._transformMessages(messages);
        return callback(null, responseObject);
    }


    /* TRANSFORMERS */

    /**
     * Transforms a raw unipile API messages response to a unified message resource
     *
     * @param {Object[]} messagesArray - Array of messages in the format returned by the unipile API
     * @returns {Array.<MessageResource>} - Array of unified message resources
     */
    _transformMessages(messagesArray) {

        if (!messagesArray || messagesArray.length === 0) {
            return [];
        }

        const messages = Array.isArray(messagesArray) ? messagesArray : [messagesArray];

        return messages.map((message) => {

            const formattedMessage = {
                email_message_id: message.message_id,
                service_message_id: message.id,
                service_thread_id: null, // TODO: Unipile doesn't seem to support threads
                date: new Date(message.date),
                subject: message.subject,
                folders: message.folders,
                attachments: message.files.length > 0,
                addresses: internals.getAddressesObject(message),
                in_reply_to: message.in_reply_to,
                service_type: this.name,
                headers: message.headers,
                files: message.files
            };

            if (message.body || message.body_plain) {
                const bodyArray = [];

                if (message.body) {
                    bodyArray.push({ content: message.body, type: 'text/html' });
                }

                if (message.body_plain) {
                    bodyArray.push({ content: message.body_plain, type: 'text/plain' });
                }

                formattedMessage.body = bodyArray;
            }

            return formattedMessage;
        });
    }

    /**
     * Transforms a raw unipile API response to a unified file resource
     *
     * @param {Array.<Object> | Object} filesArray - Array of messages in the format returned by the unipile API
     *
     * @returns {Array.<FileResource>} - Array of unified file resources
     */
    _transformFiles(filesArray) {

        const files = Array.isArray(filesArray) ? filesArray : [filesArray];

        return files.map((file) => {

            const formattedFile = {
                type: file.metadata.mime,
                size: file.metadata.size,
                service_message_id: file.message.id || file.message.service_message_id,
                service_thread_id: file.message.threadId || file.message.service_thread_id,
                email_message_id: file.message.message_id || file.message.id,
                file_name: file.metadata.name,
                service_file_id: file.metadata.id,
                content_disposition: file.metadata.contentDisposition,
                content_id: file.metadata.cid || file.metadata.content_id,
                service_type: this.name,
                is_embedded: file.metadata.isInline,
                addresses: internals.getAddressesObject(file.message),
                date: new Date(file.message.date)
            };

            if (file.download) {
                formattedFile.data = file.download.toString('base64');
            }

            if (file.message.body.includes(`cid:${formattedFile.content_id}`)) {
                formattedFile.is_embedded = true;
                if (formattedFile.content_disposition) {
                    formattedFile.content_disposition = formattedFile.content_disposition.replace('attachment;', 'inline;');
                }
                else {
                    formattedFile.content_disposition = 'inline';
                }
            }
            else if (!formattedFile.content_disposition) {
                formattedFile.content_disposition = 'attachment';
            }

            return formattedFile;
        });
    }

    // Dummy implementation since unipile access tokens are actually just account ids
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
        from: { name: message.from_attendee.display_name, email: message.from_attendee.identifier?.toLowerCase() },
        to: message.to_attendees.map((contact) => {

            return { name: contact.display_name, email: contact.identifier?.toLowerCase() };
        }),
        cc: message.cc_attendees.map((contact) => {

            return { name: contact.display_name, email: contact.identifier?.toLowerCase() };
        }),
        bcc: message.bcc_attendees.map((contact) => {

            return { name: contact.display_name, email: contact.identifier?.toLowerCase() };
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
    mimeMessage += `From: ${message.from_attendee.display_name ? `${message.from_attendee.display_name} <${message.from_attendee.identifier}>` : message.from_attendee.identifier}\n`;
    mimeMessage += `To: ${message.to_attendees.map((recipient) => (recipient.display_name ? `${recipient.display_name} <${recipient.identifier}>` : recipient.identifier)).join(', ')}\n`;

    if (message.cc?.length > 0) {
        mimeMessage += `Cc: ${message.cc_attendees.map((recipient) => (recipient.display_name ? `${recipient.display_name} <${recipient.identifier}>` : recipient.identifier)).join(', ')}\n`;
    }

    if (message.bcc?.length > 0) {
        mimeMessage += `Bcc: ${message.bcc_attendees.map((recipient) => (recipient.display_name ? `${recipient.display_name} <${recipient.identifier}>` : recipient.identifier)).join(', ')}\n`;
    }

    mimeMessage += `Subject: ${message.subject}\n`;
    mimeMessage += `Date: ${new Date(message.date).toISOString()}\n`;
    mimeMessage += `Message-ID: ${message.message_id}\n`;

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

    return internals.client.email.getAllFolders({ account_id: auth.access_token }).then((folderResponse) => {

        const folders = folderResponse.items;
        const inboxFolders = folders.filter((folder) => folder.role === 'inbox');
        const sentFolders = folders.filter((folder) => folder.role === 'sent');

        if (inboxFolders?.length > 0 && sentFolders?.length > 0) {
            return callback(null, folders);
        }

        if (!inboxFolders || inboxFolders.length === 0) {
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

        if (!sentFolders || sentFolders.length === 0) {
            const sentFoldersGuess = folders.filter((folder) => folder.name.replaceAll('INBOX.', '').toLowerCase().includes('sent'));
            if (sentFoldersGuess.length === 0) {
                return callback(Boom.notFound('Sent folder not found'));
            }

            let sent;
            if (sentFoldersGuess.length > 1) {
                // Get the folder in the results with the highest totalCount
                for ( const folder of sentFoldersGuess) {
                    if (!sent || folder.totalCount > sent.totalCount) {
                        sent = folder;
                    }
                }
            }
            else {
                sent = sentFoldersGuess[0];
            }

            sent.attributes.push('\\Sent');
        }

        return callback(null, folders);
    }).catch((err) => {

        return callback(Boom.boomify(err, { statusCode: 500 }));
    });

};

internals.removeUnwantedFolderResults = (includeDrafts, messages) => {

    const folderRoles = ['trash', 'spam'];
    if (!includeDrafts) {
        folderRoles.push('drafts');
    }

    return messages.filter((message) => {

        return !folderRoles.includes(message.role);
    });
};

internals.isDateObjectOrString = (pageToken) => {

    // Check if pageToken is a Date object
    if (pageToken instanceof Date) {
        return true;
    }

    // Check if pageToken is a string
    if (typeof pageToken === 'string') {
        // Try to parse the string as a date
        const parsedDate = new Date(pageToken);
        // Check if the parsed date is valid
        if (!Number.isNaN(parsedDate.getTime())) {
            return true;
        }
    }

    return false;
};

module.exports = UnipileConnector;
