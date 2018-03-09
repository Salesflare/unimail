'use strict';

const EventEmitter = require('events');
const MicrosoftGraph = require("@microsoft/microsoft-graph-client").Client;
const Oauth2 = require('simple-oauth2');
const Async = require('async');
const Boom = require('boom');

const Utils = require('../lib/utils');

const internals = {};

/**
 * TODO:
 * If no mailbox data doesn't exits for an outlook account responses might be:
 *  => HTTP error: 404 / Error code: MailboxNotEnabledForRESTAPI or MailboxNotSupportedForRESTAPI / Error message: “REST API is not yet supported for this mailbox
 */

class Office365Connector extends EventEmitter {

    /**
     * @constructor
     * 
     * @param {object} config - Configuration object
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

        this.oauthCredentials = {
            client: {
                id: this.clientId,
                secret: this.clientSecret
            },
            auth: {
                tokenHost: 'https://login.microsoftonline.com',
                authorizePath: 'common/oauth2/v2.0/authorize',
                tokenPath: 'common/oauth2/v2.0/token'
            }
        };

        this.oauth2 = Oauth2.create(this.oauthCredentials);
        this.name = 'office365';
    }

    /* MESSAGES */

    /**
     * Get a single message with a certain id
     * @param {Object} auth
     * @param {string} auth.access_token
     * @param {string} auth.refresh_token
     * @param {string} auth.expiration_date - Expiration date of the access token
     *
     * 
     * @param {Object} params 
     * @param {string} params.id - the message id
     * 
     * @param {Object} options
     * @param {boolean} options.raw - If true, the response will not be transformed to the unified object
     * @param {function} callback
     * 
     * @returns {MessageResource | Object | null} Returns a unified message resource when options.raw is falsy or a raw response from the API when truthy.
     * If the message id couldn't be found by the API it was most likely a draft and null will be returned.
     */
    getMessage(auth, params, options, callback) {

        if (!internals.isValidAuthentication(auth)) {
            throw new Error('Invalid authentication. One of the needed properties is missing. Please refer to the documentation to find the required fields.');
        }

        if (!params || !params.id) {
            throw new Error('Invalid configuration. Please refer to the documentation to get the required fields.');
        }

        if (typeof options === 'function') {
            callback = options;
            options = {};
        }

        options = options || {};

        return this._refreshTokenNeeded(auth, (err, token) => {

            if (err) {
                return callback(internals.wrapError(err));
            }

            auth = token;

            const client = internals.getClient(auth.access_token);

            return client.api(`/me/messages/${params.id}`)
                .select('id', 'from', 'toRecipients', 'ccRecipients', 'bccRecipients',
                    'receivedDateTime', 'subject', 'internetMessageId', 'conversationId', 'body', 'hasAttachments', 'SingleValueExtendedProperties')
                .expand(`SingleValueExtendedProperties($filter=id eq 'String 0x7D')`)
                .get((err, resMessage) => {

                    if (err) {
                        return callback(internals.wrapError(err));
                    }

                    if (options.raw) {
                        return callback(null, resMessage);
                    }

					return this._transformMessages(resMessage, (err, transformedMessages) => {

					    if (err) {
					        return callback(err);
                        }

					    const transformedMessage = transformedMessages[0];
                        const hasInlineImg = (resMessage.body.content.indexOf('cid:') > 0);

                        if (!transformedMessage.attachments && !hasInlineImg) {
                            return callback(null, transformedMessage);
                        }

                        return client.api(`/me/messages/${transformedMessage.service_message_id}/attachments`).get((err, resFiles) => {

                            if (err) {
                                return callback(internals.wrapError(err));
                            }

                            if (resFiles.value.length === 0) {
                                transformedMessage.files = [];
                            }
                            else {
                                resFiles.value.forEach((file) => {

                                    file.messageInfo = resMessage;
                                });

                                transformedMessage.files = this._transformFiles(resFiles.value);
                            }

                            return callback(null, transformedMessage);
                        });
                    });
                });
        });
    }

    /**
     * Returns a list of messages
     * Paging when working with certain folders is not supported!
     * @param {Object} auth - Authentication object
     * @param {string} auth.access_token - Access token
     * @param {string} auth.refresh_token - Refresh token
     * @param {string} auth.expiration_date - Expiration date of the access token
     * 
     * @param {Object} params
     * @param {number} params.limit - Maximum amount of messages in response
     * @param {boolean} params.hasAttachment - If true, only return messages that have attachments
     * @param {Date} params.before - Only return messages before this date
     * @param {Date} params.after - Only return messages after this date
     * @param {string} params.pageToken - Token used to retrieve a certain page in the list
     * @param {string} params.from - Only return messages sent from this address
     * @param {string} params.to - Only return messages sent to this address
     * @param {string[]} params.participants - Array of email addresses: only return messages where at least one of these participants is involved
     * @param {string[]} params.folder - Only return messages from these folders TODO: NOT SUPPORTED!!
     * @param {boolean} params.includeDrafts - Whether to include drafts or not, defaults to false
     * 
     * @param {Object}  options
     * @param {boolean} options.raw - If true the response will not be transformed to the unified object
     * 
     * @returns {MessageListResource | Object[]} - Returns an array of unified message resources when options.raw is falsy, or the raw response of the API when truthy
     */
    listMessages(auth, params, options, callback) {

        if (!internals.isValidAuthentication(auth)) {
            throw new Error('Invalid authentication. One of the needed properties is missing. Please refer to the documentation to find the required fields.');
        }
        
        if (typeof options === 'function') {
            callback = options;
            options = {};
        }
        options = options || {};
        params.original_limit = params.limit;

        return this._refreshTokenNeeded(auth, (err, token) => {

            if (err) {
                return callback(internals.wrapError(err));
            }

            auth = token;

            const client = internals.getClient(auth.access_token);
            let uri = ``;

            if (params.pageToken) {
                if (params.folder) {
                    return callback(new Error(`Requesting messages from certain folders doesn't support paging at this moment.`));
                }

                return this._get(client, params.pageToken, (err, resMessages) => {

                    if (err) {
                        return callback(internals.wrapError(err));
                    }

                    if (options.raw) {
                        return callback(null, resMessages);
                    }

                    return this._transformMessages(resMessages.value, (err, transformedMessages) => {

                        if (err) {
                            return callback(err);
                        }

                        const messageListResource = {
                            messages: transformedMessages
                        };

                        if (resMessages['@odata.nextLink']) {
                            messageListResource.next_page_token = resMessages['@odata.nextLink'];
                        }

                        return callback(null, messageListResource);
                    });
                });
            }

            // get a big part of the uri. a callback is needed, since a call to the graph API might be needed.
            return this._getUri(client, params, (err, resUri) => {

                if (err) {
                    return callback(err);
                }

                uri = resUri;

                return this._getMessages(client, params, uri, (err, resMessages) => {

                    if (err) {
                        return callback(internals.wrapError(err));
                    }

                    if (!resMessages || !resMessages.value || resMessages.value.length === 0) {
                        return callback(null, { messages: [] });
                    }

                    if (options.raw) {
                        return callback(null, resMessages);
                    }

                    return this._transformMessages(resMessages.value, (err, transformedMessages) => {

                        if (err) {
                            return callback(err);
                        }

                        const messageListResource = {
                            messages: transformedMessages
                        };

                        // TODO: only send next page token when not searching on certain folders
                        if (resMessages['@odata.nextLink']) {
                            messageListResource.next_page_token = resMessages['@odata.nextLink'];
                        }

                        return callback(null, messageListResource);
                    });
                });
            });
        });
    }

    /* FILES */

    /**
     * @param {*} auth 
     * @param {*} params 
     * @param {*} options 
     * @param {*} callback 
     */
    listFiles(auth, params, options, callback) {

        if (!internals.isValidAuthentication(auth)) {
            throw new Error('Invalid authentication. The authentication information is incorrect or expired.');
        }

        // Set params.hasAttachment to true
        params.hasAttachment = true;

        if (typeof options === 'function') {
            callback = options;
            options = {};
        }
        options = options || {};

        return this._refreshTokenNeeded(auth, (err, token) => {

            if (err) {
                return callback(internals.wrapError(err));
            }

            auth = token;

            const client = internals.getClient(auth.access_token);
            let files = [];
            let uri = ``;

            /**
             * TODO Page tokens, UPDATE: nvm, not super imported right now.
             */
            this._getUri(client, params, (err, resUri) => {

                if (err) {
                    return callback(err);
                }

                uri = resUri;

                return client.api(`/me/messages${uri}`)
                    .select('id', 'from', 'toRecipients', 'ccRecipients', 'bccRecipients',
                        'receivedDateTime', 'subject', 'internetMessageId', 'conversationId')
                    .get((err, resMessages) => {

                        if (err) {
                            return callback(internals.wrapError(err));
                        }

                        if (!resMessages || !resMessages.value || !resMessages.value.length) {
                            return callback(null, {files: []});
                        }

                        return Async.each(resMessages.value, (message, callback) => {

                            return client.api(`/me/messages/${message.id}/attachments`).get((err, resFiles) => {

                                if (err) {
                                    return callback(err);
                                }

                                if (!options.raw) {
                                    resFiles.value.forEach((resFile) => {

                                        resFile.messageInfo = message;
                                    });
                                }

                                files = files.concat(resFiles.value);

                                return callback();
                            });
                        }, (err) => {

                            if (err) {
                                return callback(internals.wrapError(err));
                            }

                            if (options.raw) {
                                return callback(null, files);
                            }

                            return callback(null, { files: this._transformFiles(files) });
                        });
                    });
            });
        });
    }

    /**
     * @param {*} auth 
     */
    getFile(auth, params, callback) {
        
        if (!internals.isValidAuthentication(auth)) {
            throw new Error('Invalid authentication. One of the needed properties is missing. Please refer to the documentation to find the required fields.');
        }

        if (!params || !params.id || !params.messageId) {
            throw new Error('Invalid configuration. Please refer to the documentation to get the required fields.');
        }

        return this._refreshTokenNeeded(auth, (err, token) => {

            if (err) {
                return callback(internals.wrapError(err));
            }

            auth = token;

            const client = internals.getClient(auth.access_token);

            return client.api(`me/messages/${params.messageId}/attachments/${params.id}`)
                .get((err, resFile) => {

                    if (err) {
                        return callback(err);
                    }

                    return callback(null, {
                        size: resFile.size,
                        file_name: resFile.name,
                        type: resFile.contentType,
                        service_file_id: resFile.id,
                        service_message_id: params.messageId,
                        data: resFile.contentBytes,
                        file_id: resFile.contentId
                    });
                });
        });
    };

    refreshAuthCredentials(auth, callback) {

        this._refreshTokenNeeded(auth, (err, resAuth) => {

            if (err) {
                return callback(err);
            }
            auth = resAuth;
            return callback(null, auth);
        })
    }

    /**
     * 
     * INTERNAL METHODS
     * 
     */

    /**
     * Checks if drafts can be included. If not, the amount of drafts for the initial search query needs to be counted,
     * and the limit of the searchQuery needs to be increased by that amount
     *
     * @param {Object} client
     * @param {Object} params
     * @param {function} callback
     * 
     * @returns A big part of the uri string needed to make the API call
     */
    _getUri(client, params, callback) {

        // If no participants or to are specified, we can use filters
        if (!params.participants && !params.to) {
            const filterQuery = internals.createFilterQuery(params);
            return callback(null, filterQuery);
        }
        else {
            // if participants are specified, we have to use search instead of filter (blame microsoft)
            const searchQuery = internals.createSearchQuery(params);
        
            if (params.includeDrafts) {
                return callback(null, searchQuery);
            }
            else {
                this._countDraftMessages(client, searchQuery, (err, resCount) => {

                    if (err) {
                        return callback(err);
                    }
                    if (resCount > 0) {
                        params.limit = params.limit + resCount;
                        return callback(null, internals.createSearchQuery(params));
                    }
                    else {
                        return callback(null, searchQuery);
                    }
                });
            }
        }
    }


    /**
     *
     * @param {*} client 
     * @param {*} params 
     * @param {*} uri 
     */
    _getMessages(client, params, uri, callback) {

        // If no specific folders are needed, a more straightforward flow can be followed
        uri = `me/messages` + uri;

        return this._get(client, uri, (err, resMessages) => {

            if (err) {
                return callback(err);
            }

            if (!params.includeDrafts && (resMessages.value.length > params.original_limit)) {
                resMessages.value = internals.applyLimit(resMessages.value, params.original_limit);
            }

            return callback(null, resMessages);
        });
    }

    /**
     * Get messages for some folders and concat them
     * Not used at the moment
     * 
     * @param {*} client 
     * @param {*} uri 
     * @param {*} requestFolders 
     * @param {*} callback 
     */
    _getFolderMessages(client, uri, requestFolders, callback) {
        
        try {
            client.api('me/mailFolders')
                .select('displayName', 'id')
                .get((err, resFolders) => {

                if (err) {
                    return callback(err);
                }
                
                const folderIds = [];
                for (let i = 0; i < requestFolders.length; ++i) {
                    var foundFolder = resFolders.value.find((folder) => {
                        return folder.displayName.toUpperCase() === requestFolders[i].toUpperCase();
                    });
                    if (foundFolder !== undefined) {
                        folderIds.push(foundFolder.id);
                    }
                }

                if (folderIds.length === 0) {
                    return callback(new Error('No corresponding folders found for your query'));
                }
                if (folderIds.length !== 0) {
                    requestFolders = folderIds;
                }
                let allMessages = [];
                return Async.each(requestFolders, (folderId, callback) => {

                    const folderUri = `me/mailFolders/${folderId}/messages` + uri;
                    client.api(folderUri).get((err, resMessages) => {

                        if (err) {
                            return callback(err);
                        }
                        allMessages = allMessages.concat(resMessages.value);
                        return callback(null, null);
                    });

                },(err, res) => {

                    if (err) {
                        return callback(err);
                    }

                    return callback(null, allMessages);
                });
            }); 
        } catch (err) {
            return callback(err);
        }
    };
    
    /**
     * Execute a GET method on a given uri, with a given client object
     * 
     * @param {Object} client - The microsoft graph client 
     * @param {string} uri - the uri of the call that needs to be made
     *
     */
    _get(client, uri, callback) {

        try {
            client.api(uri)
                .get((err, res) => {
                    if (err) {
                        return callback(err);
                    }

                    return callback(null, res);
                });
        } catch (err) {
            return callback(err);
        }
    };

    /**
     * Counts the amount of draft messages available for a certain search query (to be able to filter them out later)
     * @param {*} client 
     * @param {*} searchQuery 
     * @param {*} callback 
     */
    _countDraftMessages(client, searchQuery, callback) {

        try{
            client.api(`me/mailFolders/Drafts/messages` + searchQuery)
                .count(true)
                .get((err, resCount) => {

                    if (err) {
                        return callback(err);
                    }
                    return callback(null, resCount['@odata.count']);
                });
        } catch (err) {
            return callback(err);
        }
    }

    /**
     * TRANSFORMATION METHODS
     */

    _transformFiles(filesArray) {

        const files = Array.isArray(filesArray) ? filesArray : [filesArray];

        return files.map((file) => {

            const transformedFile = {
                type: file.contentType,
                size: file.size,
                file_name: file.name,
                content_id: file.contentId,
                service_file_id: file.id,
                is_embedded: file.isInline,
                date: new Date(file.lastModifiedDateTime).getTime(),
                service_type: this.name
            };

            if (file['@odata.type'] === "#microsoft.graph.fileAttachment") {
                if (file.isInline) {
                    transformedFile.content_disposition = `inline`
                }
                else{
                    transformedFile.content_disposition = `attachment; filename="${file.name}"`;
                }
            }
            else{
                transformedFile.content_disposition = undefined;
            }

            if (file.messageInfo) {
                transformedFile.service_message_id = file.messageInfo.id;
                transformedFile.service_thread_id = file.messageInfo.conversationId;
                transformedFile.email_message_id = file.messageInfo.internetMessageId;
                transformedFile.addresses = {
                    from: internals.getEmailAddressObjects(file.messageInfo.from),
                    to: internals.getEmailAddressObjects(file.messageInfo.toRecipients),
                    cc: internals.getEmailAddressObjects(file.messageInfo.ccRecipients),
                    bcc: internals.getEmailAddressObjects(file.messageInfo.bccRecipients)
                };
            }

            return transformedFile;
        });
    }

    /** 
     * @param {array} messagesArray - Array of messages
     * 
     * @returns {array | Object} - Returns array of transformed messages or one message object if only one message was provided as parameter
     */
    _transformMessages(messagesArray, callback) {

        const messages = Array.isArray(messagesArray) ? messagesArray : [messagesArray];

        return Async.map(messages, (message, callback) => {

            let bodyContentMimeType = message.body && message.body.contentType;

            if (bodyContentMimeType) {
                if (bodyContentMimeType === 'html') {
                    bodyContentMimeType = 'text/html';
                }

                if (bodyContentMimeType === 'text') {
                    bodyContentMimeType = 'text/plain';
                }
            }

            const transformedMessage = {
                service_message_id: message.id,
                service_thread_id: message.conversationId,
                email_message_id: message.internetMessageId,
                subject: message.subject,
                body: [{
                    type: bodyContentMimeType,
                    content: message.body && message.body.content
                }],
                addresses: {
                    from: internals.getEmailAddressObjects(message.from),
                    to: internals.getEmailAddressObjects(message.toRecipients),
                    cc: internals.getEmailAddressObjects(message.ccRecipients),
                    bcc: internals.getEmailAddressObjects(message.bccRecipients)
                },
                date: message.receivedDateTime ? new Date(message.receivedDateTime).getTime() : null,
                folders: [],
                attachments: message.hasAttachments,
                service_type: this.name
            };

            if (!message.singleValueExtendedProperties || message.singleValueExtendedProperties === 0) {
                return callback(null, transformedMessage);
            }

            return Utils.parseRawMail(message.singleValueExtendedProperties[0].value, (err, mail) => {

                if (err) {
                    return callback(internals.wrapError(err));
                }

                transformedMessage.headers = mail.headers;

                if (mail.headers['in-reply-to'] && mail.headers['in-reply-to'].length > 0) {
                    transformedMessage.in_reply_to = mail.headers['in-reply-to'];
                }
                else {
                    transformedMessage.in_reply_to = null;
                }

                return callback(null, transformedMessage);
            });
        }, callback);
    }

    /**
     * AUTHENTICATION METHODS
     */

    /**
     * Checks if the access token is still valid and gets a new one if needed.
     * 
     * @param {string} refreshToken
     * @param {string} expirationDateString - The expiration date for the access token related to the refresh token. 
     * 
     * @returns {Object | boolean} - If the access token was expired, a new one is returned. If it was still valid, false is returned.
     */
    _refreshTokenNeeded(authObject, callback) {
        
        if (!(authObject.access_token && authObject.refresh_token && authObject.expiration_date)) {
            throw new Error('Authentication object is missing properties. Refer to the docs for more info.');
        }

        const expirationDate = new Date(authObject.expiration_date);

        if (expirationDate < new Date()) {
            return this.oauth2.accessToken
                .create({ refresh_token: authObject.refresh_token, access_token: authObject.access_token })
                .refresh((err, resAuth) => {

                    if (err) {
                        // Modify error a bit to match expected error format
                        if (err.context && err.context.error) {
                            err.message = err.context.error;
                        }

                        err.statusCode = err.status;

                        return callback(err);
                    }

                    authObject.access_token = resAuth.token.access_token;
                    authObject.expiration_date = resAuth.token.expiration_date;

                    resAuth.token.refresh_token = authObject.refresh_token;

                    this._tokensUpdated(resAuth.token);

                    return callback(null, authObject);
                });
        }

        return callback(null, authObject);
    }


    /**
     * Makes sure authentication information gets updated when the access token has been renewed
     * 
     * @param {Object} newAuthObject - The new authentication information
     */
    _tokensUpdated (newAuthObject) {
    
        const authToUpdate = {
            access_token: newAuthObject.access_token,
            refresh_token: newAuthObject.refresh_token,
            expiry_date: newAuthObject.expires_at.toISOString()
        };

        const a = { ...authToUpdate };

        this.emit('newAccessToken', { ...authToUpdate });
    }
};

/**
 * 
 * INTERNAL FUNCTIONS
 * 
 */

/**
 * Initiates the Microsoft Graph client with an access token.
 * 
 * @param {string} access_token 
 * 
 * @returns {Object} - The API client
 */
internals.getClient = (access_token) => {

    const client = MicrosoftGraph.init({
        defaultVersion: 'v1.0',
        authProvider: (done) => {
            done(null, access_token);
        },
        debugLogging: false
    });

    return client;
};

/**
 *
 * @param errorObject
 * @returns {*}
 */
internals.wrapError = (errorObject) => {

    let error = errorObject;
    if (!(errorObject instanceof Error)) {
        error = new Error(errorObject.message);
    }

    return Boom.boomify(error, { statusCode: errorObject.statusCode || 500 });
};

/**
 *
 * @param recipients
 * @returns {*}
 */
internals.getEmailAddressObjects = (recipients) => {

    if (!recipients) {
        return [];
    }

    if (!Array.isArray(recipients)) {
        return {
            name: recipients.emailAddress && recipients.emailAddress.name,
            email: recipients.emailAddress && recipients.emailAddress.address && recipients.emailAddress.address.toLowerCase()
        }
    }

    if (recipients.length === 0) {
        return recipients;
    }

    return recipients.map((recipient) => {

        return {
            name: recipient.emailAddress && recipient.emailAddress.name,
            email: recipient.emailAddress && recipient.emailAddress.address && recipient.emailAddress.address.toLowerCase()
        }
    });
};

/**
 * 
 * @param {Object} authObject - The authentication object, containing all the needed information to authenticate microsoft API calls 
 * 
 * @returns {boolean} - Returns true or false, depending on the information available
 */
internals.isValidAuthentication = (authObject) => {

    return (authObject && authObject.access_token && authObject.refresh_token && authObject.expiration_date);
};

/**
 * Create a query string based on the given parameters object
 * 
 * @param {Object} params - the parameters for the request to be made
 *
 * @returns {string | string[]} - the resulting query string used for the actual call to the graph API 
 */
internals.createSearchQuery = (params) => {

    let searchQuery = '';

    if (params && !(Object.keys(params).length === 0 && params.constructor === Object)) {
        if (params.hasAttachment) {
            if (searchQuery) {
                searchQuery += ` AND `;
            }
            else {
                searchQuery += `?$search="`;
            }
            searchQuery += `hasAttachments:true`;
        }

        if (params.before) {
            if (searchQuery) {
                searchQuery += ` AND `;
            }
            else {
                searchQuery += `?$search="`
            }
            searchQuery += `received<${params.before.toISOString()}`;
        }

        if (params.after) {
            if (searchQuery) {
                searchQuery += ` AND `;
            }
            else {
                searchQuery += `?$search="`
            }
            searchQuery += `received>${params.after.toISOString()}`;
        }

        if (params.from) {
            if (searchQuery) {
                searchQuery += ` AND `;
            }
            else {
                searchQuery += `?$search="`
            }

            searchQuery += `from:${params.from}`;
        }

        if (params.to) {
            if (searchQuery) {
                searchQuery += ` AND `;
            }
            else {
                searchQuery += `?$search="`
            }

            searchQuery += `to:${params.to}`;
        }

        if (params.subject) {
            if (searchQuery) {
                searchQuery += ` AND `;
            }
            else {
                searchQuery += `?$search="`
            }

            // does not match literally, e.g. params.subject = 'test' would also match email subject 'this is a test'
            searchQuery += `subject:${params.subject}`;
        }

        if (params.participants) {
            if (searchQuery) {
                searchQuery += ` AND `;
            }
            else {
                searchQuery += `?$search="`
            }

            if (params.participants.length) {
                searchQuery += `(`;

                for (let i = 0; i < params.participants.length; ++i) {
                    if (i > 0) {
                        searchQuery += ` OR `;
                    }

                    searchQuery += `from:${params.participants[i]} OR to:${params.participants[i]} OR cc:${params.participants[i]}`;
                }

                searchQuery += `)`;
            }
        }

        // End search="" if needed
        if (searchQuery) {
            searchQuery += `"`;
        }

        if (params.limit) {
            searchQuery += `${searchQuery ? '&' : '?'}$top=${params.limit}`;
        }

        return searchQuery;
    }
};

/**
 * Create a query string based on the given parameters object
 * IMPORTANT: NO participants and NO toRecipients supported!!!!
 * 
 * @param {Object} params - the parameters for the request to be made
 * 
 * @returns {string | string []} - the resulting query string used for the actual call to the graph API
 */
internals.createFilterQuery = (params) => {

    let filterQuery = '';

    if (params && !(Object.keys(params).length === 0 && params.constructor === Object)) {
        if (params.hasAttachment) {
            if (filterQuery) {
                filterQuery += ` and `;
            }
            else {
                filterQuery += `?$filter=`;
            }
            filterQuery +=`hasAttachments eq true`
        }
        else if (params.hasAttachment === false) {
            if (filterQuery) {
                filterQuery += ` and `;
            }
            else {
                filterQuery += `?$filter=`;
            }
            filterQuery +=`hasAttachments eq false`
        }

        if (params.before) {
            if (filterQuery) {
                filterQuery += ` and `;
            }
            else {
                filterQuery += `?$filter=`;
            }
            filterQuery += `ReceivedDateTime lt ${params.before.toISOString()}`;
        }

        if (params.after) {
            if (filterQuery) {
                filterQuery += ` and `;
            }
            else {
                filterQuery += `?$filter=`;
            }
            filterQuery += `ReceivedDateTime gt ${params.after.toISOString()}`;
        }

        if (params.from) {
            if (filterQuery) {
                filterQuery += ` and `;
            }
            else {
                filterQuery += `?$filter=`;
            }
            filterQuery += `(from/emailAddress/address) eq '${params.from}'`;
        }

        if (params.subject) {
            if (filterQuery) {
                filterQuery += ` and `;
            }
            else {
                filterQuery += `?$filter=`;
            }
            filterQuery += `subject eq '${params.subject}'`;
        }

        if (params.isDraft) {
            if (filterQuery) {
                filterQuery += ` and `;
            }
            else {
                filterQuery += `?$filter=`;
            }
            filterQuery += `IsDraft eq true`;
        }
        else {
            if (filterQuery) {
                filterQuery += ` and `;
            }
            else {
                filterQuery += `?$filter=`;
            }
            filterQuery += `IsDraft eq false`;
        }

        if (params.limit) {
            filterQuery += `${filterQuery ? '&' : '?'}$top=${params.limit}`;
        }

        return filterQuery;
    }
};

/**
 * 
 * @param {Object[]} messages 
 */
internals.sortMessagesOnDate = (messages) => {

    return messages.sort((a, b) => {

        return new Date(b.receivedDateTime) - new Date(a.receivedDateTime);
    });
};

/** 
 * Filter out all draft messages 
 * 
 * @param {Object[]} messages 
 */
internals.removeDrafts = (messages) => {

    return messages.filter((message) => {
        
        return !message.isDraft;
    });
};

/**
 *
 * @param messages
 * @param original_limit
 * @returns {*}
 */
internals.applyLimit = (messages, original_limit) => {

    let cleanMessages = messages;

    if (original_limit) {
        cleanMessages = internals.removeDrafts(messages);
        if (cleanMessages.length > original_limit) {
            cleanMessages = cleanMessages.slice(0, original_limit - 1);
        }
    }

    return cleanMessages;
};

module.exports = Office365Connector;
