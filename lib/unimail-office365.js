'use strict';

const EventEmitter = require('events');
const MicrosoftGraph = require("@microsoft/microsoft-graph-client").Client;
const Oauth2 = require('simple-oauth2');
const Async = require('async');
const Boom = require('boom');

const internals = {};

/**
 * TODO:
 * addition header: x-AnchorMailbox:john@contoso.com
 * create test account here: https://signup.live.com/signup?wa=wsignin1.0&rpsnv=12&ct=1454618383&rver=6.4.6456.0&wp=MBI_SSL_SHARED&wreply=https://mail.live.com/default.aspx&id=64855&cbcxt=mai&bk=1454618383&uiflavor=web&uaid=b213a65b4fdc484382b6622b3ecaa547&mkt=E-US&lc=1033&lic=1
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
     * @returns {MessageResource | Object} Returns a unified message resource when options.raw is falsy or a raw response from the API when truthy 
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

        const client = internals.getClient(auth.access_token);

        client.api(`/me/messages/${params.id}`)
            .select('id', 'from', 'toRecipients', 'ccRecipients', 'bccRecipients',
                'sentDateTime', 'subject', 'internetMessageId', 'conversationId', 'body', 'hasAttachments', 'SingleValueExtendedProperties')
            .expand(`SingleValueExtendedProperties($filter=id eq 'String 0x7D')`)
            .get((err, resMessage) => {

            if (err) {
                return callback(internals.wrapError(err));
            }

            if (options.raw) {
                return callback(null, resMessage);
            }

            const transformedMessage = this.transformMessages(resMessage)[0];

            if (!transformedMessage.attachments) {
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

                    transformedMessage.files = this.transformFiles(resFiles.value);
                }

                return callback(null, transformedMessage);
            });
        });
    }

    /**
     * Returns a list of messages
     * TODO: paging when working with certain folders
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
     * @param {string[]} params.folder - Only return messages from these folders
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

        const client = internals.getClient(auth.access_token);
        let uri = ``;

        if (params.pageToken) {
            if (params.folder) {
                return callback(new Error(`Requesting messages from certain folders doesn't support paging at this moment.`));
            }

            this._get(client, params.pageToken, (err, resMessages) => {

                if (err) {
                    return callback(internals.wrapError(err));
                }

                if (options.raw) {
                    return callback(null, resMessages);
                }

                const messageListResource = {
                    messages: this.transformMessages(resMessages.value)
                };

                if (resMessages['@odata.nextLink']) {
                    messageListResource.next_page_token = resMessages['@odata.nextLink'];
                }

                return callback(null, messageListResource);
            });
        }
        
        // get a big part of the uri. a callback is needed, since a call the the graph API might be needed. 
        this._getUri(client, params, (err, resUri) => {
            
            if (err) {
                return callback(err);
            }
            uri = resUri;

            this._getMessages(client, params, uri, (err, resMessages) => {

                if (err) {
                    return callback(internals.wrapError(err));
                }

                if (!resMessages || !resMessages.value || resMessages.value.length === 0) {
                    return callback(null, { messages: [] });
                }

                if (options.raw) {
                    return callback(null, resMessages);
                }

                const messageListResource = {
                    messages: this.transformMessages(resMessages.value)
                };

                if (resMessages['@odata.nextLink']) {
                    messageListResource.next_page_token = resMessages['@odata.nextLink'];
                }

                return callback(null, messageListResource);
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

        const client =  internals.getClient(auth.access_token);
        let files = [];
        let uri = ``;

        /**
         * TODO: Page tokens
         */
        this._getUri(client, params, (err, resUri) => {

            if (err) {
                return callback(err);
            }

            uri = resUri;

            return client.api(`/me/messages${uri}`)
                .select('id', 'from', 'toRecipients', 'ccRecipients', 'bccRecipients',
                    'sentDateTime', 'subject', 'internetMessageId', 'conversationId')
                .get((err, resMessages) => {

                    if (err) {
                        return callback(internals.wrapError(err));
                    }

                    if (!resMessages || !resMessages.value || !resMessages.value.length) {
                        return callback(null, { files: [] });
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

                        return callback(null, this.transformFiles(files));
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

            if (!params.includeDrafts) {
                resMessages.value = internals.applyLimit(resMessages.value, params.original_limit);
            }

            return callback(null, resMessages);
        });
    }

    /**
     * Get messages for some folders and concat them
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
                        console.log('_countDraftMessages', err);
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

    transformFiles(filesArray) {

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
                transformedFile.subject = file.messageInfo.subject;
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
    transformMessages(messagesArray) {

        const messages = Array.isArray(messagesArray) ? messagesArray : [messagesArray];

        return messages.map((message) => {

            const transformedMessage = {
                service_message_id: message.id,
                service_thread_id: message.conversationId,
                email_message_id: message.internetMessageId,
                subject: message.subject,
                body: [{
                    mimeType: message.body && message.body.contentType,
                    content: message.body && message.body.content
                }],
                in_reply_to: (message.replyTo && message.replyTo[0]) || null,  //TODO: look in headers
                addresses: {
                    from: internals.getEmailAddressObjects(message.from),
                    to: internals.getEmailAddressObjects(message.toRecipients),
                    cc: internals.getEmailAddressObjects(message.ccRecipients),
                    bcc: internals.getEmailAddressObjects(message.bccRecipients)
                },
                date: message.sentDateTime ? new Date(message.sentDateTime).getTime() : null,
                folders: [],
                attachments: message.hasAttachments,
                service_type: this.name
            };


            //TODO: parse header
            if (message.singleValueExtendedProperties) {
                transformedMessage.headers = message.singleValueExtendedProperties[0].value/*.replace(/\t/g,"")*/;
            }

            return transformedMessage;
        });
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
                .create({refresh_token: authObject.refresh_token})
                .refresh((err, resAuth) => {

                if (err) {
                    return callback(err);
                }
                authObject.access_token = resAuth.token.access_token;
                authObject.refresh_token = resAuth.token.refresh_token;
                authObject.expiration_date = resAuth.token.expiration_date;
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
            expiration_date: newAuthObject.expires_at.toString()
        };

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
        debugLogging: true
    });

    return client;
};

/**
 *
 * @param errorObject
 * @returns {*}
 */
internals.wrapError = (errorObject) => {

    const error = new Error(errorObject.message);

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
        return recipients.emailAddress;
    }

    if (recipients.length === 0) {
        return recipients;
    }

    return recipients.map((recipient) => {

        return recipient.emailAddress;
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
        searchQuery += `?`;

        // include draft is handled in the countDraftMessages method

        if (params.hasAttachment || params.before || params.after || params.participants) {

            if (params.hasAttachment) {
                if (searchQuery !== '?') {
                    searchQuery += ` AND `;
                }
                else {
                    searchQuery += `$search="`
                }
                searchQuery += `hasAttachments:true`;
            }

            if (params.before) {
                if (searchQuery !== '?') {
                    searchQuery += ` AND `;
                }
                else {
                    searchQuery += `$search="`
                }
                searchQuery += `received<${params.before.toISOString()}`;
            }

            if (params.after) {
                if (searchQuery !== '?') {
                    searchQuery += ` AND `;
                }
                else {
                    searchQuery += `$search="`
                }
                searchQuery += `received>${params.after.toISOString()}`;
            }

            if (params.from) {
                if (searchQuery !== '?') {
                    searchQuery += ` AND `;
                }
                else {
                    searchQuery += `$search="`
                }

                searchQuery += `from:${params.from}`;
            }

            if (params.to) {
                if (searchQuery !== '?') {
                    searchQuery += ` AND `;
                }
                else {
                    searchQuery += `$search="`
                }

                searchQuery += `to:${params.to}`;
            }

            if (params.participants) {
                if (searchQuery !== '?') {
                    searchQuery += ` AND `;
                }
                else {
                    searchQuery += `$search="`
                }

                if (params.participants.length) {
                    searchQuery += `(`;

                    for (let i = 0; i < params.participants.length; ++i) {
                        if (i > 0) {
                            searchQuery += ` OR `;
                        }

                        searchQuery += `from:${params.participants[i]} OR to:${params.participants[i]}`;
                    }

                    searchQuery += `)`;
                }
            }

            // End search="" if needed
            if (searchQuery !== `?`) {
                searchQuery += `"`;
            }

            if (params.limit) {
                if (searchQuery !== `?`) {
                    searchQuery += `&`;
                }
                searchQuery += `$top=${params.limit}`;
            }
        }

        return searchQuery;
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
