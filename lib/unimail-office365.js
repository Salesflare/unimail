'use strict';

const EventEmitter = require('events');
const MicrosoftGraph = require("@microsoft/microsoft-graph-client").Client;
const FOLDER_LABELS = ['INBOX', 'SENT ITEMS', 'DRAFTS', 'SPAM', 'DELETED ITEMS', 'IMPORTANT'];
const Oauth2 = require('simple-oauth2');
const Async = require('async');

const internals = {
};

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

        const client =  internals.getClient(auth.access_token);

        const uri = `me/messages/${params.id}`;

        return this._get(client, uri, (err, resMessage) => {

            if (err) {
                return callback(err);
            }

            if (options.raw) {
                return callback(null, resMessage);
            }
            return callback(null, this.transformMessages(resMessage));
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

        const client =  internals.getClient(auth.access_token);
        let uri = ``;

        if (params.pageToken) {
            if (params.folder) {
                return callback(new Error(`Requesting messages from certain folders doesn't support paging at this moment.`));
            }
            this._get(client, params.pageToken, (err, resMessages) => {

                if (err) {
                    return callback(err);
                }

                if (options.raw) {
                    return callback(null, resMessages);
                }
                return callback(null, this.transformMessages(resMessages));

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
                    return callback(err);
                }

                if (options.raw) {
                    return callback(null, resMessages);
                }
                return callback(null, this.transformMessages(resMessages));

            });
        });
    }

    /**
     * TODO: notification url, clientState
     * https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/webhooks
     */
    watchMessages(auth, notificationUrl, callback) {

        if (!internals.isValidAuthentication(auth)) {
            throw new Error('Invalid authentication. The authentication information is incorrect or expired.')
        }

        const client =  internals.getClient(auth.access_token);
        const uri = `subscriptions`;
        const expirationDateTime = new Date();

        // 4230 minutes long, is the maximum expiration date
        expirationDateTime.setMinutes(4230);

        const requestPayload = {
            changeType: "created,updated",
            notificationUrl: "",
            resource: "/me/messages/",
            expirationDateTime: expirationDateTime.toISOString(),
            clientState: "SecretClientState"
        };

        try{
            client.api(uri)
                .header('Content-Type', 'application/json')
                .post(requestPayload, (err, res) => {

                    //gets here
                    if (err) {
                        return callback(err);
                    }
                    return callback(null, res);
                });
        } catch(err) {
            return callback(err);
        }
    }

    /* FILES */

    /**
     * TODO: Implement params/options
     * @param {*} auth 
     * @param {*} params 
     * @param {*} options 
     * @param {*} callback 
     */
    listFiles(auth, params, options, callback) {

        if (!internals.isValidAuthentication(auth)) {
            throw new Error('Invalid authentication. The authentication information is incorrect or expired.');
        }

        if (typeof options === 'function') {
            callback = options;
            options = {};
        }
        options = options || {};

        const client =  internals.getClient(auth.access_token);
        const messageIds = [];
        let files = [];

        let responseCounter = 0;

        try {
            client.api(`/me/messages`)
                .select('id')
                .filter("hasAttachments eq true")
                .get((err, resMessages) => {
                    if (err) {
                        console.log('listFiles', err);
                        return;
                    }

                    resMessages.value.forEach(email => {
                        messageIds.push(email.id);
                    });

                    for(let i = 0; i < messageIds.length; ++i) {
                        client.api(`/me/messages/${messageIds[i]}/attachments`)
                        .get((err, resFiles) => {
                            
                            responseCounter += 1;
                            if (err) {
                                console.log('listFiles', err);
                                return callback(err);
                            }

                            for (let iterator = 0; iterator < resFiles.value.length; ++iterator){
                                resFiles.value[iterator].messageId = messageIds[i];
                            }

                            files = files.concat(resFiles.value);
                            if (responseCounter === messageIds.length) {
                                callback(null, this.transformFiles(files));
                            }
                        });
                    }
                });
        } catch (err) {
            return callback(err);
        }
    }

    /**
     * TODO: implement
     * @param {*} auth 
     */
    getFile(auth, params, callback) {
        
        if (!internals.isValidAuthentication(auth)) {
            throw new Error('Invalid authentication. One of the needed properties is missing. Please refer to the documentation to find the required fields.');
        }

        if (!params || !params.id || !params.messageId) {
            throw new Error('Invalid configuration. Please refer to the documentation to get the required fields.');
        }

        const client =  internals.getClient(auth.access_token);

        const uri = `me/messages/${params.messageId}/attachments/${params.id}`;

        try {
            client.api(uri)
                .get((err, resFile) => {

                    if (err) {
                        return callback(err);
                    }
                    return callback(null, this.transformFiles(resFile));
                });
        }
        catch(err) {
            return callback(err);
        }
    };

    refreshTokenCheck(auth, callback) {

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
     * the graph API doesn't offer a nice way to filter on folders,
     * so we get messages from all the folders we want, concat them, sort them and cut the result to the correct size
     *
     * @param {*} client 
     * @param {*} params 
     * @param {*} uri 
     */
    _getMessages(client, params, uri, callback) {

        if (params.folder) {
            if (!Array.isArray(params.folder)) {
                const folders = [];
                folder.push(params.folder);
                params.folder = folders;
            }
            this._getFolderMessages(client, uri, params.folder,(err, resFolderMessages) => {

                if (err) {
                    return callback(err);
                }
                    
                //sort and apply limit
                let sortedMessages = internals.sortMessagesOnDate(resFolderMessages);
                if (params.limit && params.limit > 0 && params.limit < sortedMessages.length) {
                    sortedMessages = sortedMessages.slice(0, params.limit);
                }

                if (!params.includeDrafts){
                    sortedMessages = internals.applyLimit(sortedMessages, params.original_limit);
                }

                return callback(null, sortedMessages);                
            });
        }
        // If no specific folders are needed, a more straightforward flow can be followed
        else {
            uri = `me/messages` + uri;

            return this._get(client, uri, (err, resMessages) => {

                if (err) {
                    return callback(err);
                }
                
                if (!params.includeDrafts){
                    resMessages = internals.applyLimit(resMessages, params.original_limit);
                }
                return callback(null, resMessages);
            });
        }
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

        try{
            client.api(uri)
                .get((err, res, rawResponse) => {
                    if (err) {
                        console.log('_get', err);
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
            client.api(`me/mailFolders/Drafts/messages/` + searchQuery)
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

    /**
     * TODO: next token???
     */
    transformFiles(files) {

        if (!files.length) {
            return internals.transformFile(files);
        }
        else {
            const fileListResource = {};
            const transformedFiles = [];

            for (let i = 0; i < files.length; ++i) {
                const file = internals.transformFile(files[i]);
                transformedFiles.push(file);
            }
            fileListResource.files = transformedFiles;
            
            // if (messages['@odata.nextLink']){
            //     messageListResource.next_page_token = messages['@odata.nextLink'];
            // }
            return fileListResource;
        }
    }

    /** 
     * @param {array} messages - Array of messages
     * 
     * @returns {array | Object} - Returns array of transformed messages or one message object if only one message was provided as parameter
     */
    transformMessages(messages) {

        if (!messages.value) {
            return internals.transformMessage(messages);
        }
        else {
            const messageListResource = {};
            const transformedMessages = [];

            for (let i = 0; i < messages.value.length; ++i) {
                const message = internals.transformMessage(messages.value[i]);
                transformedMessages.push(message);
            }
            messageListResource.messages = transformedMessages;
            
            if (messages['@odata.nextLink']){
                messageListResource.next_page_token = messages['@odata.nextLink'];
            }
            return messageListResource;
        }
    }
    

    /**
     * TODO: implement
     * @param {*} part 
     */
    _extractNonContainerParts(part) {

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
            return callback(new Error('The auth object wasnt complete.'));
        }
        const expirationDate = new Date(authObject.expiration_date);
        if (expirationDate < new Date()) {
            this.oauth2.accessToken
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
        else{
            return callback(null, authObject);
        }
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
        }

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
internals.getClient = ( access_token ) => {

    const client = MicrosoftGraph.init({
        defaultVersion: 'v1.0',
        authProvider: (done) => {
            done(null, access_token);
        },
        debugLogging: true
    });

    return client;
}

/**
 * 
 * @param {Object} authObject - The authentication object, containing all the needed information to authenticate microsoft API calls 
 * 
 * @returns {boolean} - Returns true or false, depending on the information available
 */
internals.isValidAuthentication = (authObject) => {

    if (!authObject || !authObject.access_token || !authObject.refresh_token || !authObject.expiration_date) {
        return false;
    }
    return true;
}

/**
 * Create a query string based on the given parameters object
 * 
 * @param {Object} params - the parameters for the request to be made
 * @param {string} method - the method to be called on the api (e.g. messages)
 * 
 * @returns {string | string[]} - the resulting query string used for the actual call to the graph API 
 */
internals.createSearchQuery = (params, method) => {

    let searchQuery = ``;

    if (params && !(Object.keys(params).length === 0 && params.constructor === Object)) {
        searchQuery += `?`; 

        // include draft is handled in the countDraftMessages method

        if (params.hasAttachment || params.before || params.after || params.participants) {
            
            if (params.hasAttachment) {
                if (searchQuery !== `?`) {
                    searchQuery += ` AND `;
                }
                else {
                    searchQuery += `$search="`
                }
                searchQuery += `hasAttachments=true`;
            }

            if (params.before) {
                if (searchQuery !== `?`) {
                    searchQuery += ` AND `;
                }
                else {
                    searchQuery += `$search="`
                }
                searchQuery += `received < ${params.before.toISOString()}`;
            }

            if (params.after) {
                if (searchQuery !== `?`) {
                    searchQuery += ` AND `;
                }
                else {
                    searchQuery += `$search="`
                }
                searchQuery += `received > ${params.after.toISOString()}`;
            }

            if (params.from) {
                if (searchQuery !== `?`) {
                    searchQuery += ` AND `;
                }
                else {
                    searchQuery += `$search="`
                }

                searchQuery += `from:${params.from}`;
            }

            if (params.to) {
                if (searchQuery !== `?`) {
                    searchQuery += ` AND `;
                }
                else {
                    searchQuery += `$search="`
                }

                searchQuery += `to:${params.to}`;
            }

            if (params.participants) {
                if (searchQuery !== `?`) {
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
                        searchQuery += `participants:${params.participants[i]}`; 
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
    }

    return searchQuery;
}

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

    return messages.value.filter((message) => {
        
        return !message.isDraft;
    });
};

/**
 * 
 */
internals.applyLimit = (messages, original_limit) => {

    let cleanMessages = messages;
    if (original_limit) {
        cleanMessages = internals.removeDrafts(messages);
        if (cleanMessages.length > original_limit) {
            cleanMessages = cleanMessages.slice(0, params.original_limit -1);
        }
    }
    return cleanMessages;
};

internals.transformMessage = (message) => {

    const transformedMessage = {};

    transformedMessage.office365_message_id = message.id;
    // TODO: extra call needed???
    transformedMessage.office365_thread_id;
    transformedMessage.email_message_id = message.internetMessageId;
    transformedMessage.subject = message.subject;
    // TODO: extra call needed???
    transformedMessage.headers;
    transformedMessage.body = {};
    transformedMessage.body.mimeType = message.body.contentType;
    transformedMessage.body.content = message.body.content;
    transformedMessage.in_reply_to = message.replyTo[0] || null;
    transformedMessage.addresses = {
        from: message.from,
        to: message.toRecipients,
        cc: message.ccRecipients,
        bcc: message.bccRecipients 
    };
    transformedMessage.date = message.sentDateTime;
    // TODO: extra call needed???
    transformedMessage.folders;
    // TODO: extra call needed???
    transformedMessage.files;

    return transformedMessage;
};

internals.transformFile = (file) => {

    const transformedFile = {};

    transformedFile.type = file.contentType;
    transformedFile.size = file.size;
    transformedFile.file_name = file.name;
    transformedFile.content_id = file.contentId;
    transformedFile.content_disposition = '';
    transformedFile.file_id = file.clientId;
    transformedFile.is_embedded = file.isInline;
    transformedFile.office365_message_id = '';
    transformedFile.office365_thread_id = '';
    transformedFile.email_message_id = '';
    transformedFile.subject = '';
    transformedFile.addresses = {};
    transformedFile.date = '';

    return transformedFile;
};
module.exports = Office365Connector
