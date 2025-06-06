'use strict';

const EventEmitter = require('events');

const _ = require('lodash');
const Async = require('async');
const Boom = require('@hapi/boom');
const MicrosoftGraph = require('@microsoft/microsoft-graph-client');
const Oauth2 = require('simple-oauth2');
const Wreck = require('@hapi/wreck');

const Utils = require('../lib/utils');

const internals = {
    paramTypes: {
        search: 'search',
        filter: 'filter'
    }
};

// Needed because of https://docs.microsoft.com/en-us/graph/throttling#outlook-service-limits and fix for https://github.com/Salesflare/Server/issues/6616
const concurrentLimitFiles = 4;

/**
 * TODO:
 * If no mailbox data exits for an outlook account responses might be:
 * => HTTP error: 404 / Error code: MailboxNotEnabledForRESTAPI or MailboxNotSupportedForRESTAPI / Error message: “REST API is not yet supported for this mailbox
 */

/**
 * @typedef {import('./index').MessageResource} MessageResource
 * @typedef {import('./index').MessageListResource} MessageListResource
 * @typedef {import('./index').FileResource} FileResource
 * @typedef {import('./index').FileListResource} FileListResource
 * @typedef {import('./index').MessageRecipient} MessageRecipient
 *
 * @typedef {import('@hapi/boom').Boom} Boom
 */

class Office365Connector extends EventEmitter {

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

    /**
     * @typedef {Object} Auth - Authentication object
     * @property {String} access_token
     * @property {String} refresh_token
     * @property {Date} expiration_date
     * @property {*} [id] - will be passed back when emitting `newAccessToken`
     */

    /* MESSAGES */

    /**
     * Get a single message with a certain id
     *
     * Returns a unified message resource when options.raw is falsy or a raw response from the API when truthy.
     * If the message id couldn't be found by the API it was most likely a draft and null will be returned.
     *
     *
     * @throws
     *
     * @param {Auth} auth
     *
     * @param {Object} params
     * @param {String} params.id - the message id
     * @param {String} [params.rfc2822Format=false] - Return the email in rfc2822 format https://www.ietf.org/rfc/rfc2822.txt
     *
     * @param {Object} options
     * @param {Boolean} [options.raw] - If true, the response will not be transformed to the unified object
     * @param {function(Error, (MessageResource | Object | String)?):void} callback
     *
     * @returns {void}
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

        return this._refreshTokenIfNeeded(auth, (err, token) => {

            if (err) {
                return callback(internals.wrapError(err));
            }

            auth = token;

            if (params.rfc2822Format) {
                const client = internals.getClient(auth.access_token, 'beta');

                return client.api(`/me/messages/${params.id}/$value`)
                    .get()
                    .then((resMessage) => {

                        return callback(null, resMessage);
                    })
                    .catch((err_) => {

                        return callback(internals.wrapError(err_));
                    });
            }

            const client = internals.getClient(auth.access_token);

            return client.api(`/me/messages/${params.id}`)
                .select(['id', 'from', 'toRecipients', 'ccRecipients', 'bccRecipients', 'sentDateTime', 'subject', 'internetMessageId', 'conversationId', 'body', 'hasAttachments', 'SingleValueExtendedProperties'])
                .expand('SingleValueExtendedProperties($filter=id eq \'String 0x7D\')')
                .get()
                .then((resMessage) => {

                    if (options.raw) {
                        return callback(null, resMessage);
                    }

                    return this._transformMessages(resMessage, auth, (err, transformedMessages) => {

                        if (err) {
                            return callback(err);
                        }

                        const transformedMessage = transformedMessages[0];
                        const hasInlineImg = (resMessage.body.content.indexOf('cid:') > 0);

                        if (!transformedMessage.attachments && !hasInlineImg) {
                            return callback(null, transformedMessage);
                        }

                        // Only select the properties we need so we don't fetch the content bytes (for performance reasons)
                        return client.api(`/me/messages/${transformedMessage.service_message_id}/attachments`)
                            .select(['microsoft.graph.fileAttachment/contentId', 'microsoft.graph.fileAttachment/contentType', 'id', 'size', 'name', 'lastModifiedDateTime', 'isInline'])
                            .get()
                            .then((resFiles) => {

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
                            })
                            .catch((err_) => {

                                return callback(internals.wrapError(err_));
                            });
                    });
                })
                .catch((err_) => {

                    return callback(internals.wrapError(err_));
                });
        });
    }

    /**
     * @typedef {Object} ListOptions
     * @property {Boolean} [raw] If true the response will not be transformed to the unified object
     * @param {Boolean} [idsOnly] - If true the response will only contain the ids of the messages
     * @property {String} [wellKnownFolderName] Restrict messages to a certain folder, currently only supports 'sent'
     */

    /**
     * Returns a list of messages
     * Paging when working with certain folders is not supported!
     *
     * https://docs.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=http
     *
     * @param {Auth} auth
     *
     * @param {Object} params
     * @param {Number} params.limit - Maximum amount of messages in response
     * @param {Boolean} params.hasAttachment - If true, only return messages that have attachments
     * @param {Date} params.before - Only return messages before this date
     * @param {Date} params.after - Only return messages after this date
     * @param {String} params.pageToken - Token used to retrieve a certain page in the list
     * @param {String} params.from - Only return messages sent from this address
     * @param {String} params.to - Only return messages sent to this address
     * @param {String[]} params.participants - Array of email addresses: only return messages where at least one of these participants is involved
     * @param {String} params.folder - Only return messages in a specific folder, currently only supports 'sent'!
     * @param {Boolean} params.includeDrafts - Whether to include drafts or not, defaults to false
     * @param {Boolean} params.onboarding - Wether call is made to process an onboarding
     *
     * @param {ListOptions}  options
     *
     * @param {function(Error, MessageListResource | { messages: Array.<String> }):void} callback - Returns an array of unified message resources when options.raw is falsy, or the raw response of the API when truthy
     *
     * @returns {void}
     */
    listMessages(auth, params, options, callback) {

        if (!internals.isValidAuthentication(auth)) {
            throw new Error('Invalid authentication. One of the needed properties is missing. Please refer to the documentation to find the required fields.');
        }

        if (params.folder && params.folder !== 'sent') {
            throw new Error('Invalid configuration. Filtering messages by folder currently only supports the well known folder name "sent".');
        }

        if (typeof options === 'function') {
            callback = options;
            options = {};
        }

        options = options || {};
        params.original_limit = params.limit;

        return this._refreshTokenIfNeeded(auth, (err, token) => {

            if (err) {
                return callback(internals.wrapError(err));
            }

            auth = token;

            const client = internals.getClient(auth.access_token);
            let uri = '';

            // PageToken for Office means `nextLink` so it is a uri
            if (params.pageToken) {
                return this._get(client, params.pageToken, undefined, (err, resMessages) => {

                    if (err) {
                        return callback(internals.wrapError(err));
                    }

                    /**
                     * When we don't want drafts and we removed some we recursively fetch more messages in the chunk
                     * to get to the amount asked
                     */
                    if (!params.includeDrafts) {
                        const noDraftsList = internals.removeDrafts(resMessages.value);

                        // When there is no next page just continue without the drafts to prevent an infinite loop of fetching the same messages
                        if (!resMessages['@odata.nextLink']) {
                            resMessages.value = noDraftsList;
                        }
                        else if (noDraftsList.length < resMessages.value.length) {
                            const recursiveParams = {
                                ...params,
                                pageToken: resMessages['@odata.nextLink']
                            };
                            return this.listMessages(auth, recursiveParams, options, (err, recursiveResponse) => {

                                if (err) {
                                    return callback(err);
                                }

                                if (options.raw) {
                                    return callback(null, [...resMessages, ...recursiveResponse.messages]);
                                }
                                else if (options.idsOnly) {
                                    return callback(null, internals.mapIdsOnlyResponse([...resMessages, ...recursiveResponse.messages]));
                                }

                                return this._transformMessages(noDraftsList, auth, (err, transformedMessages) => {

                                    if (err) {
                                        return callback(err);
                                    }

                                    const messageListResource = {
                                        messages: [...transformedMessages, ...recursiveResponse.messages]
                                    };

                                    if (resMessages['@odata.nextLink']) {
                                        messageListResource.next_page_token = resMessages['@odata.nextLink'];
                                    }

                                    return callback(null, messageListResource);
                                });
                            });
                        }
                    }

                    if (options.raw) {
                        return callback(null, resMessages);
                    }
                    else if (options.idsOnly) {
                        return callback(null, internals.mapIdsOnlyResponse(resMessages));
                    }

                    return this._transformMessages(resMessages.value, auth, (err, transformedMessages) => {

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

            const paramsArray = [];

            // We split participants up in chunks of 20 since the MS Graph API seems to struggle with longer search queries
            if (params.participants && params.participants.length > 0) {
                const participantsChunks = _.chunk(params.participants, 20);

                participantsChunks.forEach((participants) => {

                    paramsArray.push({ ...params, participants });
                });
            }
            else {
                paramsArray.push(params);
            }

            let nextPageToken = null;

            return Async.map(paramsArray, (callParams, callback) => {

                // Get a big part of the uri. a callback is needed, since a call to the graph API might be needed.
                uri = this._getUri(callParams);

                return this._getMessages(client, {
                    ...options,
                    wellKnownFolderName: callParams.folder
                }, uri, (err, resMessages) => {

                    if (err) {
                        return callback(err);
                    }

                    if (!resMessages || !resMessages.value || resMessages.value.length === 0) {
                        return callback(null, []);
                    }

                    if (paramsArray.length === 1 && resMessages['@odata.nextLink']) {
                        nextPageToken = resMessages['@odata.nextLink'];
                    }

                    /**
                     * When we don't want drafts and we removed some we recursively fetch more messages in the chunk
                     * to get to the amount asked
                     */
                    if (!callParams.includeDrafts) {
                        const noDraftsList = internals.removeDrafts(resMessages.value);

                        // When there is no next page just continue without the drafts to prevent an infinite loop of fetching the same messages
                        if (!nextPageToken) {
                            resMessages.value = noDraftsList;
                        }
                        else if (noDraftsList.length < resMessages.value.length) {
                            const recursiveParams = {
                                ...callParams, // Use callParams since we want to go recursive for this chunk only
                                pageToken: nextPageToken
                            };
                            return this.listMessages(auth, recursiveParams, options, (err, recursiveResponse) => {

                                if (err) {
                                    return callback(err);
                                }

                                if (options.raw) {
                                    return callback(null, [...resMessages, ...recursiveResponse.messages]);
                                }
                                else if (options.idsOnly) {
                                    return callback(null, internals.mapIdsOnlyResponse([...resMessages, ...recursiveResponse.messages]).messages);
                                }

                                return this._transformMessages(noDraftsList, auth, (err, transformedMessages) => {

                                    if (err) {
                                        return callback(err);
                                    }

                                    return callback(null, [...transformedMessages, ...recursiveResponse.messages]);
                                });
                            });
                        }
                    }

                    if (options.raw) {
                        return callback(null, resMessages);
                    }
                    else if (options.idsOnly) {
                        return callback(null, internals.mapIdsOnlyResponse(resMessages).messages);
                    }

                    return this._transformMessages(resMessages.value, auth, (err, transformedMessages) => {

                        if (err) {
                            return callback(err);
                        }

                        return callback(null, transformedMessages);
                    });
                });
            }, (err, allMessages) => {

                if (err) {
                    return callback(internals.wrapError(err));
                }

                const messageListResource = {
                    messages: allMessages.flat().filter((x) => !!x)
                };

                if (nextPageToken) {
                    messageListResource.next_page_token = nextPageToken;
                }

                if (paramsArray.length > 1) {
                    messageListResource.messages = _.orderBy(messageListResource.messages, 'date', 'desc').slice(0, params.limit);
                }

                return callback(null, messageListResource);
            });
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
     * @param {MessageRecipient} params.from
     * @param {MessageRecipient[]} params.to
     * @param {MessageRecipient[]} params.cc
     * @param {MessageRecipient[]} params.bcc
     * @param {{ name: String, url: String, contentBytes: any }[]} params.attachments
     *
     * @param {Object} options
     *
     * @param {function(Error, Object):void} callback
     *
     * @returns {void}
     */
    sendMessage(auth, params, options, callback) {

        if (!internals.isValidAuthentication(auth)) {
            throw new Error('Invalid authentication. The authentication information is incorrect or expired.');
        }

        if (typeof options === 'function') {
            callback = options;
            options = {};
        }

        return Async.parallel({
            refreshToken: Async.apply(this._refreshTokenIfNeeded.bind(this), auth),
            convertAttachments: (callback) => {

                if (!params.attachments || params.attachments.length === 0) {
                    return callback();
                }

                return Async.map(params.attachments, (attachment, callback) => {

                    if (attachment.contentBytes) {
                        return callback(null, {
                            ...attachment,
                            '@odata.type': '#microsoft.graph.FileAttachment'
                        });
                    }

                    return Wreck.get(attachment.url).then((result) => {

                        return callback(null, {
                            '@odata.type': '#microsoft.graph.FileAttachment',
                            name: attachment.name,
                            contentLocation: attachment.url,
                            contentBytes: Buffer.from(result.payload).toString('base64')
                        });
                    })
                        .catch((err) => {

                            if (err) {
                                if (err.data && err.data.payload) {
                                    return callback(err.data.payload);
                                }

                                return callback(err);
                            }
                        });
                }, (err, convertedAttachments) => {

                    if (err) {
                        return callback(err);
                    }

                    return callback(null, convertedAttachments);
                });
            }
        }, (err, results) => {

            if (err) {
                return callback(internals.wrapError(err));
            }

            const contentType = params.html ? 'html' : 'text';
            const message = {
                subject: params.subject,
                from: internals.convertUnimailToMSGraphRecipient(params.from),
                toRecipients: params.to.map(internals.convertUnimailToMSGraphRecipient),
                body: {
                    contentType,
                    content: params[contentType]
                }
            };

            let version = 'v1.0';
            const attachments = results.convertAttachments;

            if (attachments && attachments.length > 0) {
                message.hasAttachments = true;
                message.attachments = attachments;

                if (attachments.some((attachment) => attachment['@odata.type'] === '#microsoft.graph.ReferenceAttachment')) {
                    version = 'beta';
                }
            }

            const client = internals.getClient(results.refreshToken.access_token, version);

            if (params.cc && params.cc.length > 0) {
                message.ccRecipients = params.cc.map(internals.convertUnimailToMSGraphRecipient);
            }

            if (params.bcc && params.bcc.length > 0) {
                message.bccRecipients = params.bcc.map(internals.convertUnimailToMSGraphRecipient);
            }

            if (params.inReplyTo) {
                return this._reply(message, params, client, callback);
            }

            return this._sendMessage(message, client, callback);
        });
    }

    /* FILES */

    /**
     * @typedef {Object} FilesResponse
     * @property {Object[]} files
     * @property {String} [next_page_token]
     */

    /**
     * @param {Auth} auth
     * @param {Object} params
     * @param {Object} options
     * @param {Boolean} [options.raw]
     * @param {function(Error, FilesResponse):void} callback
     *
     * @returns {void}
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

        return this._refreshTokenIfNeeded(auth, (err, token) => {

            if (err) {
                return callback(internals.wrapError(err));
            }

            auth = token;

            const client = internals.getClient(auth.access_token);
            let files = [];
            let uri = '';

            /**
             * TODO Page tokens, UPDATE: nvm, not super imported right now.
             */
            uri = this._getUri(params);

            return client.api(`/me/messages${uri}`)
                .select(['id', 'from', 'toRecipients', 'ccRecipients', 'bccRecipients', 'sentDateTime', 'subject', 'internetMessageId', 'conversationId'])
                .get()
                .then((resMessages) => {

                    if (!resMessages || !resMessages.value || resMessages.value.length === 0) {
                        return callback(null, { files: [] });
                    }

                    return Async.eachLimit(resMessages.value, concurrentLimitFiles, (message, callback) => {

                        // Only select the properties we need so we don't fetch the content bytes (for performance reasons)
                        return client.api(`/me/messages/${message.id}/attachments`)
                            .select(['microsoft.graph.fileAttachment/contentId', 'microsoft.graph.fileAttachment/contentType', 'id', 'size', 'name', 'lastModifiedDateTime', 'isInline'])
                            .get()
                            .then((resFiles) => {

                                if (!resFiles) {
                                    return callback();
                                }

                                if (!options.raw) {
                                    resFiles.value.forEach((resFile) => {

                                        resFile.messageInfo = message;
                                    });
                                }

                                files = [...files, ...resFiles.value];

                                return callback();
                            })
                            .catch((err) => {

                                return callback(err);
                            });
                    }, (err) => {

                        if (err) {
                            return callback(internals.wrapError(err));
                        }

                        if (options.raw) {
                            return callback(null, files);
                        }

                        const filesResponse = { files: this._transformFiles(files) };

                        if (resMessages['@odata.nextLink']) {
                            filesResponse.next_page_token = resMessages['@odata.nextLink'];
                        }

                        return callback(null, filesResponse);
                    });
                })
                .catch((err_) => {

                    return callback(internals.wrapError(err_));
                });
        });
    }

    getFolders(auth, callback) {

        return callback(new Error('Method not implemented'));
    }

    /**
     * @throws
     *
     * @param {Auth} auth
     * @param {Object} params
     * @param {Object} options
     * @param {function(Error, FileResource):void} callback
     *
     * @returns {void}
     */
    getFile(auth, params, options, callback) {

        if (!internals.isValidAuthentication(auth)) {
            throw new Error('Invalid authentication. One of the needed properties is missing. Please refer to the documentation to find the required fields.');
        }

        if (!params || !params.id || !params.messageId) {
            throw new Error('Invalid configuration. Please refer to the documentation to get the required fields.');
        }

        if (typeof options === 'function') {
            callback = options;
        }

        return this._refreshTokenIfNeeded(auth, (err, token) => {

            if (err) {
                return callback(internals.wrapError(err));
            }

            auth = token;

            const client = internals.getClient(auth.access_token);

            return client.api(`me/messages/${params.messageId}/attachments/${params.id}`)
                .get()
                .then((resFile) => {

                    return callback(null, {
                        size: resFile.size,
                        file_name: resFile.name,
                        type: resFile.contentType,
                        service_file_id: resFile.id,
                        service_message_id: params.messageId,
                        data: resFile.contentBytes,
                        file_id: resFile.contentId
                    });
                })
                .catch((err_) => {

                    return callback(err_);
                });
        });
    }

    refreshAuthCredentials(auth, callback) {

        this._refreshTokenIfNeeded(auth, (err, resAuth) => {

            if (err) {
                return callback(err);
            }

            auth = resAuth;
            return callback(null, auth);
        });
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
     * @param {Object} params
     *
     * @returns {String} A big part of the uri string needed to make the API call
     */
    _getUri(params) {

        // If no participants or to are specified, we can use filters
        if (!params.participants && !params.to && !params.onboarding) {
            return internals.createFilterQuery(params);
        }
        else {
            // If participants are specified, we have to use search instead of filter (blame microsoft)
            return internals.createSearchQuery(params);
        }
    }

    /**
     * https://docs.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=http
     *
     * @param {MicrosoftGraph.Client} client
     * @param {ListOptions} params
     * @param {String} uri
     * @param {function(Error, { value: Array.<Object> }?):void} callback
     * @returns {void}
     */
    _getMessages(client, params, uri, callback) {

        let messagesUri;

        if (params.wellKnownFolderName === 'sent') {
            messagesUri = `/me/mailFolders('sentItems')/messages${uri}`;
        }
        else {
            messagesUri = `/me/messages${uri}`;
        }

        let select;

        if (params.idsOnly) {
            select = ['id'];
        }

        return this._get(client, messagesUri, select, callback);
    }

    /**
     * Execute a GET method on a given uri, with a given client object
     *
     * @param {MicrosoftGraph.Client} client - The microsoft graph client
     * @param {String} uri - the uri of the call that needs to be made
     * @param {Array.<String>} select - If you only need specific fields you can pass these here
     * @param {function(Error, Object):void} callback
     *
     * @returns {void}
     */
    _get(client, uri, select = [], callback) {

        const getRequest = client.api(uri);

        if (select.length > 0) {
            getRequest.select(select);
        }

        getRequest.get()
            .then((res) => {

                return callback(null, res);
            })
            .catch((err) => {

                return callback(err);
            });
    }

    /**
     * @param {Object} message
     * @param {MicrosoftGraph.Client} client - The microsoft graph client
     * @param {function(Error, Object):void} callback
     * @returns {void}
     */
    _sendMessage(message, client, callback) {

        // `/sendMail` doesn't return an id, or anything for that matter.
        // So we have to create a draft first and then send
        client.api('/me/messages').post(message)
            .then((result) => {

                return client.api(`/me/messages/${result.id}/send`).post({})
                    .then(() => {

                        // If we don't get an id back we want to notify this without failing the call
                        // since the message was still sent
                        // The emit allows the caller to still log non-blocking errors
                        if (!result.internetMessageId) {
                            const err = new Error('No internetMessageId returned on Office365 message creation');
                            err.id = result.id;
                            err.mail_message = message;
                            this.emit('error', err);
                        }

                        return callback(null, result.internetMessageId);
                    })
                    .catch((err) => {

                        return callback(internals.wrapError(err));
                    });
            })
            .catch((err) => {

                return callback(internals.wrapError(err));
            });
    }

    /**
     * @param {Object} message
     * @param {{ inReplyTo: String }} params
     * @param {MicrosoftGraph.Client} client - The microsoft graph client
     * @param {function(Error, Object):void} callback
     * @returns {void}
     */
    _reply(message, params, client, callback) {

        // `/reply` doesn't return an id, or anything for that matter.
        // So we have to create a draft first and set everything we want manually and then send
        client.api(`/me/messages/${params.inReplyTo}/createReply`).post({})
            .then((response) => {

                // Remove the attachments from message to prevent double attachments if PATCH /messages starts supporting attachments
                const attachments = [...(message.attachments || [])];
                delete message.attachments;

                return client.api(`/me/messages/${response.id}`).patch(message)
                    .then((updateResponse) => {

                        const attachmentPromises = attachments.map((attachment) => {

                            return client.api(`/me/messages/${response.id}/attachments`).post(attachment);
                        });

                        return Promise.all(attachmentPromises)
                            .then(() => {

                                return client.api(`/me/messages/${response.id}/send`).post({}).then(() => {

                                    return callback(null, updateResponse.internetMessageId);
                                }).catch((err) => {

                                    return callback(internals.wrapError(err));
                                });
                            })
                            .catch((err) => {

                                return callback(internals.wrapError(err));
                            });
                    })
                    .catch((err) => {

                        return callback(internals.wrapError(err));
                    });
            })
            .catch((err) => {

                // This most likely means we are replying to an email not in our mailbox, if so just send a regular message
                // Hotmail for example seems to return a 404 resource not found while Outlook returns a 400 bad request
                // Office365 mailboxes can also return this vague 'Mailbox move in progress' error when attempting to reply to a message from another inbox.
                if ((err.statusCode >= 400 && err.statusCode <= 404) || (err.statusCode === 503 && err.message?.toLowerCase()?.includes('mailbox move in progress'))) {
                    return this._sendMessage(message, client, callback);
                }

                return callback(internals.wrapError(err));
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

    /**
     * TRANSFORMATION METHODS
     */

    /**
     * @param {Array.<Object> | Object} filesArray
     *
     * @returns {Array.<FileResource>} - Array of unified file resources
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
                date: new Date(file.messageInfo.sentDateTime).getTime(),
                service_type: this.name,
                content_disposition: undefined
            };

            if (file['@odata.type'] === '#microsoft.graph.fileAttachment') {
                if (file.isInline) {
                    transformedFile.content_disposition = 'inline';
                }
                else {
                    transformedFile.content_disposition = `attachment; filename="${file.name}"`;
                }
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
     * @param {Array.<Object> | Object} messagesArray - Array of messages
     * @param {Auth} auth - Auth will be used to add folder data
     * @param {function(Boom, Array.<MessageResource>):void} callback - Returns array of transformed messages or one message object if only one message was provided as parameter
     *
     * @returns {void}
     */
    _transformMessages(messagesArray, auth, callback) {

        const messages = Array.isArray(messagesArray) ? messagesArray : [messagesArray];

        /**
         * The graph api does not return any meaningful folder data (only an id).
         * So we fetch all the folders and match the folders ourselves.
         * Note that the graph api also doesn't support filtering by id 🤷‍♀️
         */
        return internals.getFolders(auth, (err, folders) => {

            if (err) {
                return callback(internals.wrapError(err));
            }

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
                    date: message.sentDateTime ? new Date(message.sentDateTime).getTime() : null,
                    folders: internals.findFolders(folders, message.parentFolderId),
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
        });
    }

    /**
     * AUTHENTICATION METHODS
     */

    /**
     * Checks if the access token is still valid and gets a new one if needed.
     *
     * @param {Auth} authObject
     *
     * @param {(function(Error):void) | (function(null, Auth):void)} callback
     *
     * @returns {void} - If the access token was expired, a new one is returned. If it was still valid, false is returned.
     */
    _refreshTokenIfNeeded(authObject, callback) {

        if (!(authObject.access_token && authObject.refresh_token && authObject.expiration_date)) {
            throw new Error('Authentication object is missing properties. Refer to the docs for more info.');
        }

        const token = this.oauth2.accessToken
            .create({
                refresh_token: authObject.refresh_token,
                access_token: authObject.access_token,
                expires_at: authObject.expiration_date
            });

        if (!token.expired()) {
            return callback(null, authObject);
        }

        token.refresh((err, resAuth) => {

            if (err) {
                // Modify error a bit to match expected error format
                if (err.context && err.context.error) {
                    err.message = err.context.error;
                }

                err.statusCode = err.status;

                return callback(err);
            }

            authObject.access_token = resAuth.token.access_token;
            authObject.expiration_date = resAuth.token.expires_at;

            this._tokensUpdated({
                refresh_token: authObject.refresh_token,
                access_token: resAuth.token.access_token,
                expires_at: resAuth.token.expires_at,
                id: authObject.id
            });

            return callback(null, authObject);
        });
    }


    /**
     * Makes sure authentication information gets updated when the access token has been renewed
     *
     * @param {Object} newAuthObject - The new authentication information
     * @param {String} newAuthObject.access_token
     * @param {String} newAuthObject.refresh_token
     * @param {Date} newAuthObject.expires_at
     * @param {any} [newAuthObject.id]
     *
     * @returns {void}
     */
    _tokensUpdated(newAuthObject) {

        const authToUpdate = {
            access_token: newAuthObject.access_token,
            refresh_token: newAuthObject.refresh_token,
            expiry_date: newAuthObject.expires_at.toISOString(),
            id: newAuthObject.id
        };

        this.emit('newAccessToken', authToUpdate);
    }
}

/**
 *
 * INTERNAL FUNCTIONS
 *
 */

/**
 * Initiates the Microsoft Graph client with an access token.
 *
 * @param {String} accessToken
 * @param {String} [version='v1.0']
 *
 * @returns {MicrosoftGraph.Client} - The API client
 */
internals.getClient = (accessToken, version = 'v1.0') => {

    const client = MicrosoftGraph.Client.initWithMiddleware({
        defaultVersion: version,
        authProvider: {
            getAccessToken: () => accessToken
        },
        debugLogging: false
    });

    return client;
};

/**
 *
 * @param {Error} errorObject
 * @returns {Boom}
 */
internals.wrapError = (errorObject) => {

    let error = errorObject;

    if (!(errorObject instanceof Error)) {
        error = new Error(errorObject.message);
        Object.entries(errorObject).forEach((keyVal) => {

            error[keyVal[0]] = keyVal[1];
        });
    }

    let statusCode = Number.parseInt(errorObject.statusCode);

    if (Number.isNaN(statusCode) || statusCode < 400) {
        statusCode = 500;
    }

    return Boom.boomify(error, { statusCode });
};

/**
 *
 * @typedef {Object} EmailAddress
 * @property {String} [name] - optional name of recipient
 * @property {String} address - email address of recipient
 */

/**
 * @param {Array<{ emailAddress: EmailAddress }> | { emailAddress: EmailAddress }} [recipients]
 *
 * @returns {Array | Array.<MessageRecipient> | MessageRecipient}
 */
internals.getEmailAddressObjects = (recipients) => {

    if (!recipients) {
        return [];
    }

    if (!Array.isArray(recipients)) {
        return {
            name: recipients.emailAddress && recipients.emailAddress.name,
            email: recipients.emailAddress && recipients.emailAddress.address && recipients.emailAddress.address.toLowerCase()
        };
    }

    if (recipients.length === 0) {
        return recipients;
    }

    return recipients.map((recipient) => {

        return {
            name: recipient.emailAddress && recipient.emailAddress.name,
            email: recipient.emailAddress && recipient.emailAddress.address && recipient.emailAddress.address.toLowerCase()
        };
    });
};

/**
 *
 * @param {Auth} authObject - The authentication object, containing all the needed information to authenticate microsoft API calls
 *
 * @returns {Boolean} - Returns true or false, depending on the information available
 */
internals.isValidAuthentication = (authObject) => {

    return (authObject && authObject.access_token && authObject.refresh_token && !!authObject.expiration_date);
};

/**
 * Create a query string based on the given parameters object
 *
 * @param {Object} params - the parameters for the request to be made
 *
 * @returns {String} - the resulting query string used for the actual call to the graph API
 */
internals.createSearchQuery = (params) => {

    let searchQuery = '';

    if (params && !(Object.keys(params).length === 0 && params.constructor === Object)) {
        if (params.hasAttachment) {
            if (searchQuery) {
                searchQuery += ' AND ';
            }
            else {
                searchQuery += '?$search="';
            }

            searchQuery += 'hasAttachments:true';
        }

        if (params.before) {
            if (searchQuery) {
                searchQuery += ' AND ';
            }
            else {
                searchQuery += '?$search="';
            }

            searchQuery += `sent<${params.before.toISOString()}`;
        }

        if (params.after) {
            if (searchQuery) {
                searchQuery += ' AND ';
            }
            else {
                searchQuery += '?$search="';
            }

            searchQuery += `sent>${params.after.toISOString()}`;
        }

        if (params.from) {
            if (searchQuery) {
                searchQuery += ' AND ';
            }
            else {
                searchQuery += '?$search="';
            }

            searchQuery += `from:${internals.encodeParam(params.from, internals.paramTypes.search)}`;
        }

        if (params.to) {
            if (searchQuery) {
                searchQuery += ' AND ';
            }
            else {
                searchQuery += '?$search="';
            }

            searchQuery += `to:${internals.encodeParam(params.to, internals.paramTypes.search)}`;
        }

        if (params.subject) {
            if (searchQuery) {
                searchQuery += ' AND ';
            }
            else {
                searchQuery += '?$search="';
            }

            /*
             * Does not match literally, e.g. params.subject = 'test' would also match email subject 'this is a test'
             * Remove & symbol since MS Graph API can't deal with it
             */
            searchQuery += `subject:${params.subject.replace(/&/g, '')}`;
        }

        if (params.participants) {

            // Filter out empty participant values, there was an unhandled exception where internals.encodeParams, could not handle 'null' values
            params.participants = params.participants.filter((participant) => participant !== null && participant !== undefined);

            if (searchQuery) {
                searchQuery += ' AND ';
            }
            else {
                searchQuery += '?$search="';
            }

            if (params.participants.length > 0) {
                searchQuery += '(';

                for (let i = 0; i < params.participants.length; ++i) {
                    if (i > 0) {
                        searchQuery += ' OR ';
                    }

                    const encodedParticipant = internals.encodeParam(params.participants[i], internals.paramTypes.search);
                    searchQuery += `from:${encodedParticipant} OR to:${encodedParticipant} OR cc:${encodedParticipant}`;
                }

                searchQuery += ')';
            }
        }

        // End search="" if needed
        if (searchQuery) {
            searchQuery += '"';
        }

        if (params.limit >= 0) {
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
 * @returns {String} - the resulting query string used for the actual call to the graph API
 */
internals.createFilterQuery = (params) => {

    let filterQuery = '';

    if (params && !(Object.keys(params).length === 0 && params.constructor === Object)) {
        // We need to make sure `sentDateTime` is the first thing in the filter query
        // This is a requisite from the Graph API to make sure we can order by `sentDateTime`
        // https://docs.microsoft.com/en-us/graph/query-parameters#orderby-parameter
        if (!params.before && !params.after) {
            filterQuery += '?$filter=SentDateTime gt 1970-01-01T00:00:00Z';
        }

        if (params.before) {

            if (filterQuery) {
                filterQuery += ' and ';
            }
            else {
                filterQuery += '?$filter=';
            }

            filterQuery += `SentDateTime lt ${params.before.toISOString()}`;
        }

        if (params.after) {
            if (filterQuery) {
                filterQuery += ' and ';
            }
            else {
                filterQuery += '?$filter=';
            }

            filterQuery += `SentDateTime gt ${params.after.toISOString()}`;
        }

        if (params.hasAttachment) {
            if (filterQuery) {
                filterQuery += ' and ';
            }
            else {
                filterQuery += '?$filter=';
            }

            filterQuery += 'hasAttachments eq true';
        }
        else if (params.hasAttachment === false) {
            if (filterQuery) {
                filterQuery += ' and ';
            }
            else {
                filterQuery += '?$filter=';
            }

            filterQuery += 'hasAttachments eq false';
        }

        if (params.from) {
            if (filterQuery) {
                filterQuery += ' and ';
            }
            else {
                filterQuery += '?$filter=';
            }

            const encodedParam = internals.encodeParam(params.from);
            // We use `startswith` here, because GRAPH API results for 'eq' are inconsistent
            filterQuery += `(startswith(from/emailAddress/address,'${encodedParam}') or from/emailAddress/address eq '${encodedParam}')`;
        }

        if (params.subject) {
            if (filterQuery) {
                filterQuery += ' and ';
            }
            else {
                filterQuery += '?$filter=';
            }

            const encodedParam = internals.encodeParam(params.subject);
            filterQuery += `subject eq '${encodedParam}'`;
        }

        if (params.isDraft) {
            if (filterQuery) {
                filterQuery += ' and ';
            }
            else {
                filterQuery += '?$filter=';
            }

            filterQuery += 'IsDraft eq true';
        }
        else {
            if (filterQuery) {
                filterQuery += ' and ';
            }
            else {
                filterQuery += '?$filter=';
            }

            filterQuery += 'IsDraft eq false';
        }

        if (params.limit) {
            filterQuery += `${filterQuery ? '&' : '?'}$top=${params.limit}`;
        }

        // The MS Graph API default sorting order is by `sentDateTime` ascending, which we want to reverse
        filterQuery += `${filterQuery ? '&' : '?'}$orderby=SentDateTime desc`;

        return filterQuery;
    }
};

/**
 * Filter query is wrapped in single quotes and when the value contains a quote a badData is returned. Escaped by doubling the single quote.
 * Search query is wrapped in double quotes and when the value contains a double quote a badData is returned or results may not be correct. Escaped by inserting \ before the double quote
 * Wrap the param in encodeURIComponent to avoid characters like + and & breaking the filter on Microsoft's end.
 *
 * @param {String} param
 * @param {"filter" | "query"} paramType
 * @returns {String}
 */
internals.encodeParam = (param, paramType) => {

    let escapedParam = '';
    if (paramType === internals.paramTypes.search) {
        escapedParam = param.replace(/"/g, '\\"');
    }
    else {
        escapedParam = param.replace(/'/g, '\'\'');
    }

    return encodeURIComponent(escapedParam);
};

/**
 *
 * @param {Object[]} messages
 * @returns {Object[]} sorted messages
 */
internals.sortMessagesOnDate = (messages) => {

    return messages.sort((a, b) => {

        return new Date(b.sentDateTime) - new Date(a.sentDateTime);
    });
};

/**
 * Filter out all draft messages
 *
 * @param {{ isDraft: Boolean }[]} messages
 * @returns {{ isDraft: Boolean }[]}
 */
internals.removeDrafts = (messages) => {

    return messages.filter((message) => {

        return !message.isDraft;
    });
};

/**
 *
 * @param {{ isDraft: Boolean }[]} messages
 * @param {Number} originalLimit
 * @returns {{ isDraft: Boolean }[]} messages
 */
internals.applyLimit = (messages, originalLimit) => {

    let cleanMessages = messages;

    if (originalLimit) {
        cleanMessages = internals.removeDrafts(messages);
        if (cleanMessages.length > originalLimit) {
            cleanMessages = cleanMessages.slice(0, originalLimit - 1);
        }
    }

    return cleanMessages;
};

/**
 *
 * @param {{ email: String }} recipient
 * @returns {{ emailAddress: { address: String } }}
 */
internals.convertUnimailToMSGraphRecipient = (recipient) => {

    return {
        emailAddress: {
            address: recipient.email
        }
    };
};

/**
 * @typedef {Map.<String, { name: String, parentFolderId: String? }>} FoldersMap
 */

/**
 * @param {Auth} auth
 * @param {function(Error, FoldersMap):void} callback
 * @returns {void}
 */
internals.getFolders = (auth, callback) => {

    /**
     * We need the beta for 2 reasons:
     * - 1.0 doesn't return folders in folders in folders (sub-sub folders) in any way
     * - 1.0 doesn't return `wellKnownName` (https://docs.microsoft.com/en-us/graph/api/resources/mailfolder?view=graph-rest-beta)
     *
     * Note that beta also returns sub folders on the top level but since they are dupes we can just dedupe those
     */
    const client = internals.getClient(auth.access_token, 'beta');

    client.api('/me/mailFolders')
        .select(['id', 'displayName', 'wellKnownName', 'childFolders'])
        .expand('childFolders')
        .top(250) // We want all the folders at once
        .get()
        .then((resMessage) => {

            const rawFolders = resMessage.value;

            // As far as I could tell you only get folders back 1 level deep
            // sub-sub folders end up under the sub folder in the top level
            const flattenedFolders = [];
            rawFolders.forEach((folder) => {

                flattenedFolders.push(folder);
                if (folder.childFolders) {
                    flattenedFolders.push(...folder.childFolders);
                }
            });

            const folders = new Map();
            flattenedFolders.forEach((folder) => {

                // From what I could test the ones with the parent id come first so we want to keep that data
                // We want the parent id to later on be able to return all the folders of a message
                if (!folders.has(folder.id)) {
                    folders.set(folder.id, {
                        name: folder.wellKnownName || folder.displayName,
                        parentFolderId: folder.parentFolderId // We also want parent so we can match the whole tree
                    });
                }
            });

            return callback(null, folders);
        })
        .catch((err) => {

            return callback(err);
        });
};

/**
 *
 * @param {FoldersMap} folders
 * @param {String} folderId
 * @returns {Array.<String>}
 */
internals.findFolders = (folders, folderId) => {

    const result = [];
    let parentId = folderId;

    // As long as we find results with parentFolderId's we keep resolving those
    do {
        const foundFolder = folders.get(parentId);

        if (!foundFolder) {
            break;
        }

        result.push(foundFolder.name);
        parentId = foundFolder.parentFolderId;
    } while (parentId);

    return result;
};

/**
 * @param {{ value: Array.<{ id: String }> }} response
 * @returns {{ messages: Array.<String> }}
 */
internals.mapIdsOnlyResponse = (response) => {

    return { messages: response.value.map((value) => value.id) };
};

module.exports = Office365Connector;
