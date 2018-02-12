'use strict';

/**
 * Message resource, represents a message and its metadata
 * This object structure should be returned by all custom connectors for messages
 *
 * @global
 * @typedef {Object} MessageResource
 * @property {string} service_message_id - Service specific message id. This id should always be usable with the getMessage method of the specific connector
 * @property {string} service_thread_id - Service specific thread id
 * @property {string} email_message_id - General email message id of the message that contains the file
 * @property {string} subject - General email message id of the message that contains the file
 * @property {Object} headers - Object representing the headers of the message. Keys are the distinct header names and the value is an array of the values of the header
 * @property {Object[]} body - Array of body parts, most emails have both a plain (text/plain) and rich text (text/html) body part
 * @property {string} body.mimeType - MIME type of the body part
 * @property {string} body.content - Actual content of the body part
 * @property {string} in_reply_to - Message id of the message this message is a reply to or null
 * @property {Object} addresses - Object representing the from, to, cc and bcc email addresses of the message
 * @property {Number} date - Unix timestamp of the sending date of the message
 * @property {string} service_type - Service name, same as the connector name
 * @property {string[]} folders - Array of folders names that contain the message
 * @property {boolean} attachments - Whether a file has attachments
 */

/**
 *
 * @global
 * @typedef {Object} MessageListResource
 * @property {MessageResource[]} messages - List of message resources
 * @property {string} next_page_token - Token for the next page of messages
 */

/**
 * File resource, represents a file and its metadata
 * This object structure should be returned by all custom connectors for files
 *
 * @global
 * @typedef {Object} FileResource
 * @property {string} type - MIME type of the file
 * @property {Number} size - Size in kb of the file
 * @property {string} file_name - Name of the file
 * @property {string} content_id - Content id of the file
 * @property {string} content_disposition - Content disposition of the file
 * @property {string} service_file_id - id of the file, specific to each email provider. This id should always be usable with the getFile method of the specific connector
 * @property {Boolean} is_embedded - true for inline files (embedded in the message body), false for attachments
 * @property {string} service_message_id - Service specific message id
 * @property {string} service_thread_id - Service specific thread id
 * @property {string} email_message_id - General email message id of the message that contains the file
 * @property {string} subject - General email message id of the message that contains the file
 * @property {string} service_type - Service name, same as the connector name
 * @property {Object} addresses - Object representing the from, to, cc and bcc email addresses of the message that contains the file
 * @property {Number} date - Unix timestamp of the sending date of the message
 */

/**
 *
 * @global
 * @typedef {Object} FileListResource
 * @property {FileResource[]} files - List of file resources
 * @property {string} next_page_token - Token for the next page of files
 */

const connectors = Symbol('connectors');

class Unimail {

    constructor() {
        this[connectors] = new Map();

        this.messages = {
            list: (connectorName, auth, params, options, callback) => {

                return this.callMethod(connectorName, 'listMessages', auth, params, options, callback);
            },
            get: (connectorName, auth, params, options, callback) => {

                return this.callMethod(connectorName, 'getMessage', auth, params, options,callback);
            },
            transform: (connectorName, data) => {

                return this.callMethod(connectorName, 'transformMessages', data)
            }
        };

        this.files = {
            list: (connectorName, auth, params, options, callback) => {

                return this.callMethod(connectorName, 'listFiles', auth, params, options, callback);
            },
            get: (connectorName, auth, params, options, callback) => {

                return this.callMethod(connectorName, 'getFile', auth, params, options, callback);
            },
            transform: (connectorName, data) => {

                return this.callMethod(connectorName, 'transformFiles', data)
            }
        };

        this.auth = {
            refreshCredentialsIfExpired: (connectorName, auth, callback) => {

                return this.callMethod(connectorName, 'refreshAuthCredentials', auth, callback);
            }
        }
    }

    use(connector) {

        if (!connector) {
            throw new Error('Connector cannot be undefined');
        }

        if (!connector.name) {
            throw new Error('Connector must have a name');
        }

        this[connectors].set(connector.name.toLowerCase(), connector);
    }

    listConnectors() {

        return Array.from(this[connectors].keys());
    }

    callMethod(connectorName, methodName, auth, params, options, callback) {

        if (!connectorName) {
            throw new Error('You should specify a connector name!');
        }

        const name = connectorName.toLowerCase();

        if (!this[connectors].has(name)) {
            throw new Error(`Unknown connector: ${connectorName}`)
        }

        const connector = this[connectors].get(name);

        if (!(methodName in connector)) {
            throw new Error(`This connector does not implement ${methodName}()`);
        }

        return connector[methodName](auth, params, options, callback);
    }
}

Unimail.GmailConnector = require('./unimail-gmail.js');
Unimail.Office365Connector = require('./unimail-office365.js');
Unimail.NylasConnector = require('./unimail-nylas.js');

module.exports = Unimail;
