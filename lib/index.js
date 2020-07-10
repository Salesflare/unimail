'use strict';

/**
 * Message resource, represents a message and its metadata
 * This object structure should be returned by all custom connectors for messages
 *
 * @global
 * @typedef {Object} MessageResource
 * @property {String} service_message_id - Service specific message id. This id should always be usable with the getMessage method of the specific connector
 * @property {String} service_thread_id - Service specific thread id
 * @property {String} email_message_id - General email message id of the message that contains the file
 * @property {String} subject - General email message id of the message that contains the file
 * @property {Object} headers - Object representing the headers of the message. Keys are the distinct header names and the value is an array of the values of the header
 * @property {Array.<Body>} body - Array of body parts, most emails have both a plain (text/plain) and rich text (text/html) body part
 * @property {String} in_reply_to - Message id of the message this message is a reply to or null
 * @property {Object} addresses - Object representing the from, to, cc and bcc email addresses of the message
 * @property {Number} date - Unix timestamp of the sending date of the message
 * @property {String} service_type - Service name, same as the connector name
 * @property {Array.<String>} folders - Array of folders names that contain the message
 * @property {Boolean} [attachments=false] - Whether a file has attachments
 * @property {Array.<FileResource>} files
 *
 * @typedef {Object} Body
 * @property {String} type - MIME type of the body part
 * @property {String} content - Actual content of the body part
 */

/**
 *
 * @global
 * @typedef {Object} MessageListResource
 * @property {Array.<MessageResource>} messages - List of message resources
 * @property {String} [next_page_token] - Token for the next page of messages
 */

/**
 *
 * @global
 * @typedef {Object} MessageRecipient
 * @property {String} [name] - optional name of recipient
 * @property {String} email - email address of recipient
 */

/**
 * File resource, represents a file and its metadata
 * This object structure should be returned by all custom connectors for files
 *
 * @global
 * @typedef {Object} FileResource
 * @property {String} type - MIME type of the file
 * @property {Number} size - Size in kb of the file
 * @property {String} file_name - Name of the file
 * @property {String} content_id - Content id of the file
 * @property {String} content_disposition - Content disposition of the file
 * @property {String} service_file_id - id of the file, specific to each email provider. This id should always be usable with the getFile method of the specific connector
 * @property {Boolean} is_embedded - true for inline files (embedded in the message body), false for attachments
 * @property {String} service_message_id - Service specific message id
 * @property {String} service_thread_id - Service specific thread id
 * @property {String} email_message_id - General email message id of the message that contains the file
 * @property {String} service_type - Service name, same as the connector name
 * @property {Object} addresses - Object representing the from, to, cc and bcc email addresses of the message that contains the file
 * @property {Number} date - Unix timestamp of the sending date of the message
 * @property {String} [data] - File data as a base64 string
 */

/**
 *
 * @global
 * @typedef {Object} FileListResource
 * @property {Array.<FileResource>} files - List of file resources
 * @property {String} [next_page_token] - Token for the next page of files
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

                return this.callMethod(connectorName, 'getMessage', auth, params, options, callback);
            },
            send: (connectorName, auth, params, options, callback) => {

                return this.callMethod(connectorName, 'sendMessage', auth, params, options, callback);
            }
        };

        this.files = {
            list: (connectorName, auth, params, options, callback) => {

                return this.callMethod(connectorName, 'listFiles', auth, params, options, callback);
            },
            get: (connectorName, auth, params, options, callback) => {

                return this.callMethod(connectorName, 'getFile', auth, params, options, callback);
            }
        };

        this.auth = {
            refreshCredentialsIfExpired: (connectorName, auth, callback) => {

                return this.callMethod(connectorName, 'refreshAuthCredentials', auth, callback);
            }
        };
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

        return [...this[connectors].keys()];
    }

    callMethod(connectorName, methodName, auth, params, options, callback) {

        if (!connectorName) {
            throw new Error('You should specify a connector name!');
        }

        const name = connectorName.toLowerCase();

        if (!this[connectors].has(name)) {
            throw new Error(`Unknown connector: ${connectorName}`);
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
