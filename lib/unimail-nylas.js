const EventEmitter = require('events');

const Nylas = require('nylas');

const internals = {};

internals.getMessages = (accessToken, callback)=> {

	const nylas = Nylas.with(accessToken);

	nylas.messages.list({}, (err, messages) => {
		console.log(messages.length);
		return callback();
	});

};

class NylasConnector extends EventEmitter {

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

		this.name = 'nylas';

		Nylas.config({
			appId: this.clientId,
			appSecret: this.clientSecret
		});
	}

	/* MESSAGES */

	/**
	 *
	 * @param {Object} auth - Authentication object
	 * @param {string} auth.access_token - Access token
	 *
	 * @param {Object} params
	 * @param {string} params.id - Nylas message id
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

		const nylas = Nylas.with(auth.access_token);

		nylas.messages.find(params.id, (err, response) => {

			if (err) {
				callback(err);
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
	//TODO: implement hasAttachment
	//TODO: implement participants
	listMessages(auth, params, options, callback) {

		if (typeof options === 'function') {
			callback = options;
			options = {};
		}
		options = options || {};

		const nylas = Nylas.with(auth.access_token);

		const nylasParams = {
			limit: params.limit,
			received_before: params.before,
			received_after: params.after,
			to: params.to,
			from: params.from,
			in: params.folder
		};

		if (params.pageToken) {
			nylasParams.offset = parseInt(params.pageToken)
		}

		nylas.messages.list(nylasParams, (err, response) => {

			if (err) {
				callback(err);
			}

			const responseObject = {
				nextPageToken: (params.limit + (nylasParams.offset || 0)).toString()
			};

			if (!params.includeDrafts) {

				responseObject.messages = response.filter((message) => {

					return message.folder.name !== 'drafts';
				});

				if (params.limit && response.length === params.limit && responseObject.messages.length < params.limit) {

					params.limit = params.limit - responseObject.messages.length;

					const filterOptions = options;

					filterOptions.raw = true;

					return this.listMessages(auth, params, filterOptions, (err, res) => {

						responseObject.messages = responseObject.messages.concat(res);

						if (options.raw) {
							return callback(null, responseObject);
						}

						responseObject.messages = this.transformMessages(responseObject.messages);

						return callback(null, responseObject)
					});
				}
			}

			responseObject.messages = response;

			if (options.raw) {
				return callback(null, responseObject);
			}

			responseObject.messages = this.transformMessages(responseObject.messages);

			return callback(null, responseObject)
		});
	}

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
	//TODO: implement decent solution
	listFiles(auth, params, options, callback) {

		return callback(null, []);
	}

	/**
	 *
	 * @param {Object} auth - Authentication object
	 * @param {string} auth.access_token - Access token
	 *
	 * @param {Object} params
	 * @param {string} params.id - Nylas attachement id
	 *
	 * @param {Object} options
	 * @param {boolean} options.raw - If true the response will not be transformed to the unified object
	 *
	 * @returns {FileResource | Object} Returns a unified file resource when options.raw is falsy or the raw response of the API when truthy
	 */
	getFile(auth, params, options, callback) {

		if (!params || !params.id) {
			throw new Error('Invalid configuration. Please refer to the documentation to get the required fields.');
		}

		if (typeof options === 'function') {
			callback = options;
			options = {};
		}
		options = options || {};

		const nylas = Nylas.with(auth.access_token);

		const file = nylas.files.build({id: params.id});

		const fileObject = {};

		file.metadata((err, metadata) => {

			if (err) {
				callback(err);
			}

			fileObject.metadata = metadata;

			file.download((err, download) => {

				if (err) {
					callback(err);
				}

				fileObject.download = download;

				return this.getMessage(auth, {id: metadata.message_ids[0]}, options, (err, message) => {

					if (err) {
						callback(err);
					}

					fileObject.message = message;

					if (options.raw) {
						return callback(null, fileObject);
					}

					return callback(null, this.transformFiles(fileObject)[0]);
				});
			});
		});
	}

	/* TRANSFORMERS */

	/**
	 * Transforms a raw Nylas API messages response to a unified message resource
	 *
	 * @param {Object[]} messagesArray - Array of messages in the format returned by the Nylas API
	 *
	 * @returns {MessageResource[]} - Array of unified message resources
	 */
	transformMessages(messagesArray) {

		const messages = Array.isArray(messagesArray) ? messagesArray : [messagesArray];

		return messages.map((message) => {

			const formattedMessage = {
				//email_message_id: internals.getHeaderValue(message.payload.headers, 'message-id'),
				service_message_id: message.id,
				service_thread_id: message.threadId,
				date: Number(message.date),
				subject: message.subject,
				folders: message.folder.id,
				attachments: message.files.length > 0,
				body: message.body,
				addresses: internals.getAddressesObject(message),
				in_reply_to: message.replyTo,
				service_type: this.name
			};

			//formattedMessage.attachments = internals.getFilesObject(formattedMessage, message.files);

			return formattedMessage;
		});
	}

	/**
	 * Transforms a raw Nylas API response to a unified file resource
	 *
	 * @param {Object[]} filesArray - Array of messages in the format returned by the Nylas API
	 *
	 * @returns {FileResource[]} - Array of unified file resources
	 */
	transformFiles(filesArray) {

		const files = Array.isArray(filesArray) ? filesArray : [filesArray];

		return files.map((file) => {

			return {
				type: file.metadata.content_type,
				size: file.metadata.size,
				service_message_id: file.message.service_message_id,
				service_thread_id: file.message.service_thread_id,
				//email_message_id: internals.getHeaderValue(message.payload.headers, 'message-id'),
				subject: file.message.subject,
				date: Number(file.download.date),
				addresses: file.message.addresses,
				file_name: file.metadata.filename,
				content_id: file.metadata.contentId,
				content_disposition: file.download['content-disposition'],
				file_id: file.metadata.id,
				is_embedded: file.download['content-disposition'] ? file.download['content-disposition'].startsWith('inline;') : false,
				service_type: this.name
			}
		});

	}
}

/* Internal utility functions */

internals.getAddressesObject = (message) => {

	return {
		from: message.from.map((contact) => {

			delete contact.connection;
			return contact;
		}),
		to: message.to.map((contact) => {

			delete contact.connection;
			return contact;
		}),
		cc: message.cc.map((contact) => {

			delete contact.connection;
			return contact;
		}),
		bcc: message.bcc.map((contact) => {

			delete contact.connection;
			return contact;
		})
	};
};

module.exports = NylasConnector;
