const EventEmitter = require('events');
const Async = require('async');

const Nylas = require('nylas');
const Boom = require('boom');


const Utils = require('../lib/utils');

const internals = {};

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

		nylas.messages.find(params.id).then(message => {

			if (options.raw) {
				return callback(null, message);
			}

			return message.getRaw().then(raw => {

				return Utils.parseRawMail(raw, (err, parsedMail) => {

					if (err) {
						return callback(Boom.boomify(new Error(err.message), { statusCode: 500 }));
					}

					message.headers = parsedMail.headers;
					message.textBody = parsedMail.textBody;
					message.in_reply_to = parsedMail.headers['in-reply-to'];
					message.email_message_id = parsedMail.messageId;

					if (message.files && message.files.length > 0) {

						return Async.map(message.files, (file, callback) => {

							const fileObject = {};

							return file.metadata((err, metadata) => {

								if (err) {
									return callback(Boom.boomify(new Error(err.message), { statusCode: err.message.indexOf('Couldn\'t find') > -1 ? 404 : 500 }));
								}

								fileObject.metadata = metadata;

								return file.download((err, download) => {

									if (err) {
										return callback(Boom.boomify(new Error(err.message), { statusCode: err.message.indexOf('Couldn\'t find') > -1 ? 404 : 500 }));
									}

									fileObject.download = download;

									fileObject.message = message;

									if (options.raw) {
										return callback(null, fileObject);
									}

									return callback(null, this.transformFiles(fileObject)[0]);
								});
							});

						}, (err, files) => {

							if (err) {
								return callback(Boom.boomify(new Error(err.message), { statusCode: 500 }));
							}

							message.files = files;

							return callback(null, this.transformMessages(message)[0]);


						});
					}

					return callback(null, this.transformMessages(message)[0]);

				});
			});

		}, err => {

			let statusCode = 500;

			if (err.message) {
				if (err.message.indexOf('Couldn\'t find') > -1) {
					statusCode = 404;
				}
				if (err.message.indexOf('Too many concurrent query requests') > -1) {
					statusCode = 429;
				}
			}

			return callback(Boom.boomify(new Error(err.message), { statusCode: statusCode }));
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
	 * @param {string[]} params.participants - Array of email addresses: only return messages with at least one of these participants are involved.
	 * Due to Nylas api limitation the participants filter will only be applied when an 'after' filter is applied and limit and offset will be ignored
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
	//TODO: rewrite ugly files logic
	listMessages(auth, params, options, callback) {

		if (typeof options === 'function') {
			callback = options;
			options = {};
		}
		options = options || {};

		const nylas = Nylas.with(auth.access_token);

		let participant = null;

		if (params.after && params.participants && params.participants.length > 0) {

			if (params.participants.length > 1) {

				return Async.map(params.participants, (participant, callback) => {

					params.participants = [participant];

					return this.listMessages(auth, params, options, callback);

				}, (err, results) => {

					if (err) {
						return callback(err);
					}

					Async.reduce(results, { messages: [], nextPageToken: 0}, (responseObject, result, callback) => {

						responseObject.messages = responseObject.messages.concat(result.messages.filter((message) => {

							return !responseObject.messages.find((m) => {

								return m.service_message_id === message.service_message_id;
							});
						}));

						return callback(null, responseObject);

					}, (err, responseObject) => {

						if (err) {
							return callback(Boom.boomify(new Error(err.message), 500));
						}

						return Async.sortBy(responseObject.messages, (message, callback) => {

							return callback(null, message.date*-1);
						}, (err, result) => {

							if (err) {
								return callback(Boom.boomify(new Error(err.message), 500));
							}

							responseObject.messages = result;

							return callback(null, responseObject);
						})
					})
				})
			}

			participant = params.participants[0];
		}

		const nylasParams = {
			limit: params.limit,
			received_before: params.before,
			received_after: params.after,
			to: params.to,
			from: params.from,
			in: params.folder
		};

		if (participant) {
			nylasParams.any_email = participant;
		}

		if (params.pageToken) {
			nylasParams.offset = parseInt(params.pageToken)
		}

		nylas.messages.list(nylasParams).then(response => {

			const limit = params.limit || response.length;

			const responseObject = {
				nextPageToken: (limit + (nylasParams.offset || 0)).toString()
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

						if (err) {
							return callback(err);
						}

						responseObject.messages = responseObject.messages.concat(res);

						if (options.raw) {
							return callback(null, responseObject);
						}

						return Async.map(responseObject.messages, (message, callback) => {

							return message.getRaw().then(raw => {

								return Utils.parseRawMail(raw, (err, parsedMail) => {

									if (err) {
										return callback(Boom.boomify(new Error(err.message), { statusCode: 500 }));
									}

									message.headers = parsedMail.headers;
									message.textBody = parsedMail.textBody;
									message.in_reply_to = parsedMail.headers['in-reply-to'];
									message.email_message_id = parsedMail.messageId;

									if (message.files && message.files.length > 0) {

										return Async.map(message.files, (file, callback) => {

										const fileObject = {};

										return file.metadata((err, metadata) => {

											if (err) {
												return callback(Boom.boomify(new Error(err.message), { statusCode: err.message.indexOf('Couldn\'t find') > -1 ? 404 : 500 }));
											}

											fileObject.metadata = metadata;

											return file.download((err, download) => {

												if (err) {
													return callback(Boom.boomify(new Error(err.message), { statusCode: err.message.indexOf('Couldn\'t find') > -1 ? 404 : 500 }));
												}

												fileObject.download = download;

													fileObject.message = message;

													if (options.raw) {
														return callback(null, fileObject);
													}

													return callback(null, this.transformFiles(fileObject)[0]);
											});
										});

										}, (err, files) => {

											if (err) {
												return callback(Boom.boomify(new Error(err.message), { statusCode: 500 }));
											}

											message.files = files;

											return callback(null, this.transformMessages(message)[0]);


										});
									}

									return callback(null, this.transformMessages(message)[0]);

								});
							});

						}, (err, results) => {

							if (err) {
								return callback(Boom.boomify(new Error(err.message), { statusCode: 500 }));
							}
							//responseObject.messages = this.transformMessages(responseObject.messages);
							responseObject.messages = results;

							return callback(null, responseObject);
						});
					});
				}
			}

			responseObject.messages = response;

			if (options.raw) {
				return callback(null, responseObject);
			}

			return Async.map(responseObject.messages, (message, callback) => {

				return message.getRaw().then(raw => {

					return Utils.parseRawMail(raw, (err, parsedMail) => {

						if (err) {
							return callback(Boom.boomify(new Error(err.message), { statusCode: 500 }));
						}

						message.headers = parsedMail.headers;
						message.textBody = parsedMail.textBody;
						message.in_reply_to = parsedMail.headers['in-reply-to'];
						message.email_message_id = parsedMail.messageId;

						if (message.files && message.files.length > 0) {

							return Async.map(message.files, (file, callback) => {

								const fileObject = {};

								return file.metadata((err, metadata) => {

									if (err) {
										return callback(Boom.boomify(new Error(err.message), { statusCode: err.message.indexOf('Couldn\'t find') > -1 ? 404 : 500 }));
									}

									fileObject.metadata = metadata;

									return file.download((err, download) => {

										if (err) {
											return callback(Boom.boomify(new Error(err.message), { statusCode: err.message.indexOf('Couldn\'t find') > -1 ? 404 : 500 }));
										}

										fileObject.download = download;

										fileObject.message = message;

										if (options.raw) {
											return callback(null, fileObject);
										}

										return callback(null, this.transformFiles(fileObject)[0]);
									});
								});

							}, (err, files) => {

								if (err) {
									return callback(Boom.boomify(new Error(err.message), { statusCode: 500 }));
								}

								message.files = files;

								return callback(null, this.transformMessages(message)[0]);


							});
						}

						return callback(null, this.transformMessages(message)[0]);

					});
				});

			}, (err, results) => {

				if (err) {
					return callback(Boom.boomify(new Error(err.message), { statusCode: 500 }));
				}
				//responseObject.messages = this.transformMessages(responseObject.messages);
				responseObject.messages = results;

				return callback(null, responseObject);
			});

			return callback(null, responseObject)
		}, err => {

			let statusCode = 500;

			if (err.message) {
				if (err.message.indexOf('Couldn\'t find') > -1) {
					statusCode = 404;
				}
				if (err.message.indexOf('Too many concurrent query requests') > -1) {
					statusCode = 429;
				}
			}

			return callback(Boom.boomify(new Error(err.message), { statusCode: statusCode }));
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

		if (typeof options === 'function') {
			callback = options;
			options = {};
		}
		options = options || {};

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
				return callback(Boom.boomify(new Error(err.message), { statusCode: err.message.indexOf('Couldn\'t find') > -1 ? 404 : 500 }));
			}

			fileObject.metadata = metadata;

			file.download((err, download) => {

				if (err) {
					return callback(Boom.boomify(new Error(err.message), { statusCode: err.message.indexOf('Couldn\'t find') > -1 ? 404 : 500 }));
				}

				fileObject.download = download;

				return this.getMessage(auth, {id: metadata.message_ids[0]}, options, (err, message) => {

					if (err) {
						return callback(Boom.boomify(new Error(err.message), 500));
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

		if (!messagesArray || messagesArray.length === 0) {
			return [];
		}

		const messages = Array.isArray(messagesArray) ? messagesArray : [messagesArray];

		return messages.map((message) => {

			const formattedMessage = {
				email_message_id: message.email_message_id,
				service_message_id: message.id,
				service_thread_id: message.threadId,
				date: Number(message.date),
				subject: message.subject,
				folders: [message.folder.display_name],
				attachments: message.files.length > 0,
				body: [{content: message.body, type: 'text/html'}],
				addresses: internals.getAddressesObject(message),
				in_reply_to: message.in_reply_to,
				service_type: this.name,
				headers: message.headers
			};

			formattedMessage.body.push(message.textBody);
			formattedMessage.files = message.files;

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
				service_message_id: file.message.id || file.message.service_message_id,
				file_name: file.metadata.filename,
				service_file_id: file.metadata.id,
				content_id: file.metadata.contentId,
				service_type: this.name,
				data: file.download.body.toString('base64')
			}
		});
	}

	// Dummy implementation since Nylas access tokens are not short-lived
    refreshAuthCredentials(auth, callback) {

        return callback(null, auth);
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
