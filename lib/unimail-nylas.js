'use strict';

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

		return nylas.messages.find(params.id).then((message) => {

			if (options.raw) {
				return callback(null, message);
			}

			return message.getRaw().then((raw) => {

				if (/^Message .* has not been synced yet$/.test(raw)) {
					return callback(new Error(raw));
				}

				return Utils.parseRawMail(raw, (err, parsedMail) => {

					if (err) {
						return callback(err);
					}

					message.headers = parsedMail.headers;
					message.textBody = parsedMail.textBody;
					message.email_message_id = parsedMail.messageId;

                    if (message.headers['in-reply-to'] && message.headers['in-reply-to'].length > 0) {
						message.in_reply_to = message.headers['in-reply-to'][0];
					}
                    else {
                        message.in_reply_to = null;
                    }

					if (message.files && message.files.length > 0) {
						return Async.map(message.files, (file, callback) => {

							const fileObject = {
								message
							};

							return file.metadata().then((metadata) => {

								fileObject.metadata = metadata;

								return file.download().then((download) => {

									fileObject.download = download;

									if (options.raw) {
										return callback(null, fileObject);
									}

									return callback(null, this._transformFiles(fileObject)[0]);
								})
								.catch((err) => {

									return callback(err);
								});
							})
							.catch((err) => {

								return callback(err);
							});
						}, (err, files) => {

							if (err) {
								return callback(err);
							}

							message.files = files;

							return callback(null, this._transformMessages(message)[0]);
						});
					}

					return callback(null, this._transformMessages(message)[0]);
				});
			})
			.catch((err) => {

				return callback(err);
			});
		})
		.catch((err) => {

			let statusCode = 500;

			if (err.message) {
				if (err.message.indexOf('Couldn\'t find') > -1) {
					statusCode = 404;
				}
				if (err.message.indexOf('Too many concurrent query requests') > -1) {
					statusCode = 429;
				}
			}

			return callback(Boom.boomify(err, { statusCode: statusCode }));
		});
	}

	/**
	 *
	 * @param {Object} auth - Authentication object
	 * @param {string} auth.access_token - Access token
	 *
	 * @param {Object} params
	 * @param {number} params.limit - Maximum amount of messages in response, max = 100
	 * @param {boolean} params.hasAttachment - If true, only return messages with attachments (false doesn't work yet in Nylas)
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
	//TODO: rewrite ugly files logic
	listMessages(auth, params, options, callback) {

		if (typeof options === 'function') {
			callback = options;
			options = {};
		}
		options = options || {};
		const clonedParams = params ? {...params} : {};

		const nylas = Nylas.with(auth.access_token);

		let participant = null;

        if (clonedParams.after && clonedParams.participants && clonedParams.participants.length > 0) {
            if (clonedParams.participants.length > 1) {
                return Async.mapLimit(clonedParams.participants, 5, (participant, callback) => {

                    clonedParams.participants = [participant];

                    return this.listMessages(auth, clonedParams, options, callback);
                }, (err, results) => {

                    if (err) {
                        return callback(err);
                    }

                    const responseObject = results.reduce((accumulator, currentValue) => {

                        return {
                            messages: accumulator.messages.concat(currentValue.messages.filter((message) => {

                                return !accumulator.messages.find((m) => {

                                    return m.service_message_id === message.service_message_id;
                                });
                            })),
                            next_page_token: 0
                        }
                    }, {messages: [], next_page_token: 0});

                    responseObject.messages = responseObject.messages.sort((a, b) => b.date - a.date).slice(0, clonedParams.limit);

                    return callback(null, responseObject);
                })
            }

            participant = clonedParams.participants[0];
        }

		const nylasParams = {
			limit: clonedParams.limit,
			received_before: clonedParams.before,
			received_after: clonedParams.after,
			to: clonedParams.to,
			from: clonedParams.from,
			subject: clonedParams.subject,
			in: clonedParams.folder,
			has_attachment: clonedParams.hasAttachment,
			view: 'expanded' // so we get headers (message id) back
		};

		if (participant) {
			nylasParams.any_email = participant;
		}

		if (clonedParams.pageToken) {
			nylasParams.offset = parseInt(clonedParams.pageToken)
		}

		return nylas.messages.list(nylasParams).then((response) => {

			const limit = clonedParams.limit || response.length;

			const responseObject = {
				next_page_token: (limit + (nylasParams.offset || 0)).toString()
			};

			if (!clonedParams.includeDrafts) {
				responseObject.messages = response.filter((message) => {

					return !message.folder || (message.folder && message.folder.name !== 'drafts');
				});

				if (clonedParams.limit && response.length === clonedParams.limit && responseObject.messages.length < clonedParams.limit) {
					clonedParams.limit = clonedParams.limit - responseObject.messages.length;
					clonedParams.pageToken = clonedParams.pageToken + clonedParams.limit;

					const filterOptions = Object.assign({}, options);

					filterOptions.raw = true;

					return this.listMessages(auth, clonedParams, filterOptions, (err, res) => {

						if (err) {
							return callback(err);
						}

						responseObject.messages = responseObject.messages.concat(res.messages);

						if (options.raw) {
							return callback(null, responseObject);
						}

						return Async.map(responseObject.messages, (message, callback) => {

							message.email_message_id = message.headers['Message-Id'];

							if (message.headers['In-Reply-To'] && message.headers['In-Reply-To'].length > 0) {
								message.in_reply_to = message.headers['In-Reply-To'];
							}
							else {
								message.in_reply_to = null;
							}

							if (message.files && message.files.length > 0) {
								return Async.map(message.files, (file, callback) => {

									const fileObject = {
										message
									};

									return file.metadata().then((metadata) => {

										fileObject.metadata = metadata;

										return file.download().then((download) => {

											fileObject.download = download;

											if (options.raw) {
												return callback(null, fileObject);
											}

											return callback(null, this._transformFiles(fileObject)[0]);
										})
										.catch((err) => {

											return callback(err);
										});
									})
									.catch((err) => {
										
										return callback(err);
									});
								}, (err, files) => {

									if (err) {
										return callback(err);
									}

									message.files = files;

									return callback(null, this._transformMessages(message)[0]);
								});
							}

							return callback(null, this._transformMessages(message)[0]);
						}, (err, results) => {

							if (err) {
								return callback(Boom.boomify(err, { statusCode: 500 }));
							}

							//responseObject.messages = this._transformMessages(responseObject.messages);
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

				message.email_message_id = message.headers['Message-Id'];

				if (message.headers['In-Reply-To'] && message.headers['In-Reply-To'].length > 0) {
					message.in_reply_to = message.headers['In-Reply-To'];
				}
				else {
					message.in_reply_to = null;
				}

				if (message.files && message.files.length > 0) {
					return Async.map(message.files, (file, callback) => {

						const fileObject = {
							message
						};

						return file.metadata().then((metadata) => {

							fileObject.metadata = metadata;

							return file.download().then((download) => {

								fileObject.download = download;

								if (options.raw) {
									return callback(null, fileObject);
								}

								return callback(null, this._transformFiles(fileObject)[0]);
							})
							.catch((err) => {
							
								return callback(err);
							});
						})
						.catch((err) => {
							
							return callback(err);
						});
					}, (err, files) => {

						if (err) {
							return callback(Boom.boomify(err, { statusCode: 500 }));
						}

						message.files = files;

						return callback(null, this._transformMessages(message)[0]);
					});
				}

				return callback(null, this._transformMessages(message)[0]);
			}, (err, results) => {

				if (err) {
					return callback(err);
				}

				//responseObject.messages = this._transformMessages(responseObject.messages);
				responseObject.messages = results;

				return callback(null, responseObject);
			});
		})
		.catch((err) => {

			let statusCode = 500;

			if (err.message) {
				if (err.message.indexOf('Couldn\'t find') > -1) {
					statusCode = 404;
				}

				if (err.message.indexOf('Too many concurrent query requests') > -1) {
					statusCode = 429;
				}
			}

			return callback(Boom.boomify(err, { statusCode: statusCode }));
		});
	}

    sendMessage() {

        throw new Error('Not yet implemented!');
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

		params.hasAttachment = true;

		return this.listMessages(auth, params, options, (err, response) => {

			if (err) {
				return callback(err);
			}

			return Async.reduce(response.messages, [], (files, message, callback) => {

				return callback(null, files.concat(message.files));
			}, (err, files) => {

				if (err) {
					return callback(err);
				}

				const limit = params.limit || files.length;

				return callback(null, { files: files.slice(0, limit), next_page_token: response.messages.length });
			});
		});
	}

	/**
	 *
	 * @param {Object} auth - Authentication object
	 * @param {string} auth.access_token - Access token
	 *
	 * @param {Object} params
	 * @param {string} params.id - Nylas attachment id
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

		return file.metadata().then((metadata) => {

			fileObject.metadata = metadata;

			return file.download().then((download) => {

				fileObject.download = download;

				return this.getMessage(auth, { id: metadata.message_ids[0] }, options, (err, message) => {

					if (err) {
						return callback(Boom.boomify(err, 500));
					}

					fileObject.message = message;

					if (options.raw) {
						return callback(null, fileObject);
					}

					return callback(null, this._transformFiles(fileObject)[0]);
				});
			})
			.catch((err) => {

				return callback(Boom.boomify(err, { statusCode: err.message.indexOf('Couldn\'t find') > -1 ? 404 : 500 }));
			});
		})
		.catch((err) => {
									
			return callback(Boom.boomify(err, { statusCode: err.message.indexOf('Couldn\'t find') > -1 ? 404 : 500 }));
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
	_transformMessages(messagesArray) {

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
				folders: [],
				attachments: message.files.length > 0,
				body: [{content: message.body, type: 'text/html'}],
				addresses: internals.getAddressesObject(message),
				in_reply_to: message.in_reply_to,
				service_type: this.name,
				headers: message.headers
			};

			if (message.folder) {
				formattedMessage.folders = [message.folder.display_name];
			}

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
	_transformFiles(filesArray) {

		const files = Array.isArray(filesArray) ? filesArray : [filesArray];

		return files.map((file) => {

			const formattedFile = {
				type: file.metadata.content_type,
				size: file.metadata.size,
				service_message_id: file.message.id || file.message.service_message_id,
				service_thread_id: file.message.threadId || file.message.service_thread_id,
				email_message_id: file.message.email_message_id,
				file_name: file.metadata.filename,
				service_file_id: file.metadata.id,
				content_disposition: file.download['content-disposition'],
				content_id: file.metadata['content_id'],
				service_type: this.name,
				is_embedded: file.download['content-disposition'] ? file.download['content-disposition'].startsWith('inline;') : false,
				addresses: internals.getAddressesObject(file.message),
				date: file.message.date,
			};

			if (file.download.body) {
				formattedFile.data = file.download.body.toString('base64');
			}

			if (file.metadata.content_id) {
				formattedFile.content_id= file.metadata.content_id.replace(/[<>]/g, "");
			}

			if (file.message.body.indexOf('cid:' + formattedFile.content_id) > -1) {
				formattedFile.is_embedded = true;
				formattedFile.content_disposition = formattedFile.content_disposition.replace('attachment;', 'inline;');
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

module.exports = NylasConnector;
