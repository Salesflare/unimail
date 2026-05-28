'use strict';

const Path = require('path');

require('dotenv').config({ path: Path.join(__dirname, '.env') });

const Hapi = require('@hapi/hapi');
const Inert = require('@hapi/inert');
const Boom = require('@hapi/boom');

const { unimail, connectors } = require('./unimail');

const Folders = require('./routes/folders');
const Messages = require('./routes/messages');

const internals = {};

internals.parseAuthHeader = (value) => {

    if (!value) {
        throw Boom.badRequest('Missing X-Unimail-Auth header');
    }

    let decoded;
    try {
        decoded = JSON.parse(Buffer.from(value, 'base64').toString('utf8'));
    }
    catch {
        throw Boom.badRequest('Invalid X-Unimail-Auth header (expected base64-encoded JSON)');
    }

    if (decoded.expiration_date) {
        decoded.expiration_date = new Date(decoded.expiration_date);
    }

    return decoded;
};

internals.encodeAuthHeader = (auth) => {

    const safe = { ...auth };
    if (safe.expiration_date instanceof Date) {
        safe.expiration_date = safe.expiration_date.toISOString();
    }
    else if (safe.expires_at instanceof Date) {
        safe.expiration_date = safe.expires_at.toISOString();
        delete safe.expires_at;
    }

    return Buffer.from(JSON.stringify(safe), 'utf8').toString('base64');
};

/**
 * Wraps a callback-style unimail call and captures any newAccessToken event
 * that fires for the same auth id during the call.
 *
 * @param request
 * @param namespace
 * @param method
 * @param {...any} args
 */
internals.callUnimail = (request, namespace, method, ...args) => {

    const connectorName = request.app.connector;
    const connector = connectors[connectorName];
    const authId = request.app.auth && request.app.auth.id;

    return new Promise((resolve, reject) => {

        let authUpdated = null;
        const listener = (newAuth) => {

            if (!authId || newAuth.id === authId) {
                authUpdated = newAuth;
            }
        };

        if (connector) {
            connector.on('newAccessToken', listener);
        }

        const done = (err, data) => {

            if (connector) {
                connector.removeListener('newAccessToken', listener);
            }

            if (err) {
                return reject(err);
            }

            return resolve({ data, authUpdated });
        };

        try {
            unimail[namespace][method](connectorName, ...args, done);
        }
        catch (err) {
            if (connector) {
                connector.removeListener('newAccessToken', listener);
            }

            reject(err);
        }
    });
};

const init = async () => {

    const server = Hapi.server({
        port: Number(process.env.PORT) || 3000,
        host: '127.0.0.1',
        routes: {
            cors: false
        }
    });

    await server.register(Inert);

    server.ext('onPreHandler', (request, h) => {

        if (!request.path.startsWith('/api/')) {
            return h.continue;
        }

        const connectorHeader = request.headers['x-unimail-connector'];
        if (!connectorHeader) {
            throw Boom.badRequest('Missing X-Unimail-Connector header');
        }

        if (!connectors[connectorHeader]) {
            throw Boom.badRequest(`Connector "${connectorHeader}" is not configured. Set the required env vars in example/.env and restart.`);
        }

        request.app.connector = connectorHeader;
        request.app.auth = internals.parseAuthHeader(request.headers['x-unimail-auth']);

        return h.continue;
    });

    server.ext('onPreResponse', (request, h) => {

        const response = request.response;
        const authUpdated = request.app.authUpdated;

        if (authUpdated && !response.isBoom) {
            response.header('X-Unimail-Auth-Updated', internals.encodeAuthHeader(authUpdated));
            response.header('Access-Control-Expose-Headers', 'X-Unimail-Auth-Updated');
        }

        return h.continue;
    });

    server.route({
        method: 'GET',
        path: '/connectors',
        handler: () => Object.keys(connectors)
    });

    server.route(Folders.routes(internals));
    server.route(Messages.routes(internals));

    server.route({
        method: 'GET',
        path: '/{param*}',
        handler: {
            directory: {
                path: Path.join(__dirname, 'public'),
                index: ['index.html']
            }
        }
    });

    await server.start();

    const configured = Object.keys(connectors);
    console.log(`Unimail viewer listening on http://${server.info.host}:${server.info.port}`);
    if (configured.length === 0) {
        console.log('WARNING: no connectors are configured. Copy example/.env.example to example/.env and fill in at least one provider.');
    }
    else {
        console.log(`Configured connectors: ${configured.join(', ')}`);
    }
};

process.on('unhandledRejection', (err) => {

    console.error(err);
    process.exit(1);
});

init();
