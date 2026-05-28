'use strict';

const list = (internals) => {

    return async (request) => {

        const params = {};
        const query = request.query;

        if (query.folder) {
            params.folder = query.folder;
        }

        if (query.limit) {
            params.limit = Number(query.limit);
        }

        if (query.pageToken) {
            params.pageToken = query.pageToken;
        }

        if (query.subject) {
            params.subject = query.subject;
        }

        if (query.after) {
            const after = new Date(query.after);
            if (!Number.isNaN(after.getTime())) {
                params.after = after;
            }
        }

        if (query.before) {
            const before = new Date(query.before);
            if (!Number.isNaN(before.getTime())) {
                params.before = before;
            }
        }

        if (query.participants) {
            const raw = Array.isArray(query.participants) ? query.participants : [query.participants];
            const participants = raw
                .flatMap((p) => String(p).split(/[\n,]/))
                .map((s) => s.trim())
                .filter(Boolean);

            if (participants.length > 0) {
                params.participants = participants;
            }
        }

        if (query.includeDrafts !== undefined) {
            params.includeDrafts = query.includeDrafts === 'true' || query.includeDrafts === true;
        }

        const { data, authUpdated } = await internals.callUnimail(request, 'messages', 'list', request.app.auth, params, { includeBody: false });
        request.app.authUpdated = authUpdated;

        return data;
    };
};

const get = (internals) => {

    return async (request) => {

        const params = { id: request.params.id };
        const options = {};

        if (request.query.raw === 'true' || request.query.raw === true) {
            options.raw = true;
        }

        const { data, authUpdated } = await internals.callUnimail(request, 'messages', 'get', request.app.auth, params, options);
        request.app.authUpdated = authUpdated;

        return data;
    };
};

exports.routes = (internals) => {

    return [
        {
            method: 'GET',
            path: '/api/messages',
            handler: list(internals)
        },
        {
            method: 'GET',
            path: '/api/messages/{id}',
            handler: get(internals)
        }
    ];
};
