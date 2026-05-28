'use strict';

const handler = (internals) => {

    return async (request) => {

        const { data, authUpdated } = await internals.callUnimail(request, 'folders', 'list', request.app.auth, {}, {});
        request.app.authUpdated = authUpdated;

        return data;
    };
};

exports.routes = (internals) => {

    return [{
        method: 'GET',
        path: '/api/folders',
        handler: handler(internals)
    }];
};
