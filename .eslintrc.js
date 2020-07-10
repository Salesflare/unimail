'use strict';

module.exports = {
    env: {
        node: true
    },
    plugins: ['node'],
    extends: ['plugin:node/recommended', 'eslint-config-salesflare'],
    rules: {
        // Taken from server
        'node/callback-return': [
            'error',
            [
                'callback',
                'next',
                'done'
            ]
        ],
        'node/file-extension-in-import': ['error', 'always', { '.js': 'never' }],
        'node/global-require': 'error',
        'node/handle-callback-err': 'error',
        'node/no-exports-assign': 'error',
        'node/no-missing-import': 'error',
        'node/no-sync': [
            'error',
            {
                allowAtRootLevel: true
            }
        ],
        'node/no-unpublished-require': ['error'],
        'no-underscore-dangle': ['error', { allow: ['_dirty', '_deleted'], allowAfterThis: true }], // We do allow after this unlike the server
        // Custom rules
        'class-methods-use-this': 'off'
    }
};
