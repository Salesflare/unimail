'use strict';

module.exports = {
    env: {
        node: true
    },
    plugins: ['node', 'dependencies'],
    extends: ['plugin:node/recommended', 'eslint-config-salesflare'],
    rules: {
        // Taken from server
        'global-require': 'error',
        'handle-callback-err': 'error',
        'dependencies/no-cycles': 1,
        'node/no-missing-import': 'error',
        'node/no-unpublished-require': ['error'],
        'node/shebang': 'off',
        'no-underscore-dangle': ['error', { allow: ['_dirty', '_deleted'], allowAfterThis: true }],
        'valid-jsdoc': [
            'error',
            {
                prefer: { 'return': 'returns' },
                preferType: {
                    'object': 'Object',
                    'number': 'Number',
                    'string': 'String',
                    'boolean': 'Boolean',
                    'array': 'Array'
                },
                'requireReturn': true,
                'requireParamDescription': false,
                'requireReturnDescription': false,
                requireParamType: true
            }
        ],
        // Custom rules
        'class-methods-use-this': 'off'
    }
};
