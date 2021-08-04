/**
 * @type {import('eslint').Linter.Config}
 */
const config = {
  parser: '@typescript-eslint/parser',
  extends: [
    'plugin:@typescript-eslint/recommended'
  ],
  parserOptions: {
    ecmaVersion: 2018,
    sourceType: 'module'
  },
  rules: {
    quotes: [2, 'single', { 'allowTemplateLiterals': true }],
    indent: ['error', 2, { 'ObjectExpression': 1 }],
    'arrow-parens': ['error', 'always'],
    'no-var-requires': 0
  },
  settings: {
    react: {
      version: 'detect'
    }
  }
};

module.exports = config;
