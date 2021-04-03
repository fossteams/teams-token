module.exports = {
    root: true,
  
    env: {
      node: true,
    },

    parser: '@typescript-eslint/parser',
    "extends": [
      'airbnb',
      'eslint:recommended'
    ],
  
    parserOptions: {
      ecmaVersion: 2020,
    },
  
    rules: {
      'no-console': 'off',
      'no-debugger': 'off',
    },
  };
  