import js from '@eslint/js'
import globals from 'globals'
import tseslint from 'typescript-eslint'

export default tseslint.config(
  { ignores: ['build', 'node_modules', 'logs'] },
  {
    extends: [js.configs.recommended, ...tseslint.configs.recommended],
  files: ['src/**/*.{ts,js}'],
    languageOptions: {
      ecmaVersion: 2020,
      globals: globals.node,
    },
    rules: {
      // Be permissive for MCP transport glue
      '@typescript-eslint/no-explicit-any': 'off',
      // Allow empty catch blocks for safe no-op fallbacks
      'no-empty': ['error', { allowEmptyCatch: true }],
      // Don't force ts-expect-error vs ts-ignore
      '@typescript-eslint/ban-ts-comment': 'off',
    },
  },
  // JS-only overrides (legacy helpers)
  {
    files: ['src/**/*.js'],
    rules: {
      '@typescript-eslint/no-require-imports': 'off',
      '@typescript-eslint/no-unused-vars': 'off',
    },
  },
  // Stricter rules for our active TypeScript sources
  {
    files: ['src/**/*.ts'],
    rules: {
      '@typescript-eslint/no-explicit-any': 'warn',
      '@typescript-eslint/ban-ts-comment': ['warn'],
  '@typescript-eslint/no-unused-vars': ['error', { argsIgnorePattern: '^_' }],
    },
  },
)
