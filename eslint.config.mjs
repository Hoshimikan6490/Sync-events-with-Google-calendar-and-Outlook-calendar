import js from '@eslint/js';
import json from '@eslint/json';
import globals from 'globals';
import prettier from 'eslint-config-prettier';

export default [
	// ============================================================
	// Ignore
	// ============================================================
	{
		ignores: ['node_modules/**', '.env', '*.min.js', 'dist/**', 'build/**'],
	},

	// ============================================================
	// JavaScript
	// ============================================================
	{
		files: ['**/*.js'],
		...js.configs.recommended,
		languageOptions: {
			ecmaVersion: 2022,
			sourceType: 'module',

			globals: {
				...globals.node,
			},
		},

		rules: {
			// 基本的なコード品質
			'no-unused-vars': [
				'error',
				{
					argsIgnorePattern: '^_',
					varsIgnorePattern: '^_',
				},
			],
			'no-console': 'off',
			'no-debugger': 'error',
			'no-alert': 'error',

			// 非同期処理
			'no-async-promise-executor': 'error',
			'prefer-promise-reject-errors': 'error',
			'require-await': 'error',

			// ベストプラクティス
			eqeqeq: ['error', 'always'],
			'no-var': 'error',
			'prefer-const': 'error',
			'prefer-arrow-callback': 'error',

			// セキュリティ
			'no-eval': 'error',
			'no-implied-eval': 'error',
			'no-new-func': 'error',
		},
	},

	// ============================================================
	// Google Apps Script
	// ============================================================
	{
		files: ['**/*.gs'],

		...js.configs.recommended,

		languageOptions: {
			ecmaVersion: 2022,
			sourceType: 'script',

			globals: {
				...globals.googleappsscript,
			},
		},

		rules: {
			// 基本的なコード品質
			'no-unused-vars': [
				'error',
				{
					argsIgnorePattern: '^_',
					varsIgnorePattern: '^_',
				},
			],
			'no-console': 'off',
			'no-debugger': 'error',
			'no-alert': 'error',

			// 非同期処理
			'no-async-promise-executor': 'error',
			'prefer-promise-reject-errors': 'error',
			'require-await': 'error',

			// ベストプラクティス
			eqeqeq: ['error', 'always'],
			'no-var': 'error',
			'prefer-const': 'error',
			'prefer-arrow-callback': 'error',

			// セキュリティ
			'no-eval': 'error',
			'no-implied-eval': 'error',
			'no-new-func': 'error',
		},
	},

	// ============================================================
	// JSON
	// ============================================================
	{
		files: ['**/*.json'],

		plugins: {
			json,
		},

		language: 'json/json',
	},

	// ============================================================
	// Prettier
	// ============================================================
	prettier,
];
