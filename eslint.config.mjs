import js from '@eslint/js';
import globals from 'globals';
import prettier from 'eslint-config-prettier';

export default [
	js.configs.recommended,
	prettier,
	{
		languageOptions: {
			ecmaVersion: 2022,
			sourceType: 'module',
		},
		rules: {
			// 基本的なコード品質ルール
			'no-unused-vars': ['error', { argsIgnorePattern: '^_' }],
			'no-console': 'off', // ログ出力のため許可
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
	{
		// 特定のファイルタイプの設定（CommonJS）
		files: ['**/*.js'],
		languageOptions: {
			sourceType: 'commonjs',
			globals: {
				...globals.node,
			},
		},
	},
	{
		// 無視するファイル/ディレクトリ（.gitignoreの内容を含む）
		ignores: [
			// .gitignoreから取り込み
			'node_modules/**',
			'.env',
			'lib/pomodoro/voicevox/**',
			'!lib/pomodoro/voicevox/.VOICEVOX_FILES_IS_HERE',
			'assets/audio/**',
			'!assets/audio/voiceList.json',

			// ESLint用の追加設定
			'*.min.js',
			'dist/**',
			'build/**',
		],
	},
];
