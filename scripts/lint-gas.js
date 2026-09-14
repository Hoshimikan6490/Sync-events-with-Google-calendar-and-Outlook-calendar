const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const projectRoot = path.resolve(__dirname, '..');
const tempDirectory = path.join(projectRoot, 'scripts', 'temp');
const tempFile = path.join(tempDirectory, 'gas-project.js');

/**
 * GASプロジェクトのディレクトリを指定する。
 *
 * 例:
 *   gas/
 *   ├── main.gs
 *   ├── calendar.gs
 *   └── utils.gs
 */
const gasDirectory = path.join(projectRoot);

/**
 * 指定ディレクトリ以下の.gsファイルを取得する。
 *
 * @param {string} directory
 * @returns {string[]}
 */
function getGsFiles(directory) {
	return fs
		.readdirSync(directory, { withFileTypes: true })
		.filter((entry) => entry.isFile() && entry.name.endsWith('.gs'))
		.map((entry) => path.join(directory, entry.name))
		.sort();
}

/**
 * 複数の.gsファイルを1つのJSファイルに結合する。
 *
 * @param {string[]} files
 * @returns {string}
 */
function combineGsFiles(files) {
	return files
		.map((file) => {
			const relativePath = path.relative(projectRoot, file);
			const source = fs.readFileSync(file, 'utf8');

			return [
				`// ============================================================`,
				`// ${relativePath}`,
				`// ============================================================`,
				'',
				source,
				'',
			].join('\n');
		})
		.join('\n');
}

/**
 * メイン処理
 */
function main() {
	if (!fs.existsSync(gasDirectory)) {
		console.error(`GAS directory not found: ${gasDirectory}`);
		process.exitCode = 1;
		return;
	}

	const files = getGsFiles(gasDirectory);

	if (files.length === 0) {
		console.error(`No .gs files found: ${gasDirectory}`);
		process.exitCode = 1;
		return;
	}

	fs.mkdirSync(tempDirectory, { recursive: true });

	const combinedSource = combineGsFiles(files);

	fs.writeFileSync(tempFile, combinedSource, 'utf8');

	console.log(`Combined ${files.length} .gs files:`);
	for (const file of files) {
		console.log(`  - ${path.relative(projectRoot, file)}`);
	}

	console.log('');
	console.log(`Linting: ${path.relative(projectRoot, tempFile)}`);

	try {
		execFileSync(
			process.platform === 'win32' ? 'npx.cmd' : 'npx',
			['eslint', tempFile],
			{
				cwd: projectRoot,
				stdio: 'inherit',
			},
		);
	} catch (error) {
		process.exitCode = error.status ?? 1;
	}
}

main();
