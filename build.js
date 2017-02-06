/*eslint no-console: "off" */
/*eslint no-sync: "off"*/
/*eslint no-process-exit: "off"*/

'use strict';

const fs = require('fs');
const browserify = require('browserify');
const uglifyJS = require('uglify-js');

const license = fs.readFileSync('license/license_header.txt', 'utf8');

writeBundle(writeMinify);

function writeBundle(next) {
	browserify('./src/index.js', {standalone: 'excelWriter'})
		.ignore('stream')
		.transform('eslintify')
		.transform('browserify-shim')
		.transform('babelify', {presets: ['es2015'], plugins: ['transform-runtime']})
		.bundle((err, bundle) => {
			if (err) {
				console.log(err);
				process.exit(1);
			}
			fs.writeFileSync('dist/excelWriter.js', license + bundle.toString());
			next();
		});
}

function writeMinify() {
	const minify = uglifyJS.minify('dist/excelWriter.js');

	fs.writeFileSync('dist/excelWriter.min.js', license + minify.code);
}
