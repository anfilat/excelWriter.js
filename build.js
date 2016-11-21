/*eslint no-console: "off" */
/*eslint no-sync: "off"*/
/*eslint no-process-exit: "off"*/

'use strict';

var fs = require('fs');
var browserify = require('browserify');
var uglifyJS = require('uglify-js');

var license = fs.readFileSync('license/license_header.txt', 'utf8');

writeBundle(writeMinify);

function writeBundle(next) {
	browserify('./src/index.js', {standalone: 'excelWriter'})
		.ignore('stream')
		.bundle(function (err, bundle) {
			if (err) {
				console.log(err);
				process.exit(1);
			}
			fs.writeFileSync('dist/excelWriter.js', license + bundle.toString());
			next();
		});
}

function writeMinify() {
	var minify = uglifyJS.minify('dist/excelWriter.js');

	fs.writeFileSync('dist/excelWriter.min.js', license + minify.code);
}
