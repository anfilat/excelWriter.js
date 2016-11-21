/*eslint no-console: "off" */
/*eslint no-sync: "off"*/
/*eslint global-require: "off"*/
/*eslint no-process-exit: "off"*/

'use strict';

var fs = require('fs');
var _ = require('lodash');
var runTest = require('./runTest');

process.chdir('./test');

_(fs.readdirSync('.'))
	.filter(function (fileName) {
		return !_.startsWith(fileName, '_') && _.endsWith(fileName, '.js');
	})
	.forEach(function (fileName) {
		var test = require('./test/' + fileName);

		runTest.compare(test, fileName)
			.catch(function (error) {
				console.log(error.message, 'failed');
			});
	});
