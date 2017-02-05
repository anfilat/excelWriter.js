/*eslint no-console: "off" */
/*eslint no-sync: "off"*/
/*eslint global-require: "off"*/

'use strict';

const fs = require('fs');
const _ = require('lodash');
const runTest = require('./runTest');

process.chdir('./test');

_(fs.readdirSync('.'))
	.filter(fileName => !_.startsWith(fileName, '_') && _.endsWith(fileName, '.js'))
	.forEach(fileName => {
		const test = require(`./test/${fileName}`);

		runTest.compare(test, fileName)
			.catch(function (error) {
				console.log(error.message, 'failed');
			});
	});
