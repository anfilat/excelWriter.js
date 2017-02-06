/*eslint no-console: "off" */
/*eslint global-require: "off"*/

'use strict';

const {compare, write} = require('../runTest')();

const fileName = process.argv[2];
const compareIt = process.argv[3];

if (fileName) {
	const test = require(`./${fileName}`);

	if (compareIt) {
		compare(test, fileName)
			.then(testName => {
				console.log(testName, 'passed');
			})
			.catch(error => {
				console.log(error.message, 'failed');
			});
	} else {
		write(test, fileName)
			.on('finish', () => {
				console.log('Excel file written.');
			});
	}
} else {
	console.log('node _test testName.js');
}
