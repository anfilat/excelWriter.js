/*eslint no-console: "off" */
/*eslint global-require: "off"*/

'use strict';

var runTest = require('../runTest');

var fileName = process.argv[2];
var compare = process.argv[3];

if (fileName) {
	var test = require('./' + fileName);

	if (compare) {
		runTest.compare(test, fileName)
			.then(function (testName) {
				console.log(testName, 'passed');
			})
			.catch(function (error) {
				console.log(error.message, 'failed');
			});
	} else {
		runTest.write(test, fileName)
			.on('finish', function () {
				console.log('Excel file written.');
			});
	}
} else {
	console.log('node _test testName.js');
}
