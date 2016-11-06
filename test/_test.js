/*eslint no-console: "off" */
/*eslint global-require: "off"*/
/*eslint no-sync: "off"*/

'use strict';

var JSZip = require('jszip');
var excelWriter = require('../src');
var fs = require('fs');
var path = require('path');
var fileName = process.argv[2];
var compare = process.argv[3];

if (fileName) {
	var test = require('./' + fileName);

	try {
		fs.statSync('xlsx');
	} catch (e) {
		fs.mkdirSync('xlsx');
	}
	runTest(test);
} else {
	console.log('node _test testName.js');
}

function runTest(test) {
	var testName = path.parse(fileName).name;
	var xlsxFileName = 'xlsx/' + testName + '.xlsx';
	var workbook = test(excelWriter);

	JSZip.defaults.date = new Date('2016-01-04');

	if (compare) {
		excelWriter.save(workbook, {type: 'nodebuffer'})
		.then(function (result) {
			var example = fs.readFileSync(xlsxFileName);

			console.log(testName, example.equals(result));
		});
	} else {
		excelWriter.saveAsNodeStream(workbook)
			.pipe(fs.createWriteStream(xlsxFileName))
			.on('finish', function () {
				console.log('Excel file written.');
			});
	}
}
