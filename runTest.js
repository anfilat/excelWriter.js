/*eslint no-sync: "off"*/

'use strict';

var fs = require('fs');
var path = require('path');
var JSZip = require('jszip');
var excelWriter = require('./src');

function compare(test, fileName) {
	var testName = path.parse(fileName).name;
	var xlsxFileName = 'xlsx/' + testName + '.xlsx';
	var workbook = test(excelWriter);

	JSZip.defaults.date = new Date('2016-01-04');

	return excelWriter.save(workbook, {type: 'nodebuffer'})
		.then(function (result) {
			var example = fs.readFileSync(xlsxFileName);

			if (example.equals(result)) {
				return testName;
			} else {
				throw new Error(testName);
			}
		});
}

function write(test, fileName) {
	var testName = path.parse(fileName).name;
	var xlsxFileName = 'xlsx/' + testName + '.xlsx';
	var workbook = test(excelWriter);

	JSZip.defaults.date = new Date('2016-01-04');

	try {
		fs.statSync('xlsx');
	} catch (e) {
		fs.mkdirSync('xlsx');
	}

	return excelWriter.saveAsNodeStream(workbook)
		.pipe(fs.createWriteStream(xlsxFileName));
}

module.exports = {
	write: write,
	compare: compare
};
