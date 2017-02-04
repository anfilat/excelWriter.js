/*eslint no-sync: "off"*/

'use strict';

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const excelWriter = require('./src');

function compare(test, fileName) {
	const testName = path.parse(fileName).name;
	const xlsxFileName = 'xlsx/' + testName + '.xlsx';
	const workbook = test(excelWriter);

	JSZip.defaults.date = new Date('2016-01-04');

	return excelWriter.save(workbook, {type: 'nodebuffer'})
		.then(result => {
			const example = fs.readFileSync(xlsxFileName);

			if (example.equals(result)) {
				return testName;
			} else {
				throw new Error(testName);
			}
		});
}

function write(test, fileName) {
	const testName = path.parse(fileName).name;
	const xlsxFileName = 'xlsx/' + testName + '.xlsx';
	const workbook = test(excelWriter);

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
	write,
	compare
};
