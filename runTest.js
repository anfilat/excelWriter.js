/*eslint no-sync: "off"*/
/*eslint global-require: "off"*/

'use strict';

const fs = require('fs');
const path = require('path');
let excelWriter;

function compare(test, fileName) {
	const testName = path.parse(fileName).name;
	const xlsxFileName = `xlsx/${testName}.xlsx`;
	const workbook = test(excelWriter);

	return workbook.save({type: 'nodebuffer'})
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
	const xlsxFileName = `xlsx/${testName}.xlsx`;
	const workbook = test(excelWriter);

	try {
		fs.statSync('xlsx');
	} catch (e) {
		fs.mkdirSync('xlsx');
	}

	return workbook.saveAsNodeStream()
		.pipe(fs.createWriteStream(xlsxFileName));
}

module.exports = function (testDist) {
	if (testDist) {
		global._ = require('lodash');
		global.JSZip = require('jszip');
		excelWriter = require('./dist/excelWriter.min.js');

		global.JSZip.defaults.date = new Date('2016-01-04');
	} else {
		const JSZip = require('jszip');
		excelWriter = require('./src');

		JSZip.defaults.date = new Date('2016-01-04');
	}
	return {compare, write};
};
