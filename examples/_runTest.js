/*eslint no-console: "off" */
/*eslint no-sync: "off"*/

'use strict';

const excelWriter = require('../src');
const fs = require('fs');
const path = require('path');

module.exports = function (test) {
	const fileName = path.parse(process.argv[1]).name;
	const workbook = test(excelWriter);

	try {
		fs.statSync('xlsx');
	} catch (e) {
		fs.mkdirSync('xlsx');
	}

	workbook.saveAsNodeStream()
		.pipe(fs.createWriteStream(`xlsx/${fileName}.xlsx`))
		.on('finish', () => {
			console.log('Excel file written.');
		});
};
