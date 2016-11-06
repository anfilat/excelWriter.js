/*eslint no-console: "off" */
/*eslint no-sync: "off"*/

'use strict';

var excelWriter = require('../src');
var fs = require('fs');
var path = require('path');

module.exports = function (test) {
	var fileName = path.parse(process.argv[1]).name;
	var workbook = test(excelWriter);

	try {
		fs.statSync('xlsx');
	} catch (e) {
		fs.mkdirSync('xlsx');
	}

	excelWriter.saveAsNodeStream(workbook)
		.pipe(fs.createWriteStream('xlsx/' + fileName + '.xlsx'))
		.on('finish', function () {
			console.log('Excel file written.');
		});
};
