excelWriter.js
================

A way to build excel files with javascript.

This project is fork of [excel-builder.js](https://github.com/stephenliberty/excel-builder.js). But it was rewrote in jQuery style). I hope that I'll write the documentation in a future)

Short example
---------------
    var JSZip = require('jszip');
    var excelWriter = require('excelWriter');
	var fs = require('fs');
    
    var workbook = excelWriter.createWorkbook();

   	workbook.addWorksheet()
   		.setData([
    		[2541, 'Nullam aliquet mi et nunc tempus rutrum.', 'Dolore anim', {date: 1342977404000}],
    		[2542, 'Nullam aliquet mi et nunc tempus rutrum.', 'Dolore anim', {date: 1342977405000}]
   		])
		.setColumns([
			{width: 10},
			{width: 40}
		]);
   		
	workbook.saveAsNodeStream()
		.pipe(fs.createWriteStream('./excel.xlsx'))


Distributables
---------------
excelWriter.js -> Requires lodash and jszip scripts to be loaded on the page.
excelWriter.min.js -> Minified excelWriter.js.
