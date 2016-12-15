'use strict';

module.exports = function (excel) {
	var workbook = excel.createWorkbook();
	var worksheetData = [
		[
			{date: 1342372604000},
			{date: new Date(2016, 11, 15)},
			new Date(2016, 11, 15)
		]
	];

	workbook.addNumberFormat('yyyy.mm.dd', 'date');
	workbook.addWorksheet()
		.setData(worksheetData)
		.setColumn(1, {
			width: 12,
			style: {format: 'date'}
		})
		.setColumn(2, {
			width: 12,
			style: {format: 'date'}
		})
		.setColumn(3, {
			width: 12,
			style: {format: 'date'}
		});

	return workbook;
};
