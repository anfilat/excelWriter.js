'use strict';

module.exports = function (excel) {
	const workbook = excel.createWorkbook();
	const worksheetData = [
		[
			1342372604000,
			1342372604000,
			{value: 1342372604000, type: 'date'},
			{value: 1342372604000, type: 'time'},
			{date: 1342372604000},
			{time: 1342372604000},
			new Date(2016, 11, 15, 11, 5),
			new Date(2016, 11, 15, 11, 5),
			{date: new Date(2016, 11, 15, 11, 5)},
			{time: new Date(2016, 11, 15, 11, 5)}
		]
	];

	workbook.addWorksheet()
		.setData(worksheetData)
		.setColumns([
			{width: 12, type: 'date'},
			{width: 12, type: 'time'},
			{width: 12},
			{width: 12},
			{width: 12},
			{width: 12},
			{width: 12},
			{width: 12, type: 'time'},
			{width: 12},
			{width: 12}
		]);

	return workbook;
};
