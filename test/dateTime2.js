'use strict';

module.exports = function (excel) {
	const workbook = excel.createWorkbook();

	workbook.addNumberFormat('yyyy mm dd', 'date');
	workbook.addFormat({format: 'hh:mm'}, 'time');
	workbook.addFormat({format: 'mmmm'}, 'month');
	workbook.addFormat({format: 'hh\\h mm\\m ss\\s'}, 'time2');
	workbook.addWorksheet()
		.setData([
			[
				{date: 1342372604000},
				{time: 1342372604000},
				{date: 1342372604000},
				{time: 1342372604000}
			]
		])
		.setColumns([
			{width: 12},
			{width: 12},
			{width: 12, style: 'month'},
			{width: 12, style: 'time2'}
		]);

	return workbook;
};
