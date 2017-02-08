'use strict';

module.exports = function (excel) {
	return excel.createWorkbook()
		.addWorksheet()
		.setColumns([
			{type: 'number'},
			{type: 'number'}
		])
		.setData([
			{
				data: ['First', 'Second'],
				type: 'string'
			},
			[42, 428]
		])
		.end();
};
