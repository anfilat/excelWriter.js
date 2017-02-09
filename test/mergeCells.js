'use strict';

module.exports = function (excel) {
	const row = [1234, {value: 'Incididunt et sunt ipsum aliqua excepteur nostrud.', colspan: 2}, 'Lorem nulla'];
	const bottomRow = [3, 'elit culpa'];

	return excel.createWorkbook()
		.addWorksheet()
		.setData([
			[2541, 'Aliquip aliqua ex magna pariatur in enim.', 260, 'duis laborum'],
			[2542, {value: 'Elit dolore in eiusmod exercitation reprehenderit eu.', colspan: 2, rowspan: 3}, 'dolor ut'],
			bottomRow,
			bottomRow
		])
		.setData([
			row,
			row
		])
		.setData([
			bottomRow,
			[2542, {value: 'Elit dolore in eiusmod exercitation reprehenderit eu.', colspan: 2, rowspan: 3}, 'dolor ut']
		])
		.setColumns([
			{type: 'number', width: 10},
			{type: 'string', width: 60},
			{type: 'number'},
			{type: 'string', width: 20}
		])
		.end();
};
