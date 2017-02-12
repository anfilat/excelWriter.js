'use strict';

module.exports = function (excel) {
	const row = [1234, {value: 'Incididunt et sunt ipsum aliqua excepteur nostrud.', colspan: 2}, 'Lorem nulla'];
	const bottomRow = [3, 'elit culpa'];

	return excel.createWorkbook()
		.addWorksheet()
		.setColumns([
			{type: 'number', width: 10},
			{type: 'string', width: 60},
			{type: 'number'},
			{type: 'string', width: 20}
		])
		.setData([
			[2541, 'Aliquip aliqua ex magna pariatur in enim.', 260, 'duis laborum'],
			[
				2542,
				{value: 'Elit dolore in eiusmod exercitation reprehenderit eu.', colspan: 2, rowspan: 3, style: {vertical: 'top'}},
				'dolor ut'
			],
			bottomRow,
			bottomRow
		])
		.setData([
			row,
			row
		])
		.setData([
			{
				style: {pattern: {color: '#b0ffff'}, border: {color: '#000000', style: 'thin'}},
				data: [
					123,
					['Officia deserunt elit', 'Pariatur elit duis', 'Sunt qui ipsum'],
					542,
					{value: 'Mollit aliqua', style: {vertical: 'center'}}
				]
			},
			{
				style: {pattern: {color: '#b0ffb0'}, border: {color: '#000000', style: 'thin'}},
				data: [
					245,
					{value: ['Esse consectetur ex', 'Dolore id sint', 'Anim irure pariatur'], style: {pattern: {color: '#b0b0ff'}}},
					{value: [358, 864], style: {pattern: {color: '#ffb0ff'}}},
					{value: 'Minim nisi', style: {vertical: 'center'}}
				]
			}
		])
		.setData([
			bottomRow,
			{
				data: [
					2547,
					{
						value: 'Ea commodo nostrud incididunt incididunt qui in.',
						colspan: 2,
						rowspan: 2,
						style: {pattern: {color: '#ffffb0'}, border: {color: '#000000', style: 'thin'}, vertical: 'top'}
					},
					'ex magna'
				]
			}
		])
		.end();
};
