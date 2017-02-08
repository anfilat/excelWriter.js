'use strict';

module.exports = function (excel) {
	return excel.createWorkbook()
		.addWorksheet()
		.setColumns([
			{style: {border: {color: '#000000', style: 'thin'}}}
		])
		.setData([
			[
				{value: 'First'}
			],
			[
				{value: 'Second', style: {border: {color: '#FF0000', style: 'thick'}}}
			]
		])
		.end();
};
