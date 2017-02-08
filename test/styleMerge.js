'use strict';

module.exports = function (excel) {
	return excel.createWorkbook()
		.addWorksheet()
		.setColumns([
			{style: {border: {color: '#000000', style: 'thin'}}}
		])
		.setData([
			[
				{value: 'Text 1'}
			],
			[
				{value: 'Text2', style: {border: {color: '#FF0000', style: 'thick'}}}
			],
			[
				{value: 'Text3', style: {border: {color: '#00FF00'}}}
			],
			[
				{value: 'Text4', style: {border: {style: 'thick'}}}
			]
		])
		.end();
};
