'use strict';

module.exports = function (excel) {
	return excel.createWorkbook()
		.addWorksheet()
		.setColumns([
			{width: 20, style: {border: {color: '#000000', style: 'thin'}, pattern: {color: '#b0ffff'}}},
			{width: 20, style: {border: {color: '#000000', style: 'thin'}, pattern: {color: '#b0ffff'}, horizontal: 'right'}}
		])
		.setData([
			{
				skipColumnsStyle: true,
				data: ['Header 1', 'Header 2']
			},
			['Text 1', 'Text 2'],
			['Text 3', 'Text 4']
		])
		.end();
};
