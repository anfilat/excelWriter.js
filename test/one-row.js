'use strict';

module.exports = function (excel) {
	var workbook = excel.createWorkbook();
	workbook.addFontFormat({bold: true, underline: true, color: {theme: 3}}, 'font');
	var headerBorder = workbook.addBorderFormat({
		bottom: {color: 'FFFF0000', style: 'thin'},
		top: {color: 'FFFF0000', style: 'thin'},
		left: {color: 'FFFF0000', style: 'thin'},
		right: {color: 'FFFF0000', style: 'thin'}
	});
	var greenFill = workbook.addPatternFormat({color: 'FF00FF00'});
	var currency = workbook.addNumberFormat('$ #,##0.00;$ #,##0.00;-', 'currency');
	var header = workbook.addFormat({
		font: 'font',
		border: headerBorder,
		pattern: greenFill,
		format: currency,
		horizontal: 'center'
	});
	var columnFormat = workbook.addFormat({
		gradient: {left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, start: 'FFFFFF00', end: 'FF5B9BD5'}
	});
	var currencyFormat = workbook.addFormat({format: 'currency'});
	var dateFormat = workbook.addFormat({format: 'date', horizontal: 'right', indent: 1});
	var fillFormat = workbook.addFormat({
		font: {
			italic: true,
			color: 'FFFF0000'
		},
		pattern: {type: 'darkHorizontal', color: 'FF008800', backColor: 'FF000088'},
		locked: false,
		hidden: true
	});
	var columns = [
		{width: 10},
		{width: 40, style: columnFormat},
		{style: currencyFormat}
	];
	var worksheetData = [
		{
			outlineLevel: 0,
			height: 30,
			style: fillFormat,
			data: [{value: 2541, style: header}, 'Nullam aliquet mi et nunc tempus rutrum.', 260,
				'Dolore anim', 1342372604000, {date: 1342977404000}]
		},
		{
			outlineLevel: 1,
			data: [{value: 2541, style: header}, 'Nullam aliquet mi et nunc tempus rutrum.', 260,
				'Dolore anim', 1342372604000, {date: 1342977404000}]
		},
		[{value: 2541, style: header}, 'Nullam aliquet mi et nunc tempus rutrum.', 260,
			'Dolore anim', 1342372604000, {time: 1342977404000}],
		[{value: 2541, style: header}, 'Nullam aliquet mi et nunc tempus rutrum.', 260,
			'Dolore anim', 1342372604000, {time: 1342977404000}]
	];

	workbook.addWorksheet()
		.setData(worksheetData)
		.setData(7, worksheetData)
		.setColumns(columns)
		.setColumn(5, {width: 12, style: dateFormat, type: 'date'})
		.setRow(2, {height: 30})
		.setRows(3, [{height: 30}, {height: 30}]);

	return workbook;
};
