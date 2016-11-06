'use strict';

module.exports = function (excel) {
	var data = [
		['Artist', 'Album', 'Price'],
		['Buckethead', 'Albino Slug', 8.99],
		['Buckethead', 'Electric Tears', 13.99],
		['Buckethead', 'Colma', 11.34],
		['Crystal Method', 'Vegas', 10.54],
		['Crystal Method', 'Tweekend', 10.64],
		['Crystal Method', 'Divided By Night', 8.99]
	];
	var workbook = excel.createWorkbook();
	var wholeTable = workbook.addTableElementFormat({
		pattern: {color: 'FF0088FF', type: 'solid'},
		font: {italic: true, color: 'FF880000'}
	});

	workbook.addTableElementFormat({
		horizontal: 'center',
		font: {bold: true, italic: false}
	}, 'headerRow');

	workbook.addTableFormat({
		wholeTable: wholeTable,
		headerRow: 'headerRow',
		totalRow: {gradient: {degree: 0, start: 'FF778877', end: 'FF558899'}},
		firstRowStripe: {style: {pattern: {color: 'FF338833'}}, size: 2},
		secondRowStripe: {pattern: {color: 'FF3388CC'}}
	}, 'SlightlyOffColorBlue');

	workbook.setDefaultTableStyle('SlightlyOffColorBlue');

	workbook.addWorksheet({name: 'Album List'})
		.setData(data)
		.addTable()
			.setTheme('SlightlyOffColorBlue')
			.setReferenceRange('A1', {c: 3, r: data.length})
			.addTotalRow([
				{totalsRowLabel: 'Highest Price'},
				{totalsRowLabel: 'test'},
				{totalsRowFunction: 'max'}
			])
			.end()
		.setColumns([
			{width: 20},
			{width: 20}
		]);

	return workbook;
};
