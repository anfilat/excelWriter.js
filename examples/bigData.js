/*eslint no-console: "off" */
/*eslint no-sync: "off"*/

//stress test

'use strict';

var fs = require('fs');
var path = require('path');
var _ = require('lodash');
var loremIpsum = require('lorem-ipsum');
var excel = require('../src');
var testData;
var fileName;

fileName = path.parse(process.argv[1]).name;
try {
	fs.statSync('xlsx');
} catch (e) {
	fs.mkdirSync('xlsx');
}

console.time('Generate test data');
testData = generateData();
console.timeEnd('Generate test data');

console.log('Memory before - ', process.memoryUsage());
console.time('time');

excel.saveAsNodeStream(run(testData))
	.pipe(fs.createWriteStream('xlsx/' + fileName + '.xlsx'))
	.on('finish', function () {
		console.timeEnd('time');
		console.log('Memory after - ', process.memoryUsage());
		console.log('Excel file written.');
	});

function generateData() {
	var size = 1000000;
	var testData = new Array(size);

	_.times(size, function (i) {
		testData[i] = [
			1000 + i,
			_.capitalize(loremIpsum({count: 7, units: 'words'}) + '.'),
			Math.round(100 + Math.random() * (1000 - 100)),
			loremIpsum({count: 2, units: 'words'}),
			Math.round(1410000000000 + Math.random() * (1460000000000 - 1410000000000)),
			Math.round(1410000000000 + Math.random() * (1460000000000 - 1410000000000))
		];
	});
	return testData;
}

function run(testData) {
	var workbook = excel.createWorkbook();
	var header = workbook.addFormat({
		font: {bold: true, underline: true, color: {theme: 3}},
		horizontal: 'center',
		border: {
			bottom: {color: 'FFFF0000', style: 'thin'},
			top: {color: 'FFFF0000', style: 'thin'},
			left: {color: 'FFFF0000', style: 'thin'},
			right: {color: 'FFFF0000', style: 'thin'}
		},
		pattern: {color: 'FF00FF00'}
	});
	var fillFormat = workbook.addFormat({
		font: {
			italic: true,
			color: 'FFFF0000'
		},
		pattern: {color: 'FF00FF00'}
	});
	var currency = workbook.addFormat({
		format: '$ #,##0.00;$ #,##0.00;-',
		font: {color: 'FFE9F50A'}
	});
	workbook.addFormat({format: 'date'}, 'date');
	var columns = [
		{type: 'number', width: 10},
		{type: 'string', width: 60},
		{type: 'number', style: currency},
		{type: 'string', width: 20},
		{type: 'date', style: 'date', width: 12},
		{type: 'date', style: 'date', width: 12}
	];
	var worksheetData = [
		[
			{value: 'ID', style: header, type: 'string'},
			{value: 'Name', style: header, type: 'string'},
			{value: 'Price', style: header, type: 'string'},
			{
				value: 'Location',
				style: header,
				type: 'string',
				hyperlink: {location: 'http://ya.ru', tooltip: 'yandex'}
			},
			{value: 'Start Date', style: header, type: 'string'},
			{value: 'End Date', style: header, type: 'string'}
		]
	].concat(testData);

	workbook.addWorksheet({name: 'table'})
		.setHeader([{bold: true, text: 'Generic Report'}, '', ''])
		.setFooter(['', '', 'Page &P of &N'])
		.setData(worksheetData)
		.setColumns(columns)
		.setRow(2, {
			height: 30,
			style: fillFormat
		})
		.setHyperlink({
			cell: 'B2',
			location: 'http://www.google.com',
			tooltip: 'Click me!'
		})
		.setHyperlink({
			cell: {c: 2, r: 3},
			location: 'http://www.google.com',
			tooltip: 'Click me!'
		});

	return workbook;
}
