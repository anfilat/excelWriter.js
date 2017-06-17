/*eslint no-console: "off" */
/*eslint no-sync: "off"*/

//stress test

'use strict';

const fs = require('fs');
const path = require('path');
const _ = require('lodash');
const loremIpsum = require('lorem-ipsum');
const excel = require('../src');

const fileName = path.parse(process.argv[1]).name;
try {
	fs.statSync('xlsx');
} catch (e) {
	fs.mkdirSync('xlsx');
}

console.time('Generate test data');
const testData = generateData();
console.timeEnd('Generate test data');

console.log('Memory before - ', process.memoryUsage());
console.time('time');

run(testData).saveAsNodeStream()
	.pipe(fs.createWriteStream(`xlsx/${fileName}.xlsx`))
	.on('finish', () => {
		console.timeEnd('time');
		console.log('Memory after - ', process.memoryUsage());
		console.log('Excel file written.');
	});

function generateData() {
	const size = 1000000;
	const testData = new Array(size);

	for (let i = 0; i < size; i++) {
		testData[i] = [
			1000 + i,
			_.upperFirst(loremIpsum({count: 7, units: 'words'}) + '.'),
			Math.round(100 + Math.random() * (1000 - 100)),
			{value: loremIpsum({count: 2, units: 'words'}), style: {font: {bold: true}}},
			Math.round(1410000000000 + Math.random() * (1460000000000 - 1410000000000)),
			Math.round(1410000000000 + Math.random() * (1460000000000 - 1410000000000))
		];
	}
	return testData;
}

function run(testData) {
	const workbook = excel.createWorkbook();
	workbook.addFormat({format: 'date'}, 'date');
	const columns = [
		{type: 'number', width: 10},
		{type: 'string', width: 60},
		{
			type: 'number',
			style: {
				format: '$ #,##0.00;$ #,##0.00;-',
				font: {color: '#E90AF5'}
			}
		},
		{
			type: 'string',
			style: {font: {color: '#f56606'}},
			width: 20
		},
		{type: 'date', style: 'date', width: 12},
		{type: 'date', style: 'date', width: 12}
	];
	const worksheetData = [
		{
			style: {
				font: {bold: true, underline: true, color: {theme: 3}},
				horizontal: 'center',
				border: {color: '#FF0000', style: 'thin'},
				pattern: {color: '#C0F0C0'}
			},
			type: 'string',
			data: [
				'ID',
				'Name',
				'Price',
				{
					value: 'Location',
					hyperlink: {location: 'http://ya.ru', tooltip: 'yandex'}
				},
				'Start Date',
				'End Date'
			]
		}
	].concat(testData);

	workbook.addWorksheet({name: 'table'})
		.setHeader([{bold: true, text: 'Generic Report'}, '', ''])
		.setFooter(['', '', 'Page &P of &N'])
		.setData(worksheetData)
		.setColumns(columns)
		.setRow(2, {
			height: 30,
			style: {
				font: {
					italic: true,
					color: '#FF0000'
				},
				pattern: {color: '#00FF00'}
			}
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
