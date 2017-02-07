'use strict';

const Readable = require('stream').Readable;
const _ = require('lodash');
const PrepareSave = require('./prepareSave');
const util = require('../util');
const toXMLString = require('../XMLString');

class WorksheetSave extends PrepareSave {
	_save(canStream) {
		if (canStream) {
			return new WorksheetStream({
				worksheet: this
			});
		} else {
			return saveBeforeRows(this) +
				saveData(this) +
				saveAfterRows(this);
		}
	}
}

class WorksheetStream extends (Readable || null) {
	constructor(options) {
		super(options);
		this.worksheet = options.worksheet;
		this.status = 0;
	}
	_read(size) {
		const worksheet = this.worksheet;
		let stop = false;

		if (this.status === 0) {
			stop = !this.push(saveBeforeRows(worksheet));

			this.status = 1;
			this.index = 0;
			this.len = worksheet.preparedData.length;
		}

		if (this.status === 1) {
			const data = worksheet.preparedData;
			const preparedRows = worksheet.preparedRows;
			let s = '';

			while (this.index < this.len && !stop) {
				while (this.index < this.len && s.length < size) {
					s += saveRow(data[this.index], preparedRows[this.index], this.index);
					data[this.index] = null;
					this.index++;
				}
				stop = !this.push(s);
				s = '';
			}

			if (this.index === this.len) {
				this.status = 2;
			}
		}

		if (this.status === 2) {
			this.push(saveAfterRows(worksheet));
			this.push(null);
		}
	}
}

function saveBeforeRows(worksheet) {
	return util.xmlPrefix + '<worksheet xmlns="' + util.schemas.spreadsheetml +
		'" xmlns:r="' + util.schemas.relationships + '" xmlns:mc="' + util.schemas.markupCompat + '">' +
		saveDimension(worksheet.maxX, worksheet.maxY) +
		worksheet._saveSheetView() +
		saveColumns(worksheet.preparedColumns) +
		'<sheetData>';
}

function saveAfterRows(worksheet) {
	return '</sheetData>' +
		// 'mergeCells' should be written before 'headerFoot' and 'drawing' due to issue
		// with Microsoft Excel (2007, 2013)
		worksheet._saveMergeCells() +
		worksheet._saveHyperlinks() +
		worksheet._savePrint() +
		worksheet._saveTables() +
		// the 'drawing' element should be written last, after 'headerFooter', 'mergeCells', etc. due
		// to issue with Microsoft Excel (2007, 2013)
		worksheet._saveDrawing() +
		'</worksheet>';
}

function saveData(worksheet) {
	const data = worksheet.preparedData;
	const preparedRows = worksheet.preparedRows;
	let children = '';

	for (let i = 0, len = data.length; i < len; i++) {
		children += saveRow(data[i], preparedRows[i], i);
		data[i] = null;
	}
	return children;
}

function saveRow(dataRow, row, rowIndex) {
	let rowChildren = [];

	if (dataRow) {
		const rowLen = dataRow.length;
		rowChildren = new Array(rowLen);

		for (let colIndex = 0; colIndex < rowLen; colIndex++) {
			const value = dataRow[colIndex];

			if (!value) {
				continue;
			}

			let attrs = ' r="' + util.positionToLetter(colIndex + 1, rowIndex + 1) + '"';
			if (value.styleId) {
				attrs += ' s="' + value.styleId + '"';
			}
			if (value.isString) {
				attrs += ' t="s"';
			}

			if (!_.isNil(value.value)) {
				rowChildren[colIndex] = '<c' + attrs + '><v>' + value.value + '</v></c>';
			} else if (!_.isNil(value.formula)) {
				rowChildren[colIndex] = '<c' + attrs + '><f>' + value.formula + '</f></c>';
			} else {
				rowChildren[colIndex] = '<c' + attrs + '/>';
			}
		}
	}

	return '<row' + getRowAttributes(row, rowIndex) + '>' + rowChildren.join('') + '</row>';
}

function getRowAttributes(row, rowIndex) {
	let attributes = ' r="' + (rowIndex + 1) + '"';

	if (row) {
		if (row.height !== undefined) {
			attributes += ' customHeight="1"';
			attributes += ' ht="' + row.height + '"';
		}
		if (row.styleId) {
			attributes += ' customFormat="1"';
			attributes += ' s="' + row.styleId + '"';
		}
		if (row.outlineLevel) {
			attributes += ' outlineLevel="' + row.outlineLevel + '"';
		}
	}
	return attributes;
}

function saveDimension(maxX, maxY) {
	const attributes = [];

	if (maxX !== 0) {
		attributes.push(
			['ref', 'A1:' + util.positionToLetter(maxX, maxY)]
		);
	} else {
		attributes.push(
			['ref', 'A1']
		);
	}

	return toXMLString({
		name: 'dimension',
		attributes
	});
}

function saveColumns(columns) {
	if (columns.length) {
		const children = _.map(columns, (column, index) => {
			column = column || {};

			const attributes = [
				['min', column.min || index + 1],
				['max', column.max || index + 1]
			];

			if (column.hidden) {
				attributes.push(['hidden', 1]);
			}
			if (column.bestFit) {
				attributes.push(['bestFit', 1]);
			}
			if (column.customWidth || column.width) {
				attributes.push(['customWidth', 1]);
			}
			if (column.width) {
				attributes.push(['width', column.width]);
			} else {
				attributes.push(['width', 9.140625]);
			}
			if (column.styleId) {
				attributes.push(['style', column.styleId]);
			}

			return toXMLString({
				name: 'col',
				attributes
			});
		});

		return toXMLString({
			name: 'cols',
			children
		});
	}
	return '';
}

module.exports = WorksheetSave;
