'use strict';

const Readable = require('stream').Readable;
const _ = require('lodash');
const util = require('../util');
const toXMLString = require('../XMLString');

const methods = {
	save() {
		const canStream = !!Readable;

		if (canStream) {
			return new WorksheetStream({
				worksheet: this
			});
		}
		return saveBeforeRows(this) +
			saveData(this) +
			saveAfterRows(this);
	}
};

class WorksheetStream extends (Readable || null) {
	constructor(options) {
		super(options);
		this.worksheet = options.worksheet;
		this.status = 0;
		this.index = 0;
		this.len = this.worksheet.preparedData.length;
	}
	_read(size) {
		let stop = false;

		if (this.status === 0) {
			stop = !this.push(saveBeforeRows(this.worksheet));

			this.status = 1;
		}

		if (this.status === 1) {
			while (this.index < this.len && !stop) {
				stop = !this.push(this.packChunk(size));
			}

			if (this.index === this.len) {
				this.status = 2;
			}
		}

		if (this.status === 2) {
			this.push(saveAfterRows(this.worksheet));
			this.push(null);
		}
	}
	packChunk(size) {
		const worksheet = this.worksheet;
		const data = worksheet.preparedData;
		const preparedRows = worksheet.preparedRows;
		let s = '';

		while (this.index < this.len && s.length < size) {
			s += saveRow(data[this.index], preparedRows[this.index], this.index);
			data[this.index] = null;
			this.index++;
		}
		return s;
	}
}

function saveBeforeRows(worksheet) {
	return util.xmlPrefix + '<worksheet xmlns="' + util.schemas.spreadsheetml +
		'" xmlns:r="' + util.schemas.relationships + '" xmlns:mc="' + util.schemas.markupCompat + '">' +
		saveDimension(worksheet.maxX, worksheet.maxY) +
		worksheet.sheetView.save() +
		saveColumns(worksheet.preparedColumns) +
		'<sheetData>';
}

function saveAfterRows(worksheet) {
	return '</sheetData>' +
		// 'mergeCells' should be written before 'headerFoot' and 'drawing' due to issue
		// with Microsoft Excel (2007, 2013)
		worksheet.mergedCells.save() +
		worksheet.hyperlinks.save() +
		worksheet.savePrint() +
		worksheet.tables.save() +
		// the 'drawing' element should be written last, after 'headerFooter', 'mergeCells', etc. due
		// to issue with Microsoft Excel (2007, 2013)
		worksheet.drawings.save() +
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
	if (dataRow) {
		const rowLen = dataRow.length;
		const rowChildren = new Array(rowLen);

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
		return '<row' + getRowAttributes(row, rowIndex) + '>' + rowChildren.join('') + '</row>';
	}
	return '<row' + getRowAttributes(row, rowIndex) + '></row>';
}

function getRowAttributes(row, rowIndex) {
	let attributes = ' r="' + (rowIndex + 1) + '"';

	if (row) {
		if (row.height !== undefined) {
			attributes += ' customHeight="1" ht="' + row.height + '"';
		}
		if (row.selfStyleId) {
			attributes += ' customFormat="1" s="' + row.selfStyleId + '"';
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
		const children = columns.map((column, index) => {
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
			if (column.selfStyleId) {
				attributes.push(['style', column.selfStyleId]);
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

module.exports = {methods};
