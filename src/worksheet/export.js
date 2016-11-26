'use strict';

var Readable = require('stream').Readable;
var _ = require('lodash');
var util = require('../util');
var toXMLString = require('../XMLString');

function Export(worksheet) {
	this.worksheet = worksheet;
}

Export.prototype.export = function (canStream) {
	if (canStream) {
		return new WorksheetStream({
			exporter: this,
			worksheet: this.worksheet
		});
	} else {
		return this.exportBeforeRows() +
			this.exportData() +
			this.exportAfterRows();
	}
};

function WorksheetStream(options) {
	Readable.call(this, options);

	this.exporter = options.exporter;
	this.worksheet = options.worksheet;
	this.status = 0;
}

util.inherits(WorksheetStream, Readable || {});

WorksheetStream.prototype._read = function (size) {
	var stop = false;
	var exporter = this.exporter;
	var worksheet = this.worksheet;

	if (this.status === 0) {
		stop = !this.push(exporter.exportBeforeRows());

		this.status = 1;
		this.index = 0;
		this.len = worksheet.preparedData.length;
	}

	if (this.status === 1) {
		var s = '';
		while (this.index < this.len && !stop) {
			while (this.index < this.len && s.length < size) {
				s += exporter.exportRow(worksheet.preparedData[this.index], this.index, worksheet.preparedRows);
				worksheet.preparedData[this.index] = null;
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
		this.push(exporter.exportAfterRows());
		this.push(null);
	}
};

Export.prototype.exportBeforeRows = function () {
	var worksheet = this.worksheet;

	return this.getXMLBegin() +
		this.exportDimension(worksheet.maxX, worksheet.maxY) +
		worksheet.sheetView.export() +
		this.exportColumns(worksheet.columns) +
		'<sheetData>';
};

Export.prototype.exportAfterRows = function () {
	var worksheet = this.worksheet;

	return '</sheetData>' +
		// 'mergeCells' should be written before 'headerFoot' and 'drawing' due to issue
		// with Microsoft Excel (2007, 2013)
		worksheet.mergedCells.export() +
		worksheet.hyperlinks.export() +
		worksheet.print.export() +
		worksheet.tables.export() +
		// the 'drawing' element should be written last, after 'headerFooter', 'mergeCells', etc. due
		// to issue with Microsoft Excel (2007, 2013)
		worksheet.drawing.export() +
		this.getXMLEnd();
};

Export.prototype.exportData = function () {
	var data = this.worksheet.preparedData;
	var children = '';
	var dataRow;

	for (var i = 0, len = data.length; i < len; i++) {
		dataRow = data[i];
		children += this.exportRow(dataRow, i);
		data[i] = null;
	}
	return children;
};

Export.prototype.exportRow = function (dataRow, rowIndex) {
	var rowLen;
	var rowChildren = [];
	var colIndex;
	var value;
	var attrs;

	if (dataRow) {
		rowLen = dataRow.length;
		rowChildren = new Array(rowLen);

		for (colIndex = 0; colIndex < rowLen; colIndex++) {
			value = dataRow[colIndex];

			if (!value) {
				continue;
			}

			attrs = ' r="' + util.positionToLetter(colIndex + 1, rowIndex + 1) + '"';
			if (value.style) {
				attrs += ' s="' + value.style + '"';
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

	return '<row' + this.getRowAttributes(this.worksheet.preparedRows[rowIndex], rowIndex) + '>' +
		rowChildren.join('') +
		'</row>';
};

Export.prototype.getRowAttributes = function (row, rowIndex) {
	var attributes = ' r="' + (rowIndex + 1) + '"';

	if (row) {
		if (row.height !== undefined) {
			attributes += ' customHeight="1"';
			attributes += ' ht="' + row.height + '"';
		}
		if (row.style !== undefined) {
			attributes += ' customFormat="1"';
			attributes += ' s="' + row.style + '"';
		}
		if (row.outlineLevel) {
			attributes += ' outlineLevel="' + row.outlineLevel + '"';
		}
	}
	return attributes;
};

Export.prototype.exportDimension = function (maxX, maxY) {
	var attributes = [];

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
		attributes: attributes
	});
};

Export.prototype.exportColumns = function (columns) {
	if (columns.length) {
		var children = _.map(columns, function (column, index) {
			column = column || {};

			var attributes = [
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
			if (column.style !== undefined) {
				attributes.push(['style', column.style]);
			}

			return toXMLString({
				name: 'col',
				attributes: attributes
			});
		});

		return toXMLString({
			name: 'cols',
			children: children
		});
	}
	return '';
};

Export.prototype.getXMLBegin = function () {
	return util.xmlPrefix + '<worksheet xmlns="' + util.schemas.spreadsheetml +
		'" xmlns:r="' + util.schemas.relationships + '" xmlns:mc="' + util.schemas.markupCompat + '">';
};

Export.prototype.getXMLEnd = function () {
	return '</worksheet>';
};

module.exports = Export;
