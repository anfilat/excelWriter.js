'use strict';

var _ = require('lodash');
var util = require('./util');
var toXMLString = require('./XMLString');

function Table(worksheet, config) {
	this.worksheet = worksheet;
	this.common = worksheet.common;
	this.tableId = util.uniqueId('Table');
	this.objectId = 'Table' + this.tableId;
	this.name = this.objectId;
	this.displayName = this.objectId;
	this.headerRowBorderDxfId = null;
	this.headerRowCount = 1;
	this.headerRowDxfId = null;
	this.beginCell = null;
	this.endCell = null;
	this.totalsRowCount = 0;
	this.totalRow = null;
	this.themeStyle = null;

	_.extend(this, config);
}

Table.prototype.end = function () {
	return this.worksheet;
};

Table.prototype.setReferenceRange = function (beginCell, endCell) {
	this.beginCell = util.canonCell(beginCell);
	this.endCell = util.canonCell(endCell);
	return this;
};

Table.prototype.addTotalRow = function (totalRow) {
	this.totalRow = totalRow;
	this.totalsRowCount = 1;
	return this;
};

Table.prototype.setTheme = function (theme) {
	this.themeStyle = theme;
	return this;
};

/**
 * Expects an object with the following properties:
 * caseSensitive (boolean)
 * dataRange
 * columnSort (assumes true)
 * sortDirection
 * sortRange (defaults to dataRange)
 */
Table.prototype.setSortState = function (state) {
	this.sortState = state;
	return this;
};

Table.prototype._prepare = function (worksheetData) {
	var SUB_TOTAL_FUNCTIONS = ['average', 'countNums', 'count', 'max', 'min', 'stdDev', 'sum', 'var'];
	var SUB_TOTAL_NUMS = [101, 102, 103, 104, 105, 107, 109, 110];

	if (this.totalRow) {
		var tableName = this.name;
		var beginCell = util.letterToPosition(this.beginCell);
		var endCell = util.letterToPosition(this.endCell);
		var firstRow = beginCell.y - 1;
		var firstColumn = beginCell.x - 1;
		var lastRow = endCell.y - 1;
		var headerRow = worksheetData[firstRow] || [];
		var totalRow = worksheetData[lastRow + 1];

		if (!totalRow) {
			totalRow = [];
			worksheetData[lastRow + 1] = totalRow;
		}

		_.forEach(this.totalRow, function (cell, i) {
			var headerValue = headerRow[firstColumn + i];
			var funcIndex;

			if (typeof headerValue === 'object') {
				headerValue = headerValue.value;
			}
			cell.name = headerValue;
			if (cell.totalsRowLabel) {
				totalRow[firstColumn + i] = {
					value: cell.totalsRowLabel,
					type: 'string'
				};
			} else if (cell.totalsRowFunction) {
				funcIndex = _.indexOf(SUB_TOTAL_FUNCTIONS, cell.totalsRowFunction);

				if (funcIndex !== -1) {
					totalRow[firstColumn + i] = {
						value: 'SUBTOTAL(' + SUB_TOTAL_NUMS[funcIndex] + ',' + tableName + '[' + headerValue + '])',
						type: 'formula'
					};
				}
			}
		});
	}
};

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.table.aspx
Table.prototype._export = function () {
	var attributes = [
		['id', this.tableId],
		['name', this.name],
		['displayName', this.displayName]
	];
	var children = [];
	var end = util.letterToPosition(this.endCell);
	var ref = this.beginCell + ':' + util.positionToLetter(end.x, end.y + this.totalsRowCount);

	attributes.push(['ref', ref]);
	attributes.push(['totalsRowCount', this.totalsRowCount]);
	attributes.push(['headerRowCount', this.headerRowCount]);

	children.push(exportAutoFilter(this.beginCell, this.endCell));
	children.push(exportTableColumns(this.totalRow));
	children.push(exportTableStyleInfo(this.common, this.themeStyle));

	return toXMLString({
		name: 'table',
		ns: 'spreadsheetml',
		attributes: attributes,
		children: children
	});
};

function exportAutoFilter(beginCell, endCell) {
	return toXMLString({
		name: 'autoFilter',
		attributes: ['ref', beginCell + ':' + endCell]
	});
}

function exportTableColumns(totalRow) {
	var attributes = [
		['count', totalRow.length]
	];
	var children = _.map(totalRow, function (cell, index) {
		var attributes = [
			['id', index + 1],
			['name', cell.name]
		];

		if (cell.totalsRowFunction) {
			attributes.push(['totalsRowFunction', cell.totalsRowFunction]);
		}
		if (cell.totalsRowLabel) {
			attributes.push(['totalsRowLabel', cell.totalsRowLabel]);
		}

		return toXMLString({
			name: 'tableColumn',
			attributes: attributes
		});
	});

	return toXMLString({
		name: 'tableColumns',
		attributes: attributes,
		children: children
	});
}

function exportTableStyleInfo(common, themeStyle) {
	var attributes = [
		['name', themeStyle]
	];
	var isRowStripes = false;
	var isColumnStripes = false;
	var isFirstColumn = false;
	var isLastColumn = false;
	var format = common.styles.tables.get(themeStyle);

	if (format) {
		isRowStripes = format.firstRowStripe || format.secondRowStripe;
		isColumnStripes = format.firstColumnStripe || format.secondColumnStripe;
		isFirstColumn = format.firstColumn;
		isLastColumn = format.lastColumn;
	}
	attributes.push(
		['showRowStripes', isRowStripes ? '1' : '0'],
		['showColumnStripes', isColumnStripes ? '1' : '0'],
		['showFirstColumn', isFirstColumn ? '1' : '0'],
		['showLastColumn', isLastColumn ? '1' : '0']
	);

	return toXMLString({
		name: 'tableStyleInfo',
		attributes: attributes
	});
}

module.exports = Table;
