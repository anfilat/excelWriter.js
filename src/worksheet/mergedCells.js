'use strict';

var _ = require('lodash');
var util = require('../util');
var toXMLString = require('../XMLString');

function MergedCells(worksheet) {
	this.worksheet = worksheet;
	this.relations = worksheet.relations;

	this.mergedCells = [];
}

/**
 * Merge cells in given range
 *
 * @param cell1 - A1, A2...
 * @param cell2 - A2, A3...
 */
MergedCells.prototype.merge = function (cell1, cell2) {
	this.mergedCells.push([cell1, cell2]);
};

MergedCells.prototype.insert = function (dataRow, colIndex, rowIndex, colSpan, rowSpan) {
	var i, j;
	var row;

	if (colSpan) {
		for (j = 0; j < colSpan; j++) {
			dataRow.splice(colIndex + 1, 0, {style: null, type: 'empty'});
		}
	}
	if (rowSpan) {
		colSpan += 1;

		for (i = 0; i < rowSpan; i++) {
			//todo: original data changed
			row = this.worksheet.data[rowIndex + i + 1];

			if (!row) {
				row = [];
				this.worksheet.data[rowIndex + i + 1] = row;
			}

			if (row.length > colIndex) {
				for (j = 0; j < colSpan; j++) {
					row.splice(colIndex, 0, {style: null, type: 'empty'});
				}
			} else {
				for (j = 0; j < colSpan; j++) {
					row[colIndex + j] = {style: null, type: 'empty'};
				}
			}
		}
	}
};

MergedCells.prototype.export = function () {
	if (this.mergedCells.length > 0) {
		var children = _.map(this.mergedCells, function (mergeCell) {
			return toXMLString({
				name: 'mergeCell',
				attributes: [
					['ref', util.canonCell(mergeCell[0]) + ':' + util.canonCell(mergeCell[1])]
				]
			});
		});

		return toXMLString({
			name: 'mergeCells',
			attributes: [
				['count', this.mergedCells.length]
			],
			children: children
		});
	}
	return '';
};

module.exports = MergedCells;
