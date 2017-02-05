'use strict';

const _ = require('lodash');
const util = require('../util');
const toXMLString = require('../XMLString');

module.exports = SuperClass => class MergedCells extends SuperClass {
	constructor(workbook, config) {
		super(workbook, config);
		this._mergedCells = [];
	}
	mergeCells(cell1, cell2) {
		this._mergedCells.push([cell1, cell2]);
		return this;
	}
	_insertMergeCells(dataRow, colIndex, rowIndex, colSpan, rowSpan) {
		let i, j;
		let row;

		if (colSpan) {
			for (j = 0; j < colSpan; j++) {
				dataRow.splice(colIndex + 1, 0, {style: null, type: 'empty'});
			}
		}
		if (rowSpan) {
			colSpan += 1;

			for (i = 0; i < rowSpan; i++) {
				//todo: original data changed
				row = this.data[rowIndex + i + 1];

				if (!row) {
					row = [];
					this.data[rowIndex + i + 1] = row;
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
	}
	_saveMergeCells() {
		if (this._mergedCells.length > 0) {
			const children = _.map(this._mergedCells, function (mergeCell) {
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
					['count', this._mergedCells.length]
				],
				children
			});
		}
		return '';
	}
};
