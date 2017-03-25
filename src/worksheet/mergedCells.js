'use strict';

const _ = require('lodash');
const util = require('../util');
const DrawingsExt = require('./drawing');
const toXMLString = require('../XMLString');

class MergedCells extends DrawingsExt {
	constructor() {
		super();
		this._mergedCells = [];
	}
	mergeCells(cell1, cell2) {
		this._mergedCells.push([cell1, cell2]);
		return this;
	}
	_insertMergeCells(dataRow, colIndex, rowIndex, colSpan, rowSpan, style) {
		if (colSpan) {
			dataRow = [
				...dataRow.slice(0, colIndex + 1),
				..._.times(colSpan, () => ({style, type: 'empty'})),
				...dataRow.slice(colIndex + 1)
			];
		}
		if (rowSpan) {
			_.forEach(_.range(rowIndex + 1, rowIndex + 1 + rowSpan), index => {
				let row = this.data[index] || [];

				if (_.isArray(row)) {
					row = {
						data: row
					};
					this.data[index] = row;
				}

				row.inserts = row.inserts || [];
				_.forEach(_.range(colIndex, colIndex + colSpan + 1), index => {
					row.inserts[index] = {style};
				});
			});
		}
		return dataRow;
	}
	_saveMergeCells() {
		if (this._mergedCells.length > 0) {
			const children = this._mergedCells.map(
				mergeCell => toXMLString({
					name: 'mergeCell',
					attributes: [
						['ref', util.canonCell(mergeCell[0]) + ':' + util.canonCell(mergeCell[1])]
					]
				})
			);

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
}

module.exports = MergedCells;
