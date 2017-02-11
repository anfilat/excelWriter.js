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
			const cells = _.times(colSpan, () => ({style: style, type: 'empty'}));
			dataRow = [...dataRow.slice(0, colIndex + 1), ...cells, ...dataRow.slice(colIndex + 1)];
		}
		if (rowSpan) {
			colSpan += 1;

			for (let i = 0; i < rowSpan; i++) {
				let rowData = this.data[rowIndex + i + 1] || [];
				let row;

				if (_.isArray(rowData)) {
					row = {
						data: rowData
					};
					this.data[rowIndex + i + 1] = row;
				} else {
					row = rowData;
					rowData = rowData.data;
				}

				row.inserts = row.inserts || [];
				_.times(colSpan, index => {
					row.inserts[colIndex + index] = {style};
				});
			}
		}
		return dataRow;
	}
	_saveMergeCells() {
		if (this._mergedCells.length > 0) {
			const children = _.map(this._mergedCells,
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
