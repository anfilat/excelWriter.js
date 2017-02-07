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
	_insertMergeCells(dataRow, colIndex, rowIndex, colSpan, rowSpan) {
		if (colSpan) {
			for (let j = 0; j < colSpan; j++) {
				dataRow.splice(colIndex + 1, 0, {style: null, type: 'empty'});
			}
		}
		if (rowSpan) {
			colSpan += 1;

			for (let i = 0; i < rowSpan; i++) {
				//todo: original data changed
				let row = this.data[rowIndex + i + 1];

				if (!row) {
					row = [];
					this.data[rowIndex + i + 1] = row;
				}

				if (row.length > colIndex) {
					for (let j = 0; j < colSpan; j++) {
						row.splice(colIndex, 0, {style: null, type: 'empty'});
					}
				} else {
					for (let j = 0; j < colSpan; j++) {
						row[colIndex + j] = {style: null, type: 'empty'};
					}
				}
			}
		}
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
