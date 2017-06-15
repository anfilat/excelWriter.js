'use strict';

const _ = require('lodash');
const util = require('../util');
const toXMLString = require('../XMLString');

function MergedCells(worksheet) {
	this.worksheet = worksheet;
	this.mergedCells = [];
}

MergedCells.prototype = {
	mergeCells(cell1, cell2) {
		this.mergedCells.push([cell1, cell2]);
	},
	insert(dataRow, colIndex, rowIndex, colSpan, rowSpan, style) {
		if (colSpan) {
			dataRow = [
				...dataRow.slice(0, colIndex + 1),
				..._.times(colSpan, () => ({style, type: 'empty'})),
				...dataRow.slice(colIndex + 1)
			];
		}
		if (rowSpan) {
			const data = this.worksheet.data;

			_.forEach(_.range(rowIndex + 1, rowIndex + 1 + rowSpan), index => {
				let row = data[index] || [];

				if (_.isArray(row)) {
					row = {
						data: row
					};
					data[index] = row;
				}

				row.inserts = row.inserts || [];
				_.forEach(_.range(colIndex, colIndex + colSpan + 1), index => {
					row.inserts[index] = {style};
				});
			});
		}
		return dataRow;
	},
	save() {
		if (this.mergedCells.length > 0) {
			const children = this.mergedCells.map(
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
					['count', this.mergedCells.length]
				],
				children
			});
		}
		return '';
	}
};

module.exports = MergedCells;
