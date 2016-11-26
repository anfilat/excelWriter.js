'use strict';

var _ = require('lodash');
var util = require('../util');
var toXMLString = require('../XMLString');

module.exports = {
	init: function () {
		this._mergedCells = [];
	},
	methods: {
		mergeCells: function (cell1, cell2) {
			this._mergedCells.push([cell1, cell2]);
			return this;
		},
		_insertMergeCells: function (dataRow, colIndex, rowIndex, colSpan, rowSpan) {
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
		},
		_exportMergeCells: function () {
			if (this._mergedCells.length > 0) {
				var children = _.map(this._mergedCells, function (mergeCell) {
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
					children: children
				});
			}
			return '';
		}
	}
};
