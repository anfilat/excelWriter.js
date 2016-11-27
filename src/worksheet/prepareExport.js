'use strict';

var _ = require('lodash');

module.exports = {
	_prepare: function () {
		var rowIndex;
		var len;
		var maxX = 0;
		var row;

		this._prepareTables();

		this.preparedData = [];
		this.preparedRows = [];

		for (rowIndex = 0, len = this.data.length; rowIndex < len; rowIndex++) {
			row = prepareRow(this, rowIndex);

			if (row) {
				if (row.length > maxX) {
					maxX = row.length;
				}
			}
		}

		this.data = null;
		this.rows = null;

		this.maxX = maxX;
		this.maxY = this.preparedData.length;
	}
};

function prepareRow(worksheet, rowIndex) {
	var common = worksheet.common;
	var styles = common.styles;
	var row = worksheet.rows[rowIndex];
	var dataRow = worksheet.data[rowIndex];
	var preparedRow = [];
	var rowStyle = null;
	var column;
	var colIndex;
	var value;
	var cellValue;
	var cellStyle;
	var cellType;
	var cellFormula;
	var isString;
	var colSpan;
	var rowSpan;
	var date;

	if (dataRow) {
		if (dataRow.data) {
			row = row || {};
			row.height = dataRow.height || row.height;
			row.style = dataRow.style || row.style;
			row.outlineLevel = dataRow.outlineLevel || row.outlineLevel;

			dataRow = dataRow.data;
		}
		if (row) {
			rowStyle = row.style || null;
			row.style = styles.cells.getId(row.style);
		}

		for (colIndex = 0; colIndex < dataRow.length; colIndex++) {
			column = worksheet.columns[colIndex];
			value = dataRow[colIndex];

			if (_.isNil(value)) {
				continue;
			}

			cellFormula = null;
			isString = false;
			if (value && typeof value === 'object') {
				cellStyle = value.style || null;
				if (value.formula) {
					cellValue = value.formula;
					cellType = 'formula';
				} else if (value.date) {
					cellValue = value.date;
					cellType = 'date';
				} else if (value.time) {
					cellValue = value.time;
					cellType = 'time';
				} else {
					cellValue = value.value;
					cellType = value.type;
				}

				if (value.hyperlink) {
					worksheet._insertHyperlink(colIndex, rowIndex, value.hyperlink);
				}
				if (value.image) {
					worksheet._insertDrawing(colIndex, rowIndex, value.image);
				}
				if (value.colspan || value.rowspan) {
					colSpan = (value.colspan || 1) - 1;
					rowSpan = (value.rowspan || 1) - 1;

					worksheet.mergeCells({c: colIndex + 1, r: rowIndex + 1},
						{c: colIndex + 1 + colSpan, r: rowIndex + 1 + rowSpan});
					worksheet._insertMergeCells(dataRow, colIndex, rowIndex, colSpan, rowSpan);
				}
			} else {
				cellValue = value;
				cellStyle = null;
				cellType = null;
			}

			if (cellStyle === null) {
				if (rowStyle !== null) {
					cellStyle = rowStyle;
				} else if (column && column.style) {
					cellStyle = column.style;
				}
			}

			if (!cellType) {
				if (column && column.type) {
					cellType = column.type;
				} else if (typeof cellValue === 'number') {
					cellType = 'number';
				} else {
					cellType = 'string';
				}
			}

			if (cellType === 'string') {
				cellValue = common.addString(cellValue);
				isString = true;
			} else if (cellType === 'date' || cellType === 'time') {
				date = 25569.0 + (cellValue - worksheet.timezoneOffset) / (60 * 60 * 24 * 1000);
				if (_.isFinite(date)) {
					cellValue = date;
				} else {
					cellValue = common.addString(String(cellValue));
					isString = true;
				}
			} else if (cellType === 'formula') {
				cellFormula = _.escape(cellValue);
				cellValue = null;
			}

			preparedRow[colIndex] = {
				value: cellValue,
				formula: cellFormula,
				style: styles.cells.getId(cellStyle),
				isString: isString
			};
		}
	}

	worksheet.preparedData[rowIndex] = preparedRow;
	worksheet.preparedRows[rowIndex] = row;

	return preparedRow;
}
