'use strict';

var _ = require('lodash');

module.exports = {
	_prepare: function () {
		var rowIndex;
		var len;
		var maxX = 0;
		var preparedDataRow;

		this.preparedData = [];
		this.preparedColumns = [];
		this.preparedRows = [];

		this._prepareTables();
		prepareColumns(this);
		prepareRows(this);

		for (rowIndex = 0, len = this.data.length; rowIndex < len; rowIndex++) {
			preparedDataRow = prepareDataRow(this, rowIndex);

			if (preparedDataRow.length > maxX) {
				maxX = preparedDataRow.length;
			}
		}

		this.data = null;
		this.columns = null;
		this.rows = null;

		this.maxX = maxX;
		this.maxY = this.preparedData.length;
	}
};

function prepareColumns(worksheet) {
	var styles = worksheet.common.styles;

	_.forEach(worksheet.columns, function (column, index) {
		var preparedColumn;

		if (column) {
			preparedColumn = _.clone(column);

			if (column.style) {
				preparedColumn.style = styles.addFormat(column.style);
				preparedColumn.styleId = styles._getId(preparedColumn.style);
			}
			worksheet.preparedColumns[index] = preparedColumn;
		}
	});
}

function prepareRows(worksheet) {
	var styles = worksheet.common.styles;

	_.forEach(worksheet.rows, function (row, index) {
		var preparedRow;

		if (row) {
			preparedRow = _.clone(row);

			if (row.style) {
				preparedRow.style = styles.addFormat(row.style);
			}
			worksheet.preparedRows[index] = preparedRow;
		}
	});
}

function prepareDataRow(worksheet, rowIndex) {
	var common = worksheet.common;
	var styles = common.styles;
	var row = worksheet.preparedRows[rowIndex];
	var dataRow = worksheet.data[rowIndex];
	var preparedDataRow = [];
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
		if (!_.isArray(dataRow)) {
			row = mergeDataRowToRow(row, dataRow);
			dataRow = dataRow.data;
		}
		if (row) {
			rowStyle = row.style || null;
		}

		for (colIndex = 0; colIndex < dataRow.length; colIndex++) {
			column = worksheet.preparedColumns[colIndex];
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

			preparedDataRow[colIndex] = {
				value: cellValue,
				formula: cellFormula,
				styleId: cellStyle ? styles._getId(styles.addFormat(cellStyle)) : null,
				isString: isString
			};
		}
	}

	worksheet.preparedData[rowIndex] = preparedDataRow;
	if (row) {
		row.styleId = styles._getId(row.style);
		worksheet.preparedRows[rowIndex] = row;
	}

	return preparedDataRow;
}

function mergeDataRowToRow(row, dataRow) {
	row = row || {};
	row.height = dataRow.height || row.height;
	row.style = dataRow.style || row.style;
	row.outlineLevel = dataRow.outlineLevel || row.outlineLevel;

	return row;
}
