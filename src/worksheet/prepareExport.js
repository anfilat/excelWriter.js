'use strict';

var _ = require('lodash');

function PrepareExport(worksheet) {
	this.worksheet = worksheet;
}

PrepareExport.prototype.prepare = function () {
	var worksheet = this.worksheet;
	var rowIndex;
	var len;
	var maxX = 0;
	var row;

	worksheet.tables._prepare();

	worksheet.preparedData = [];
	worksheet.preparedRows = [];

	for (rowIndex = 0, len = worksheet.data.length; rowIndex < len; rowIndex++) {
		row = this.prepareRow(rowIndex);

		if (row) {
			if (row.length > maxX) {
				maxX = row.length;
			}
		}
	}

	worksheet.data = null;
	worksheet.rows = null;

	worksheet.maxX = maxX;
	worksheet.maxY = worksheet.preparedData.length;
};

PrepareExport.prototype.prepareRow = function (rowIndex) {
	var worksheet = this.worksheet;
	var styles = worksheet.common.styles;
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
					worksheet.hyperlinks.insert(colIndex, rowIndex, value.hyperlink);
				}
				if (value.image) {
					worksheet.drawing.insert(colIndex, rowIndex, value.image);
				}
				if (value.colspan || value.rowspan) {
					colSpan = (value.colspan || 1) - 1;
					rowSpan = (value.rowspan || 1) - 1;

					worksheet.mergedCells.merge({c: colIndex + 1, r: rowIndex + 1},
						{c: colIndex + 1 + colSpan, r: rowIndex + 1 + rowSpan});
					worksheet.mergedCells.insert(dataRow, colIndex, rowIndex, colSpan, rowSpan);
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
				cellValue = worksheet.common.sharedStrings.add(cellValue);
				isString = true;
			} else if (cellType === 'date' || cellType === 'time') {
				date = 25569.0 + (cellValue - worksheet.timezoneOffset) / (60 * 60 * 24 * 1000);
				if (_.isFinite(date)) {
					cellValue = date;
				} else {
					cellValue = worksheet.common.sharedStrings.add(String(cellValue));
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
};

module.exports = PrepareExport;
