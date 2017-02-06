'use strict';

const _ = require('lodash');
const SheetView = require('./sheetView');

class PrepareSave extends SheetView {
	_prepare() {
		let maxX = 0;
		let preparedDataRow;

		this.preparedData = [];
		this.preparedColumns = [];
		this.preparedRows = [];

		this._prepareTables();
		prepareColumns(this);
		prepareRows(this);

		for (let rowIndex = 0, len = this.data.length; rowIndex < len; rowIndex++) {
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
}

function prepareColumns(worksheet) {
	const styles = worksheet.common.styles;

	_.forEach(worksheet.columns, (column, index) => {
		let preparedColumn;

		if (column) {
			preparedColumn = _.clone(column);

			if (column.style) {
				preparedColumn.style = styles.addFormat(column.style);
				if (styles._get(preparedColumn.style).fillOut) {
					preparedColumn.styleId = styles._getId(preparedColumn.style);
				} else {
					preparedColumn.styleId = styles._getId(styles._addInvisibleFormat(preparedColumn.style));
				}
			}
			worksheet.preparedColumns[index] = preparedColumn;
		}
	});
}

function prepareRows(worksheet) {
	const styles = worksheet.common.styles;

	_.forEach(worksheet.rows, (row, index) => {
		let preparedRow;

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
	const common = worksheet.common;
	const styles = common.styles;
	const preparedDataRow = [];
	let row = worksheet.preparedRows[rowIndex];
	let dataRow = worksheet.data[rowIndex];
	let rowStyle = null;

	if (dataRow) {
		if (!_.isArray(dataRow)) {
			row = mergeDataRowToRow(worksheet, row, dataRow);
			dataRow = dataRow.data;
		}
		if (row) {
			rowStyle = row.style || null;
		}

		for (let colIndex = 0; colIndex < dataRow.length; colIndex++) {
			const column = worksheet.preparedColumns[colIndex];
			const value = dataRow[colIndex];
			let cellStyle = null;
			let cellFormula = null;
			let isString = false;
			let cellValue;
			let cellType;

			if (_.isDate(value)) {
				cellValue = value;
				cellType = 'date';
			} else if (value && typeof value === 'object') {
				if (value.style) {
					cellStyle = value.style;
				}
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

				insertEmbedded(worksheet, dataRow, value, colIndex, rowIndex);
			} else {
				cellValue = value;
				cellType = null;
			}

			cellStyle = styles._merge(column ? column.style : null, rowStyle, cellStyle);

			if (!cellType) {
				if (column && column.type) {
					cellType = column.type;
				} else if (typeof cellValue === 'number') {
					cellType = 'number';
				} else if (typeof cellValue === 'string') {
					cellType = 'string';
				}
			}

			if (cellType === 'string') {
				cellValue = common.addString(cellValue);
				isString = true;
			} else if (cellType === 'date' || cellType === 'time') {
				const date = 25569.0 + ((_.isDate(cellValue) ? cellValue.valueOf() : cellValue) - worksheet.timezoneOffset) /
					(60 * 60 * 24 * 1000);
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
				styleId: styles._getId(cellStyle),
				isString
			};
		}
	}

	worksheet.preparedData[rowIndex] = preparedDataRow;
	if (row) {
		if (row.style) {
			if (styles._get(row.style).fillOut) {
				row.styleId = styles._getId(row.style);
			} else {
				row.styleId = styles._getId(styles._addInvisibleFormat(row.style));
			}
		}
		worksheet.preparedRows[rowIndex] = row;
	}

	return preparedDataRow;
}

function mergeDataRowToRow(worksheet, row, dataRow) {
	row = row || {};
	row.height = dataRow.height || row.height;
	row.style = dataRow.style ? worksheet.common.styles.addFormat(dataRow.style) : row.style;
	row.outlineLevel = dataRow.outlineLevel || row.outlineLevel;

	return row;
}

function insertEmbedded(worksheet, dataRow, value, colIndex, rowIndex) {
	if (value.hyperlink) {
		worksheet._insertHyperlink(colIndex, rowIndex, value.hyperlink);
	}
	if (value.image) {
		worksheet._insertDrawing(colIndex, rowIndex, value.image);
	}
	if (value.colspan || value.rowspan) {
		const colSpan = (value.colspan || 1) - 1;
		const rowSpan = (value.rowspan || 1) - 1;

		worksheet.mergeCells({c: colIndex + 1, r: rowIndex + 1},
			{c: colIndex + 1 + colSpan, r: rowIndex + 1 + rowSpan});
		worksheet._insertMergeCells(dataRow, colIndex, rowIndex, colSpan, rowSpan);
	}
}

module.exports = PrepareSave;
