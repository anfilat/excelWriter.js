'use strict';

const _ = require('lodash');
const SheetView = require('./sheetView');

class PrepareSave extends SheetView {
	_prepare() {
		let maxX = 0;

		this.preparedData = [];
		this.preparedColumns = [];
		this.preparedRows = [];

		this._prepareTables();
		prepareColumns(this);
		prepareRows(this);

		for (let rowIndex = 0, len = this.data.length; rowIndex < len; rowIndex++) {
			const preparedDataRow = prepareDataRow(this, rowIndex);

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
		if (column) {
			const preparedColumn = _.clone(column);

			if (column.style) {
				const style = styles.addFormat(column.style);
				const columnStyle = styles._get(style).fillOut ? style : styles._addInvisibleFormat(style);

				preparedColumn.style = style;
				preparedColumn.styleId = styles._getId(columnStyle);
			}
			worksheet.preparedColumns[index] = preparedColumn;
		}
	});
}

function prepareRows(worksheet) {
	const styles = worksheet.common.styles;

	_.forEach(worksheet.rows, (row, index) => {
		if (row) {
			const preparedRow = _.clone(row);

			if (row.style) {
				preparedRow.style = styles.addFormat(row.style);
			}
			worksheet.preparedRows[index] = preparedRow;
		}
	});
}

function prepareDataRow(worksheet, rowIndex) {
	const styles = worksheet.common.styles;
	const strings = worksheet.common.strings;
	const preparedDataRow = [];
	let row = worksheet.preparedRows[rowIndex];
	let dataRow = worksheet.data[rowIndex];

	if (dataRow) {
		let rowStyle = null;

		if (!_.isArray(dataRow)) {
			row = mergeDataRowToRow(styles, row, dataRow);
			dataRow = dataRow.data;
		}
		if (row) {
			rowStyle = row.style || null;
		}

		for (let colIndex = 0; colIndex < dataRow.length; colIndex++) {
			const column = worksheet.preparedColumns[colIndex];
			const value = dataRow[colIndex];
			let cellValue;
			let cellType;
			let cellStyle = null;
			let cellFormula = null;
			let isString = false;

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
				if (row && row.type) {
					cellType = row.type;
				} else if (column && column.type) {
					cellType = column.type;
				} else if (typeof cellValue === 'number') {
					cellType = 'number';
				} else if (typeof cellValue === 'string') {
					cellType = 'string';
				}
			}

			if (cellType === 'string') {
				cellValue = strings.add(cellValue);
				isString = true;
			} else if (cellType === 'date' || cellType === 'time') {
				const dateValue = _.isDate(cellValue) ? cellValue.valueOf() : cellValue;
				const date = 25569.0 + (dateValue - worksheet.timezoneOffset) / (60 * 60 * 24 * 1000);

				if (_.isFinite(date)) {
					cellValue = date;
				} else {
					cellValue = strings.add(String(cellValue));
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
		setRowStyleId(styles, row);
		worksheet.preparedRows[rowIndex] = row;
	}

	return preparedDataRow;
}

function mergeDataRowToRow(styles, row = {}, dataRow) {
	row.height = dataRow.height || row.height;
	row.outlineLevel = dataRow.outlineLevel || row.outlineLevel;
	row.type = dataRow.type || row.type;
	row.style = dataRow.style ? styles.addFormat(dataRow.style) : row.style;

	return row;
}

function setRowStyleId(styles, row) {
	if (row.style) {
		const rowStyle = styles._get(row.style).fillOut ? row.style : styles._addInvisibleFormat(row.style);

		row.styleId = styles._getId(rowStyle);
	}
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

		if (colSpan || rowSpan) {
			worksheet.mergeCells(
				{c: colIndex + 1, r: rowIndex + 1},
				{c: colIndex + 1 + colSpan, r: rowIndex + 1 + rowSpan}
			);
			worksheet._insertMergeCells(dataRow, colIndex, rowIndex, colSpan, rowSpan);
		}
	}
}

module.exports = PrepareSave;
