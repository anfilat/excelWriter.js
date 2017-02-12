'use strict';

const _ = require('lodash');
const SheetView = require('./sheetView');

class PrepareSave extends SheetView {
	_prepare() {
		this._prepareTables();
		this._prepareColumns();
		this._prepareRows();
		this._prepareData();
	}
	_prepareColumns() {
		this.preparedColumns = _.map(this.columns, column => {
			if (column) {
				const preparedColumn = _.clone(column);

				if (column.style) {
					const style = this.styles.addFormat(column.style);
					preparedColumn.style = style;

					const columnStyle = this.styles._addFillOutFormat(style);
					preparedColumn.styleId = this.styles._getId(columnStyle);
				}
				return preparedColumn;
			}
		});
		this.columns = null;
	}
	_prepareRows() {
		this.preparedRows = _.map(this.rows, row => {
			if (row) {
				const preparedRow = _.clone(row);

				if (row.style) {
					preparedRow.style = this.styles.addFormat(row.style);
				}
				return preparedRow;
			}
		});
		this.rows = null;
	}
	_prepareData() {
		this.maxX = 0;
		this.preparedData = [];
		for (let rowIndex = 0; rowIndex < this.data.length; rowIndex++) {
			const preparedDataRow = this._prepareDataRow(rowIndex);

			this.preparedData[rowIndex] = preparedDataRow;
			this.maxX = Math.max(this.maxX, preparedDataRow.length);
		}
		this.maxY = this.preparedData.length;
		this.data = null;
	}
	_prepareDataRow(rowIndex) {
		const strings = this.common.strings;
		const preparedDataRow = [];
		let row = this.preparedRows[rowIndex];
		let dataRow = this.data[rowIndex];

		if (dataRow) {
			let rowStyle = null;
			let skipColumnsStyle = false;
			let inserts = [];

			if (!_.isArray(dataRow)) {
				row = this._mergeDataRowToRow(row, dataRow);
				if (dataRow.inserts) {
					inserts = dataRow.inserts;
					dataRow = _.clone(dataRow.data);
				} else {
					dataRow = dataRow.data;
				}
			}
			if (row) {
				rowStyle = row.style || null;
				skipColumnsStyle = row.skipColumnsStyle;
			}
			dataRow = this._splitDataRow(row, dataRow, rowIndex);

			for (let colIndex = 0; colIndex < dataRow.length || colIndex < inserts.length; colIndex++) {
				if (inserts[colIndex]) {
					const insertCell = {style: inserts[colIndex].style, type: 'empty'};

					if (dataRow.length > colIndex) {
						dataRow.splice(colIndex, 0, insertCell);
					} else {
						dataRow[colIndex] = insertCell;
					}
				}

				const column = this.preparedColumns[colIndex];
				const columnStyle = !skipColumnsStyle && column ? column.style : null;
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
					} else if (_.isDate(value.value)) {
						cellValue = value.value;
						cellType = 'date';
					} else {
						cellValue = value.value;
						cellType = value.type;
					}

					this._insertEmbedded(value, colIndex, rowIndex);
					dataRow = this._mergeCells(dataRow, value, colIndex, rowIndex);
				} else {
					cellValue = value;
					cellType = null;
				}

				cellStyle = this.styles._merge(columnStyle, rowStyle, cellStyle);

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
					const date = 25569.0 + (dateValue - this.timezoneOffset) / (60 * 60 * 24 * 1000);

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
					styleId: this.styles._getId(cellStyle),
					isString
				};
			}
		}

		if (row) {
			this._setRowStyleId(row);
			this.preparedRows[rowIndex] = row;
		}

		return preparedDataRow;
	}
	_mergeDataRowToRow(row = {}, dataRow) {
		row.height = dataRow.height || row.height;
		row.outlineLevel = dataRow.outlineLevel || row.outlineLevel;
		row.type = dataRow.type || row.type;
		row.style = dataRow.style ? this.styles.addFormat(dataRow.style) : row.style;
		row.skipColumnsStyle = dataRow.skipColumnsStyle || row.skipColumnsStyle;

		return row;
	}
	_splitDataRow(row = {}, dataRow, rowIndex) {
		const count = this._dataRowHeight(dataRow);
		if (count === 1) {
			return dataRow;
		}

		const newRows = _.times(count, () => {
			const result = _.clone(row);
			result.data = [];
			return result;
		});
		_.forEach(dataRow, value => {
			if (_.isArray(value)) {
				_(value)
					.initial()
					.forEach((val, index) => {
						newRows[index].data.push(val);
					});

				const val = value.length < count
					? this._addRowspan(_.last(value), count - value.length + 1)
					: _.last(value);
				newRows[value.length - 1].data.push(val);
			} else {
				newRows[0].data.push(this._addRowspan(value, count));
			}
		});
		this.data.splice(rowIndex, 1, ...newRows);

		return newRows[0].data;
	}
	_dataRowHeight(dataRow) {
		let count = 1;
		_.forEach(dataRow, value => {
			if (_.isArray(value)) {
				count = Math.max(value.length, count);
			}
		});
		return count;
	}
	_addRowspan(value, rowspan) {
		if (_.isObject(value) && !_.isDate(value)) {
			value = _.clone(value);
			value.rowspan = rowspan;
			value.style = this.styles._merge({vertical: 'top'}, value.style);
			return value;
		}
		return {
			value,
			rowspan,
			style: {vertical: 'top'}
		};
	}
	_setRowStyleId(row) {
		if (row.style) {
			const rowStyle = this.styles._addFillOutFormat(row.style);
			row.styleId = this.styles._getId(rowStyle);
		}
	}
	_insertEmbedded(value, colIndex, rowIndex) {
		if (value.hyperlink) {
			this._insertHyperlink(colIndex, rowIndex, value.hyperlink);
		}

		if (value.image) {
			this._insertDrawing(colIndex, rowIndex, value.image);
		}
	}
	_mergeCells(dataRow, value, colIndex, rowIndex) {
		if (value.colspan || value.rowspan) {
			const colSpan = (value.colspan || 1) - 1;
			const rowSpan = (value.rowspan || 1) - 1;

			if (colSpan || rowSpan) {
				this.mergeCells(
					{c: colIndex + 1, r: rowIndex + 1},
					{c: colIndex + 1 + colSpan, r: rowIndex + 1 + rowSpan}
				);
				return this._insertMergeCells(dataRow, colIndex, rowIndex, colSpan, rowSpan, value.style);
			}
		}
		return dataRow;
	}
}

module.exports = PrepareSave;
