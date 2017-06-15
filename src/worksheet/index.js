'use strict';

const WorksheetSave = require('./save');
const Relations = require('../relations');

class Worksheet extends WorksheetSave {
	constructor(outerWorkbook, common, config = {}) {
		super(outerWorkbook, common, config);
		this.outerWorkbook = outerWorkbook;
		this.common = common;
		this.styles = this.common.styles;

		this.objectId = this.common.uniqueId('Worksheet');
		this.data = [];
		this.columns = [];
		this.rows = [];

		this.name = config.name;
		this.state = config.state || 'visible';
		this.timezoneOffset = new Date().getTimezoneOffset() * 60 * 1000;
		this.relations = new Relations(this.common);
	}
	end() {
		return this.outerWorkbook;
	}
	setActive() {
		this.common.setActiveWorksheet(this);
		return this;
	}
	setVisible() {
		this.setState('visible');
		return this;
	}
	setHidden() {
		this.setState('hidden');
		return this;
	}
	/**
	 * //http://www.datypic.com/sc/ooxml/t-ssml_ST_SheetState.html
	 * @param state - visible | hidden | veryHidden
	 */
	setState(state) {
		this.state = state;
		return this;
	}
	getState() {
		return this.state;
	}
	setRows(startRow, rows) {
		if (!rows) {
			rows = startRow;
			startRow = 0;
		} else {
			--startRow;
		}
		rows.forEach((row, i) => {
			this.rows[startRow + i] = row;
		});
		return this;
	}
	setRow(rowIndex, row) {
		this.rows[--rowIndex] = row;
		return this;
	}
	/**
	 * http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.column.aspx
	 */
	setColumns(startColumn, columns) {
		if (!columns) {
			columns = startColumn;
			startColumn = 0;
		} else {
			--startColumn;
		}
		columns.forEach((column, i) => {
			this.columns[startColumn + i] = column;
		});
		return this;
	}
	setColumn(columnIndex, column) {
		this.columns[--columnIndex] = column;
		return this;
	}
	setData(offset, data) {
		let startRow = this.data.length;

		if (!data) {
			data = offset;
		} else {
			startRow += offset;
		}
		data.forEach((row, i) => {
			this.data[startRow + i] = row;
		});
		return this;
	}
}

module.exports = Worksheet;
