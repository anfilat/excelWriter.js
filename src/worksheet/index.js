'use strict';

const _ = require('lodash');
const sheetView = require('./sheetView');
const hyperlinks = require('./hyperlinks');
const mergedCells = require('./mergedCells');
const print = require('./print');
const drawing = require('./drawing');
const tables = require('./tables');
const prepareExport = require('./prepareExport');
const worksheetExport = require('./export');
const Relations = require('../relations');

class Worksheet extends sheetView(hyperlinks(mergedCells(drawing(tables(print(worksheetExport(prepareExport(Object)))))))) {
	constructor(workbook, config = {}) {
		super(workbook, config);
		this.workbook = workbook;
		this.common = workbook.common;

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
		return this.workbook;
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
		_.forEach(rows, (row, i) => {
			this.rows[startRow + i] = row;
		});
		return this;
	}
	setRow(rowIndex, meta) {
		this.rows[--rowIndex] = meta;
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
		_.forEach(columns, (column, i) => {
			this.columns[startColumn + i] = column;
		});
		return this;
	}
	setColumn(columnIndex, column) {
		this.columns[--columnIndex] = column;
		return this;
	}
	setData(startRow, data) {
		if (!data) {
			data = startRow;
			startRow = 0;
		} else {
			--startRow;
		}
		_.forEach(data, (row, i) => {
			this.data[startRow + i] = row;
		});
		return this;
	}
}

module.exports = Worksheet;
