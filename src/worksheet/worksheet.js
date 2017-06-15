'use strict';

const Tables = require('./tables');
const WorksheetDrawings = require('./drawing');
const Hyperlinks = require('./hyperlinks');
const MergedCells = require('./mergedCells');
const SheetView = require('./sheetView');
const print = require('./print');
const prepareSave = require('./prepareSave');
const save = require('./save');
const Relations = require('../relations');

function Worksheet(common, config = {}) {
	this.common = common;
	this.styles = this.common.styles;

	this.objectId = this.common.uniqueId('Worksheet');
	this.data = [];
	this.columns = [];
	this.rows = [];

	this.headers = [];
	this.footers = [];

	this.name = config.name;
	this.state = config.state || 'visible';
	this.timezoneOffset = new Date().getTimezoneOffset() * 60 * 1000;
	this.relations = new Relations(this.common);

	this.tables = new Tables(this, this.common, this.relations);
	this.drawings = new WorksheetDrawings(this.common, this.relations);
	this.hyperlinks = new Hyperlinks(this.common, this.relations);
	this.mergedCells = new MergedCells(this);
	this.sheetView = new SheetView(config);
}

Worksheet.prototype = {
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
	},
	setRow(rowIndex, row) {
		this.rows[--rowIndex] = row;
	},
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
	},
	setColumn(columnIndex, column) {
		this.columns[--columnIndex] = column;
	},
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
	},
	setState(state) {
		this.state = state;
	},
	getState() {
		return this.state;
	}
};

Object.assign(Worksheet.prototype, print.methods);
Object.assign(Worksheet.prototype, prepareSave.methods);
Object.assign(Worksheet.prototype, save.methods);

module.exports = Worksheet;
