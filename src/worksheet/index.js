'use strict';

var _ = require('lodash');
var sheetView = require('./sheetView');
var hyperlinks = require('./hyperlinks');
var mergedCells = require('./mergedCells');
var print = require('./print');
var drawing = require('./drawing');
var tables = require('./tables');
var prepareExport = require('./prepareExport');
var worksheetExport = require('./export');
var RelationshipManager = require('../relationshipManager');

function Worksheet(workbook, config) {
	config = config || {};

	this.workbook = workbook;
	this.common = workbook.common;

	sheetView.init.call(this, config);
	hyperlinks.init.call(this);
	mergedCells.init.call(this);
	drawing.init.call(this);
	tables.init.call(this);
	print.init.call(this);

	this.objectId = this.common.uniqueId('Worksheet');
	this.data = [];
	this.columns = [];
	this.rows = [];

	this.name = config.name;
	this.state = config.state || 'visible';
	this.timezoneOffset = new Date().getTimezoneOffset() * 60 * 1000;
	this.relations = new RelationshipManager(this.common);
}

_.assign(Worksheet.prototype, prepareExport);
_.assign(Worksheet.prototype, worksheetExport);
_.assign(Worksheet.prototype, sheetView.methods);
_.assign(Worksheet.prototype, hyperlinks.methods);
_.assign(Worksheet.prototype, mergedCells.methods);
_.assign(Worksheet.prototype, drawing.methods);
_.assign(Worksheet.prototype, tables.methods);
_.assign(Worksheet.prototype, print.methods);

Worksheet.prototype.end = function () {
	return this.workbook;
};

Worksheet.prototype.setActive = function () {
	this.common.setActiveWorksheet(this);
	return this;
};

Worksheet.prototype.setVisible = function () {
	this.setState('visible');
	return this;
};

Worksheet.prototype.setHidden = function () {
	this.setState('hidden');
	return this;
};

/**
 * //http://www.datypic.com/sc/ooxml/t-ssml_ST_SheetState.html
 * @param state - visible | hidden | veryHidden
 */
Worksheet.prototype.setState = function (state) {
	this.state = state;
	return this;
};

Worksheet.prototype.getState = function () {
	return this.state;
};

Worksheet.prototype.setRows = function (startRow, rows) {
	var self = this;

	if (!rows) {
		rows = startRow;
		startRow = 0;
	} else {
		--startRow;
	}
	_.forEach(rows, function (row, i) {
		self.rows[startRow + i] = row;
	});
	return this;
};

Worksheet.prototype.setRow = function (rowIndex, meta) {
	this.rows[--rowIndex] = meta;
	return this;
};

/**
 * http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.column.aspx
 */
Worksheet.prototype.setColumns = function (startColumn, columns) {
	var self = this;

	if (!columns) {
		columns = startColumn;
		startColumn = 0;
	} else {
		--startColumn;
	}
	_.forEach(columns, function (column, i) {
		self.columns[startColumn + i] = column;
	});
	return this;
};

Worksheet.prototype.setColumn = function (columnIndex, column) {
	this.columns[--columnIndex] = column;
	return this;
};

Worksheet.prototype.setData = function (startRow, data) {
	var self = this;

	if (!data) {
		data = startRow;
		startRow = 0;
	} else {
		--startRow;
	}
	_.forEach(data, function (row, i) {
		self.data[startRow + i] = row;
	});
	return this;
};

module.exports = Worksheet;
