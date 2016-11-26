'use strict';

var _ = require('lodash');
var SheetView = require('./sheetView');
var Hyperlinks = require('./hyperlinks');
var MergedCells = require('./mergedCells');
var Print = require('./print');
var Drawing = require('./drawing');
var Tables = require('./tables');
var PrepareExport = require('./prepareExport');
var Export = require('./export');
var RelationshipManager = require('../relationshipManager');

function Worksheet(workbook, config) {
	config = config || {};

	this.workbook = workbook;
	this.common = workbook.common;

	this.objectId = this.common.uniqueId('Worksheet');
	this.data = [];
	this.columns = [];
	this.rows = [];

	this.name = config.name;
	this.state = config.state || 'visible';
	this.timezoneOffset = new Date().getTimezoneOffset() * 60 * 1000;

	this.relations = new RelationshipManager(this.common);
	this.sheetView = new SheetView(config);
	this.hyperlinks = new Hyperlinks(this);
	this.mergedCells = new MergedCells(this);
	this.print = new Print(this);
	this.drawing = new Drawing(this);
	this.tables = new Tables(this);
}

Worksheet.prototype.end = function () {
	return this.workbook;
};

Worksheet.prototype.addTable = function (config) {
	return this.tables.add(config);
};

Worksheet.prototype.setImage = function (image, config) {
	this.drawing.set(image, config, 'anchor');
	return this;
};

Worksheet.prototype.setImageOneCell = function (image, config) {
	this.drawing.set(image, config, 'oneCell');
	return this;
};

Worksheet.prototype.setImageAbsolute = function (image, config) {
	this.drawing.set(image, config, 'absolute');
	return this;
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

Worksheet.prototype.setHeader = function (headers) {
	this.print.setHeader(headers);
	return this;
};

Worksheet.prototype.setFooter = function (footers) {
	this.print.setFooter(footers);
	return this;
};

Worksheet.prototype.setPageMargin = function (margin) {
	this.print.setPageMargin(margin);
	return this;
};

Worksheet.prototype.setPageOrientation = function (orientation) {
	this.print.setPageOrientation(orientation);
	return this;
};

Worksheet.prototype.setPrintTitleTop = function (params) {
	this.print.setPrintTitleTop(params);
	return this;
};

Worksheet.prototype.setPrintTitleLeft = function (params) {
	this.print.setPrintTitleLeft(params);
	return this;
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

Worksheet.prototype.mergeCells = function (cell1, cell2) {
	this.mergedCells.merge(cell1, cell2);
	return this;
};

Worksheet.prototype.setHyperlink = function (hyperlink) {
	this.hyperlinks.set(hyperlink);
	return this;
};

Worksheet.prototype.setAttribute = function (name, value) {
	this.sheetView.setAttribute(name, value);
	return this;
};

Worksheet.prototype.freeze = function (column, row, cell, activePane) {
	this.sheetView.freeze(column, row, cell, activePane);
	return this;
};

Worksheet.prototype.split = function (x, y, cell, activePane) {
	this.sheetView.split(x, y, cell, activePane);
	return this;
};

Worksheet.prototype._prepare = function () {
	(new PrepareExport(this)).prepare();
};

Worksheet.prototype._export = function (canStream) {
	return (new Export(this)).export(canStream);
};

module.exports = Worksheet;
