'use strict';

var NumberFormats = require('./numberFormats');
var Fonts = require('./fonts');
var Fills = require('./fills');
var Borders = require('./borders');
var Cells = require('./cells');
var Tables = require('./tables');
var TableElements = require('./tableElements');
var toXMLString = require('../XMLString');

function Styles(common) {
	this.objectId = common.uniqueId('Styles');
	this.numberFormats = new NumberFormats(this);
	this.fonts = new Fonts(this);
	this.fills = new Fills(this);
	this.borders = new Borders(this);
	this.cells = new Cells(this);
	this.tableElements = new TableElements(this);
	this.tables = new Tables(this);
	this.defaultTableStyle = '';
}

Styles.prototype.addFormat = function (format, name) {
	return this.cells.add(format, null, name);
};

Styles.prototype._getId = function (name) {
	return this.cells.getId(name);
};

Styles.prototype._merge = function (columnFormat, rowFormat, cellFormat) {
	var count = Number(Boolean(columnFormat)) + Number(Boolean(rowFormat)) + Number(Boolean(cellFormat));

	if (count === 0) {
		return null;
	} else if (count === 1) {
		return this.cells.add(columnFormat || rowFormat || cellFormat);
	} else {
		var format = {};

		if (columnFormat) {
			format = this.cells.merge(format, this.cells.fullGet(columnFormat));
		}
		if (rowFormat) {
			format = this.cells.merge(format, this.cells.fullGet(rowFormat));
		}
		if (cellFormat) {
			format = this.cells.merge(format, this.cells.fullGet(cellFormat));
		}
		return this.cells.add(format);
	}
};

Styles.prototype.addFontFormat = function (format, name) {
	return this.fonts.add(format, null, name);
};

Styles.prototype.addBorderFormat = function (format, name) {
	return this.borders.add(format, null, name);
};

Styles.prototype.addPatternFormat = function (format, name) {
	return this.fills.add(format, 'pattern', name);
};

Styles.prototype.addGradientFormat = function (format, name) {
	return this.fills.add(format, 'gradient', name);
};

Styles.prototype.addNumberFormat = function (format, name) {
	return this.numberFormats.add(format, null, name);
};

Styles.prototype.addTableFormat = function (format, name) {
	return this.tables.add(format, null, name);
};

Styles.prototype.addTableElementFormat = function (format, name) {
	return this.tableElements.add(format, null, name);
};

Styles.prototype.setDefaultTableStyle = function (name) {
	this.tables.defaultTableStyle = name;
};

Styles.prototype.export = function () {
	return toXMLString({
		name: 'styleSheet',
		ns: 'spreadsheetml',
		children: [
			this.numberFormats.export(),
			this.fonts.export(),
			this.fills.export(),
			this.borders.export(),
			this.cells.export(),
			this.tableElements.export(),
			this.tables.export()
		]
	});
};

module.exports = Styles;
