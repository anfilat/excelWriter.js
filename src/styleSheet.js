'use strict';

var _ = require('lodash');
var numberFormats = require('./style/numberFormats');
var fonts = require('./style/fonts');
var fills = require('./style/fills');
var borders = require('./style/borders');
var cells = require('./style/cells');
var tables = require('./style/tables');
var tableElements = require('./style/tableElements');
var toXMLString = require('./XMLString');

function StyleSheet() {
	this.objectId = _.uniqueId('StyleSheet');
	this.numberFormats = new numberFormats.NumberFormats(this);
	this.fonts = new fonts.Fonts(this);
	this.fills = new fills.Fills(this);
	this.borders = new borders.Borders(this);
	this.cells = new cells.Cells(this);
	this.tableElements = new tableElements.TableElements(this);
	this.tables = new tables.Tables(this);
	this.defaultTableStyle = '';
}

StyleSheet.prototype.addFormat = function (format, name) {
	return this.cells.add(format, null, name);
};

StyleSheet.prototype.addFontFormat = function (format, name) {
	return this.fonts.add(format, null, name);
};

StyleSheet.prototype.addBorderFormat = function (format, name) {
	return this.borders.add(format, null, name);
};

StyleSheet.prototype.addPatternFormat = function (format, name) {
	return this.fills.add(format, 'pattern', name);
};

StyleSheet.prototype.addGradientFormat = function (format, name) {
	return this.fills.add(format, 'gradient', name);
};

StyleSheet.prototype.addNumberFormat = function (format, name) {
	return this.numberFormats.add(format, null, name);
};

StyleSheet.prototype.addTableFormat = function (format, name) {
	return this.tables.add(format, null, name);
};

StyleSheet.prototype.addTableElementFormat = function (format, name) {
	return this.tableElements.add(format, null, name);
};

StyleSheet.prototype.setDefaultTableStyle = function (name) {
	this.tables.defaultTableStyle = name;
};

StyleSheet.prototype._export = function () {
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

module.exports = StyleSheet;
