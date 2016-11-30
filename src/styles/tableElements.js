'use strict';

var StylePart = require('./stylePart');
var util = require('../util');
var NumberFormats = require('./numberFormats');
var Fonts = require('./fonts');
var Fills = require('./fills');
var Borders = require('./borders');
var alignment = require('./alignment');
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.differentialformats.aspx
function TableElements(styles) {
	StylePart.call(this, styles, 'dxfs', 'tableElement');
}

util.inherits(TableElements, StylePart);

TableElements.prototype.canon = function (format) {
	var result = {};

	if (format.format) {
		result.format = NumberFormats.canon(format.format);
	}
	if (format.font) {
		result.font = Fonts.canon(format.font);
	}
	if (format.pattern) {
		result.fill = Fills.canon(format.pattern, {fillType: 'pattern', isTable: true});
	} else if (format.gradient) {
		result.fill = Fills.canon(format.gradient, {fillType: 'gradient'});
	}
	if (format.border) {
		result.border = Borders.canon(format.border);
	}
	result.alignment = alignment.canon(format);
	return result;
};

TableElements.prototype.exportFormat = function (format) {
	var children = [];

	if (format.font) {
		children.push(Fonts.exportFormat(format.font));
	}
	if (format.fill) {
		children.push(Fills.exportFormat(format.fill));
	}
	if (format.border) {
		children.push(Borders.exportFormat(format.border));
	}
	if (format.format) {
		children.push(NumberFormats.exportFormat(format.format));
	}
	if (format.alignment && format.alignment.length) {
		children.push(alignment.export(format.alignment));
	}

	return toXMLString({
		name: 'dxf',
		children: children
	});
};

module.exports = TableElements;
