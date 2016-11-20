'use strict';

var StylePart = require('./stylePart');
var util = require('../util');
var numberFormats = require('./numberFormats');
var fonts = require('./fonts');
var fills = require('./fills');
var borders = require('./borders');
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
		result.format = numberFormats.canon(format.format);
	}
	if (format.font) {
		result.font = fonts.canon(format.font);
	}
	if (format.pattern) {
		result.fill = fills.canon(format.pattern, 'pattern', 'table');
	} else if (format.gradient) {
		result.fill = fills.canon(format.gradient, 'gradient');
	}
	if (format.border) {
		result.border = borders.canon(format.border);
	}
	result.alignment = alignment.canon(format);
	return result;
};

TableElements.prototype.exportFormat = function (format) {
	var children = [];

	if (format.font) {
		children.push(fonts.exportFormat(format.font));
	}
	if (format.fill) {
		children.push(fills.exportFormat(format.fill));
	}
	if (format.border) {
		children.push(borders.exportFormat(format.border));
	}
	if (format.format) {
		children.push(numberFormats.exportFormat(format.format));
	}
	if (format.alignment && format.alignment.length) {
		children.push(alignment.export(format.alignment));
	}

	return toXMLString({
		name: 'dxf',
		children: children
	});
};

module.exports = {
	TableElements: TableElements
};
