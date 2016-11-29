'use strict';

var _ = require('lodash');
var StylePart = require('./stylePart');
var util = require('../util');
var formatUtils = require('./utils');
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.fonts.aspx
function Fonts(styles) {
	StylePart.call(this, styles, 'fonts', 'font');

	this.init();
	this.lastId = this.formats.length;
}

util.inherits(Fonts, StylePart);

Fonts.canon = function (format) {
	var result = {};

	if (_.has(format, 'bold')) {
		result.bold = !!format.bold;
	}
	if (_.has(format, 'italic')) {
		result.italic = !!format.italic;
	}
	if (format.superscript) {
		result.vertAlign = 'superscript';
	}
	if (format.subscript) {
		result.vertAlign = 'subscript';
	}
	if (format.underline) {
		if (_.indexOf(['double', 'singleAccounting', 'doubleAccounting'], format.underline) !== -1) {
			result.underline = format.underline;
		} else {
			result.underline = true;
		}
	}
	if (_.has(format, 'strike')) {
		result.strike = !!format.strike;
	}
	if (format.outline) {
		result.outline = true;
	}
	if (format.shadow) {
		result.shadow = true;
	}
	if (format.size) {
		result.size = format.size;
	}
	if (format.color) {
		result.color = format.color;
	}
	if (format.fontName) {
		result.fontName = format.fontName;
	}
	return result;
};

Fonts.exportFormat = function (format) {
	var children = [];
	var attrs;

	if (format.size) {
		children.push(toXMLString({
			name: 'sz',
			attributes: [
				['val', format.size]
			]
		}));
	}
	if (format.fontName) {
		children.push(toXMLString({
			name: 'name',
			attributes: [
				['val', format.fontName]
			]
		}));
	}
	if (_.has(format, 'bold')) {
		if (format.bold) {
			children.push(toXMLString({name: 'b'}));
		} else {
			children.push(toXMLString({name: 'b', attributes: [['val', 0]]}));
		}
	}
	if (_.has(format, 'italic')) {
		if (format.italic) {
			children.push(toXMLString({name: 'i'}));
		} else {
			children.push(toXMLString({name: 'i', attributes: [['val', 0]]}));
		}
	}
	if (format.vertAlign) {
		children.push(toXMLString({
			name: 'vertAlign',
			attributes: [
				['val', format.vertAlign]
			]
		}));
	}
	if (format.underline) {
		attrs = null;

		if (format.underline !== true) {
			attrs = [
				['val', format.underline]
			];
		}
		children.push(toXMLString({
			name: 'u',
			attributes: attrs
		}));
	}
	if (_.has(format, 'strike')) {
		if (format.strike) {
			children.push(toXMLString({name: 'strike'}));
		} else {
			children.push(toXMLString({name: 'strike', attributes: [['val', 0]]}));
		}
	}
	if (format.shadow) {
		children.push(toXMLString({name: 'shadow'}));
	}
	if (format.outline) {
		children.push(toXMLString({name: 'outline'}));
	}
	if (format.color) {
		children.push(formatUtils.exportColor(format.color));
	}

	return toXMLString({
		name: 'font',
		children: children
	});
};

Fonts.prototype.init = function () {
	this.formats.push(
		{format: this.canon({})}
	);
};

Fonts.prototype.canon = Fonts.canon;

Fonts.prototype.merge = function (formatTo, formatFrom) {
	var result = _.assign(formatTo, formatFrom);

	result.color = formatFrom && formatFrom.color || formatTo && formatTo.color;
	return result;
};

Fonts.prototype.exportFormat = Fonts.exportFormat;

module.exports = Fonts;
