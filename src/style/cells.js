'use strict';

var _ = require('lodash');
var StylePart = require('./stylePart');
var alignment = require('./alignment');
var protection = require('./protection');
var util = require('../util');
var toXMLString = require('../XMLString');

var ALLOWED_PARTS = ['format', 'fill', 'border', 'font'];
var XLS_NAMES = ['numFmtId', 'fillId', 'borderId', 'fontId'];

function Cells(styleSheet) {
	StylePart.call(this, styleSheet, 'cellXfs', 'format');

	this.init();
	this.lastId = this.formats.length;
	this.exportEmpty = false;
}

util.inherits(Cells, StylePart);

Cells.prototype.init = function () {
	this.formats.push(
		{format: this.canon({})}
	);
};

Cells.prototype.canon = function (format) {
	var result = {};

	if (format.format) {
		result.format = this.styleSheet.numberFormats.add(format.format);
	}
	if (format.font) {
		result.font = this.styleSheet.fonts.add(format.font);
	}
	if (format.pattern) {
		result.fill = this.styleSheet.fills.add(format.pattern, 'pattern');
	} else if (format.gradient) {
		result.fill = this.styleSheet.fills.add(format.gradient, 'gradient');
	}
	if (format.border) {
		result.border = this.styleSheet.borders.add(format.border);
	}
	result.alignment = alignment.canon(format);
	result.protection = protection.canon(format);
	return result;
};

Cells.prototype.exportFormat = function (format) {
	var styleSheet = this.styleSheet;
	var attributes = [];
	var children = [];

	if (format.alignment && format.alignment.length) {
		children.push(alignment.export(format.alignment));
		attributes.push(['applyAlignment', 'true']);
	}
	if (format.protection && format.protection.length) {
		children.push(protection.export(format.protection));
		attributes.push(['applyProtection', 'true']);
	}

	_.forEach(format, function (value, key) {
		var xlsName;

		if (_.includes(ALLOWED_PARTS, key)) {
			xlsName = XLS_NAMES[_.indexOf(ALLOWED_PARTS, key)];

			if (key === 'format') {
				attributes.push([xlsName, styleSheet.numberFormats.getId(value)]);
				attributes.push(['applyNumberFormat', 'true']);
			} else if (key === 'fill') {
				attributes.push([xlsName, styleSheet.fills.getId(value)]);
				attributes.push(['applyFill', 'true']);
			} else if (key === 'border') {
				attributes.push([xlsName, styleSheet.borders.getId(value)]);
				attributes.push(['applyBorder', 'true']);
			} else if (key === 'font') {
				attributes.push([xlsName, styleSheet.fonts.getId(value)]);
				attributes.push(['applyFont', 'true']);
			}
		}
	});

	return toXMLString({
		name: 'xf',
		attributes: attributes,
		children: children
	});
};

module.exports = {
	Cells: Cells
};
