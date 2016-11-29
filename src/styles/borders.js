'use strict';

var _ = require('lodash');
var StylePart = require('./stylePart');
var util = require('../util');
var formatUtils = require('./utils');
var toXMLString = require('../XMLString');

var BORDERS = ['left', 'right', 'top', 'bottom', 'diagonal'];

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.borders.aspx
function Borders(styles) {
	StylePart.call(this, styles, 'borders', 'border');

	this.init();
	this.lastId = this.formats.length;
}

util.inherits(Borders, StylePart);

Borders.canon = function (format) {
	var result = {};

	_.forEach(BORDERS, function (name) {
		var border = format[name];

		if (border) {
			result[name] = {
				style: border.style,
				color: border.color
			};
		} else {
			result[name] = {};
		}
	});
	return result;
};

Borders.exportFormat = function (format) {
	var children = _.map(BORDERS, function (name) {
		var border = format[name];
		var attributes;
		var children;

		if (border) {
			if (border.style) {
				attributes = [['style', border.style]];
			}
			if (border.color) {
				children = [formatUtils.exportColor(border.color)];
			}
		}
		return toXMLString({
			name: name,
			attributes: attributes,
			children: children
		});
	});

	return toXMLString({
		name: 'border',
		children: children
	});
};

Borders.prototype.init = function () {
	this.formats.push(
		{format: this.canon({})}
	);
};

Borders.prototype.canon = Borders.canon;

Borders.prototype.merge = function (formatTo, formatFrom) {
	formatTo = formatTo || {};

	if (formatFrom) {
		_.forEach(BORDERS, function (name) {
			var borderFrom = formatFrom[name];

			if (borderFrom && (borderFrom.style || borderFrom.color)) {
				formatTo[name] = {
					style: borderFrom.style,
					color: borderFrom.color
				};
			} else if (!formatTo[name]) {
				formatTo[name] = {};
			}
		});
	}
	return formatTo;
};

Borders.prototype.exportFormat = Borders.exportFormat;

module.exports = Borders;
