'use strict';

var _ = require('lodash');
var StylePart = require('./stylePart');
var util = require('../util');
var toXMLString = require('../XMLString');

var PATTERN_TYPES = ['none', 'solid', 'darkGray', 'mediumGray', 'lightGray', 'gray125', 'gray0625',
	'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis',
	'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp',	'lightGrid', 'lightTrellis'];

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.fills.aspx
function Fills(styles) {
	StylePart.call(this, styles, 'fills', 'fill');

	this.init();
	this.lastId = this.formats.length;
}

util.inherits(Fills, StylePart);

Fills.canon = function (format, flags) {
	var result = {
		fillType: flags.merge ? format.fillType : flags.fillType
	};

	if (result.fillType === 'pattern') {
		var fgColor = (flags.merge ? format.fgColor : format.color) || 'FFFFFFFF';
		var bgColor = (flags.merge ? format.bgColor : format.backColor) || 'FFFFFFFF';
		var patternType = flags.merge ? format.patternType : format.type;

		result.patternType = _.includes(PATTERN_TYPES, patternType) ? patternType : 'solid';
		if (flags.isTable && result.patternType === 'solid') {
			result.fgColor = bgColor;
			result.bgColor = fgColor;
		} else {
			result.fgColor = fgColor;
			result.bgColor = bgColor;
		}
	} else {
		if (_.has(format, 'left')) {
			result.left = format.left || 0;
			result.right = format.right || 0;
			result.top = format.top || 0;
			result.bottom = format.bottom || 0;
		} else {
			result.degree = format.degree || 0;
		}
		result.start = format.start || 'FFFFFFFF';
		result.end = format.end || 'FFFFFFFF';
	}
	return result;
};

Fills.exportFormat = function (format) {
	var children;

	if (format.fillType === 'pattern') {
		children = [exportPatternFill(format)];
	} else {
		children = [exportGradientFill(format)];
	}

	return toXMLString({
		name: 'fill',
		children: children
	});
};

function exportPatternFill(format) {
	var attributes = [
		['patternType', format.patternType]
	];
	var children = [
		toXMLString({
			name: 'fgColor',
			attributes: [
				['rgb', format.fgColor]
			]
		}),
		toXMLString({
			name: 'bgColor',
			attributes: [
				['rgb', format.bgColor]
			]
		})
	];

	return toXMLString({
		name: 'patternFill',
		attributes: attributes,
		children: children
	});
}

function exportGradientFill(format) {
	var attributes = [];
	var children = [];
	var attrs;

	if (format.degree) {
		attributes.push(['degree', format.degree]);
	} else if (format.left) {
		attributes.push(['type', 'path']);
		attributes.push(['left', format.left]);
		attributes.push(['right', format.right]);
		attributes.push(['top', format.top]);
		attributes.push(['bottom', format.bottom]);
	}

	attrs = [['rgb', format.start]];
	children.push(toXMLString({
		name: 'stop',
		attributes: [
			['position', 0]
		],
		children : [toXMLString({name: 'color',	attributes: attrs})]
	}));

	attrs = [['rgb', format.end]];
	children.push(toXMLString({
		name: 'stop',
		attributes: [
			['position', 1]
		],
		children : [toXMLString({name: 'color',	attributes: attrs})]
	}));

	return toXMLString({
		name: 'gradientFill',
		attributes: attributes,
		children: children
	});
}

Fills.prototype.init = function () {
	this.formats.push({
		format: this.canon({type: 'none'}, {fillType: 'pattern'})
	}, {
		format: this.canon({type: 'gray125'}, {fillType: 'pattern'})
	});
};

Fills.prototype.canon = Fills.canon;

Fills.prototype.merge = function (formatTo, formatFrom) {
	return formatFrom || formatTo;
};

Fills.prototype.exportFormat = Fills.exportFormat;

module.exports = Fills;
