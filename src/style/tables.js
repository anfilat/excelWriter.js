'use strict';

var _ = require('lodash');
var StylePart = require('./stylePart');
var util = require('../util');
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestylevalues.aspx
var ELEMENTS = ['wholeTable', 'headerRow', 'totalRow', 'firstColumn', 'lastColumn',
	'firstRowStripe', 'secondRowStripe', 'firstColumnStripe', 'secondColumnStripe',
	'firstHeaderCell', 'lastHeaderCell', 'firstTotalCell', 'lastTotalCell'];
var SIZED_ELEMENTS = ['firstRowStripe', 'secondRowStripe', 'firstColumnStripe', 'secondColumnStripe'];

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestyles.aspx
function Tables(styleSheet) {
	StylePart.call(this, styleSheet, 'tableStyles', 'table');

	this.exportEmpty = false;
}

util.inherits(Tables, StylePart);

Tables.prototype.canon = function (format) {
	var result = {};
	var styleSheet = this.styleSheet;

	_.forEach(format, function (value, key) {
		if (_.includes(ELEMENTS, key)) {
			var style;
			var size = null;

			if (value.style) {
				style = styleSheet.tableElements.add(value.style);
				if (value.size > 1 && _.includes(SIZED_ELEMENTS, key)) {
					size = value.size;
				}
			} else {
				style = styleSheet.tableElements.add(value);
			}
			result[key] = {
				style: style,
				size: size
			};
		}
	});

	return result;
};

Tables.prototype.exportCollectionExt = function (attributes) {
	if (this.styleSheet.defaultTableStyle) {
		attributes.push(['defaultTableStyle', this.styleSheet.defaultTableStyle]);
	}
};

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestyleelement.aspx
Tables.prototype.exportFormat = function (format, styleFormat) {
	var styleSheet = this.styleSheet;
	var attributes = [
		['name', styleFormat.name],
		['pivot', 0]
	];
	var children = [];

	_.forEach(format, function (value, key) {
		var style = value.style;
		var size = value.size;
		var attributes = [
			['type', key],
			['dxfId', styleSheet.tableElements.getId(style)]
		];

		if (size) {
			attributes.push(['size', size]);
		}

		children.push(toXMLString({
			name: 'tableStyleElement',
			attributes: attributes
		}));
	});
	attributes.push(['count', children.length]);

	return toXMLString({
		name: 'tableStyle',
		attributes: attributes,
		children: children
	});
};

module.exports = {
	Tables: Tables
};
