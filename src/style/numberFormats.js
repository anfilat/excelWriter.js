'use strict';

var _ = require('lodash');
var StylePart = require('./stylePart');
var util = require('../util');
var toXMLString = require('../XMLString');

var PREDEFINED = {
	date: 14, //mm-dd-yy
	time: 21  //h:mm:ss
};

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformats.aspx
function NumberFormats(styleSheet) {
	StylePart.call(this, styleSheet, 'numFmts', 'numberFormat');

	this.init();
	this.lastId = 164;
}

util.inherits(NumberFormats, StylePart);

NumberFormats.prototype.init = function () {
	var self = this;

	_.forEach(PREDEFINED, function (formatId, format) {
		self.formatsByNames[format] = {formatId: formatId};
	});
};

NumberFormats.prototype.canon = canon;
NumberFormats.prototype.exportFormat = exportFormat;

function canon(format) {
	return format;
}

function exportFormat(format, styleFormat) {
	var attributes = [
		['numFmtId', styleFormat.formatId],
		['formatCode', format]
	];

	return toXMLString({
		name: 'numFmt',
		attributes: attributes
	});
}

module.exports = {
	NumberFormats: NumberFormats,
	canon: canon,
	exportFormat: exportFormat
};