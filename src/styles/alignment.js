'use strict';

var _ = require('lodash');
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.alignment.aspx
//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.horizontalalignmentvalues.aspx
var HORIZONTAL = ['left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.verticalalignmentvalues.aspx
var VERTICAL = ['top', 'center', 'bottom', 'justify', 'distributed'];

function canon(format) {
	var result = [];

	if (format.horizontal && _.includes(HORIZONTAL, format.horizontal)) {
		result.push(['horizontal', format.horizontal]);
	}
	if (format.vertical && _.includes(VERTICAL, format.vertical)) {
		result.push(['vertical', format.vertical]);
	}
	if (format.indent) {
		result.push(['indent', format.indent]);
	}
	if (format.justifyLastLine) {
		result.push(['justifyLastLine', 1]);
	}
	if (_.has(format, 'readingOrder') && _.includes([0, 1, 2], format.readingOrder)) {
		result.push(['readingOrder', format.readingOrder]);
	}
	if (format.relativeIndent) {
		result.push(['relativeIndent', format.relativeIndent]);
	}
	if (format.shrinkToFit) {
		result.push(['shrinkToFit', 1]);
	}
	if (format.textRotation) {
		result.push(['textRotation', format.textRotation]);
	}
	if (format.wrapText) {
		result.push(['wrapText', 1]);
	}

	return result.length ? result : undefined;
}

function exportAlignment(attributes) {
	return toXMLString({
		name: 'alignment',
		attributes: attributes
	});
}

module.exports = {
	canon: canon,
	export: exportAlignment
};
