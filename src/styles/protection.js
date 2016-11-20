'use strict';

var _ = require('lodash');
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.protection.aspx

function canon(format) {
	var result = [];

	if (_.has(format, 'locked') && !format.locked) {
		result.push(['locked', 0]);
	}
	if (format.hidden) {
		result.push(['hidden', 1]);
	}

	return result.length ? result : undefined;
}

function exportProtection(attributes) {
	return toXMLString({
		name: 'protection',
		attributes: attributes
	});
}

module.exports = {
	canon: canon,
	export: exportProtection
};
