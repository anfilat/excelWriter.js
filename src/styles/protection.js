'use strict';

const _ = require('lodash');
const toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.protection.aspx

function canon(format) {
	const result = {};

	if (_.has(format, 'locked')) {
		result.locked = format.locked ? 1 : 0;
	}
	if (_.has(format, 'hidden')) {
		result.hidden = format.hidden ? 1 : 0;
	}

	return _.isEmpty(result) ? null : result;
}

function merge(formatTo, formatFrom) {
	return _.assign(formatTo, formatFrom);
}

function exportFormat(format) {
	return toXMLString({
		name: 'protection',
		attributes: _.toPairs(format)
	});
}

module.exports = {
	canon,
	merge,
	exportFormat
};
