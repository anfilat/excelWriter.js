'use strict';

const _ = require('lodash');
const StylePart = require('./stylePart');
const toXMLString = require('../XMLString');

const PREDEFINED = {
	date: 14, //mm-dd-yy
	time: 21  //h:mm:ss
};

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformats.aspx
function NumberFormats(styles) {
	StylePart.call(this, styles, 'numFmts', 'numberFormat');

	this.init();
	this.lastId = 164;
}

NumberFormats.canon = function (format) {
	return format;
};

NumberFormats.saveFormat = function (format, styleFormat) {
	const attributes = [
		['numFmtId', styleFormat.formatId],
		['formatCode', format]
	];

	return toXMLString({
		name: 'numFmt',
		attributes
	});
};

NumberFormats.prototype = _.merge({}, StylePart.prototype, {
	init() {
		_.forEach(PREDEFINED, (formatId, format) => {
			this.formatsByNames[format] = {
				formatId: formatId,
				format: format
			};
		});
	},
	canon: NumberFormats.canon,
	saveFormat: NumberFormats.saveFormat,
	merge(formatTo, formatFrom) {
		return formatFrom || formatTo;
	}
});

module.exports = NumberFormats;
