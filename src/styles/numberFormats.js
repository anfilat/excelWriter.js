'use strict';

const _ = require('lodash');
const StylePart = require('./stylePart');
const toXMLString = require('../XMLString');

const PREDEFINED = {
	date: 14, //mm-dd-yy
	time: 21  //h:mm:ss
};

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformats.aspx
class NumberFormats extends StylePart {
	constructor(styles) {
		super(styles, 'numFmts', 'numberFormat');

		this.init();
		this.lastId = 164;
	}
	init() {
		_.forEach(PREDEFINED, (formatId, format) => {
			this.formatsByNames[format] = {
				formatId: formatId,
				format: format
			};
		});
	}
	static canon(format) {
		return format;
	}
	static exportFormat(format, styleFormat) {
		const attributes = [
			['numFmtId', styleFormat.formatId],
			['formatCode', format]
		];

		return toXMLString({
			name: 'numFmt',
			attributes
		});
	}
	canon(format) {
		return NumberFormats.canon(format);
	}
	merge(formatTo, formatFrom) {
		return formatFrom || formatTo;
	}
	exportFormat(format, styleFormat) {
		return NumberFormats.exportFormat(format, styleFormat);
	}
}

module.exports = NumberFormats;
