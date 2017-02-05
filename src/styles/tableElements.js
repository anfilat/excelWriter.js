'use strict';

const StylePart = require('./stylePart');
const NumberFormats = require('./numberFormats');
const Fonts = require('./fonts');
const Fills = require('./fills');
const Borders = require('./borders');
const alignment = require('./alignment');
const toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.differentialformats.aspx
class TableElements extends StylePart {
	constructor(styles) {
		super(styles, 'dxfs', 'tableElement');
	}
	canon(format) {
		const result = {};

		if (format.format) {
			result.format = NumberFormats.canon(format.format);
		}
		if (format.font) {
			result.font = Fonts.canon(format.font);
		}
		if (format.pattern) {
			result.fill = Fills.canon(format.pattern, {fillType: 'pattern', isTable: true});
		} else if (format.gradient) {
			result.fill = Fills.canon(format.gradient, {fillType: 'gradient'});
		}
		if (format.border) {
			result.border = Borders.canon(format.border);
		}
		result.alignment = alignment.canon(format);
		return result;
	}
	saveFormat(format) {
		const children = [];

		if (format.font) {
			children.push(Fonts.saveFormat(format.font));
		}
		if (format.fill) {
			children.push(Fills.saveFormat(format.fill));
		}
		if (format.border) {
			children.push(Borders.saveFormat(format.border));
		}
		if (format.format) {
			children.push(NumberFormats.saveFormat(format.format));
		}
		if (format.alignment && format.alignment.length) {
			children.push(alignment.save(format.alignment));
		}

		return toXMLString({
			name: 'dxf',
			children
		});
	}
}

module.exports = TableElements;
