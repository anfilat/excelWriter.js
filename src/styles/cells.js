'use strict';

const _ = require('lodash');
const StylePart = require('./stylePart');
const alignment = require('./alignment');
const protection = require('./protection');
const toXMLString = require('../XMLString');

const ALLOWED_PARTS = ['format', 'fill', 'border', 'font'];
const XLS_NAMES = ['numFmtId', 'fillId', 'borderId', 'fontId'];

class Cells extends StylePart {
	constructor(styles) {
		super(styles, 'cellXfs', 'format');

		this.init();
		this.lastId = this.formats.length;
		this.saveEmpty = false;
	}
	init() {
		this.formats.push(
			{format: this.canon({})}
		);
	}
	canon(format, flags) {
		const result = {};

		if (format.format) {
			result.format = this.styles.numberFormats.add(format.format);
		}
		if (format.font) {
			result.font = this.styles.fonts.add(format.font);
		}
		if (format.pattern) {
			result.fill = this.styles.fills.add(format.pattern, null, {fillType: 'pattern'});
		} else if (format.gradient) {
			result.fill = this.styles.fills.add(format.gradient, null, {fillType: 'gradient'});
		} else if (flags && flags.merge && format.fill) {
			result.fill = this.styles.fills.add(format.fill, null, flags);
		}
		if (format.border) {
			result.border = this.styles.borders.add(format.border);
		}
		const alignmentValue = alignment.canon(format);
		if (alignmentValue) {
			result.alignment = alignmentValue;
		}
		const protectionValue = protection.canon(format);
		if (protectionValue) {
			result.protection = protectionValue;
		}
		if (format.fillOut) {
			result.fillOut = format.fillOut;
		}
		return result;
	}
	fullGet(format) {
		let result = {};

		if (this.getId(format)) {
			format = this.get(format);
			if (format.format) {
				result.format = this.styles.numberFormats.get(format.format);
			}
			if (format.font) {
				result.font = _.clone(this.styles.fonts.get(format.font));
			}
			if (format.fill) {
				result.fill = _.clone(this.styles.fills.get(format.fill));
			}
			if (format.border) {
				result.border = _.clone(this.styles.borders.get(format.border));
			}
			if (format.alignment) {
				result.alignment = _.clone(format.alignment);
			}
			if (format.protection) {
				result.protection = _.clone(format.protection);
			}
		} else {
			result = this.canon(format);
		}
		return result;
	}
	cutVisible(format) {
		const result = {};

		if (format.format) {
			result.format = format.format;
		}
		if (format.font) {
			result.font = format.font;
		}
		if (format.alignment) {
			result.alignment = format.alignment;
		}
		if (format.protection) {
			result.protection = format.protection;
		}
		return result;
	}
	merge(formatTo, formatFrom) {
		if (formatTo.format || formatFrom.format) {
			formatTo.format = this.styles.numberFormats.merge(formatTo.format, formatFrom.format);
		}
		if (formatTo.font || formatFrom.font) {
			formatTo.font = this.styles.fonts.merge(formatTo.font, formatFrom.font);
		}
		if (formatTo.fill || formatFrom.fill) {
			formatTo.fill = this.styles.fills.merge(formatTo.fill, formatFrom.fill);
		}
		if (formatTo.border || formatFrom.border) {
			formatTo.border = this.styles.borders.merge(formatTo.border, formatFrom.border);
		}
		if (formatTo.alignment || formatFrom.alignment) {
			formatTo.alignment = alignment.merge(formatTo.alignment, formatFrom.alignment);
		}
		if (formatTo.protection || formatFrom.protection) {
			formatTo.protection = protection.merge(formatTo.protection, formatFrom.protection);
		}
		return formatTo;
	}
	saveFormat(format) {
		const styles = this.styles;
		const attributes = [];
		const children = [];

		if (format.alignment) {
			children.push(alignment.saveFormat(format.alignment));
			attributes.push(['applyAlignment', 'true']);
		}
		if (format.protection) {
			children.push(protection.saveFormat(format.protection));
			attributes.push(['applyProtection', 'true']);
		}

		_.forEach(format, (value, key) => {
			if (_.includes(ALLOWED_PARTS, key)) {
				const xlsName = XLS_NAMES[_.indexOf(ALLOWED_PARTS, key)];

				if (key === 'format') {
					attributes.push([xlsName, styles.numberFormats.getId(value)]);
					attributes.push(['applyNumberFormat', 'true']);
				} else if (key === 'fill') {
					attributes.push([xlsName, styles.fills.getId(value)]);
					attributes.push(['applyFill', 'true']);
				} else if (key === 'border') {
					attributes.push([xlsName, styles.borders.getId(value)]);
					attributes.push(['applyBorder', 'true']);
				} else if (key === 'font') {
					attributes.push([xlsName, styles.fonts.getId(value)]);
					attributes.push(['applyFont', 'true']);
				}
			}
		});

		return toXMLString({
			name: 'xf',
			attributes,
			children
		});
	}
}

module.exports = Cells;
