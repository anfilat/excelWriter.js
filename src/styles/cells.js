'use strict';

var _ = require('lodash');
var StylePart = require('./stylePart');
var alignment = require('./alignment');
var protection = require('./protection');
var util = require('../util');
var toXMLString = require('../XMLString');

var ALLOWED_PARTS = ['format', 'fill', 'border', 'font'];
var XLS_NAMES = ['numFmtId', 'fillId', 'borderId', 'fontId'];

function Cells(styles) {
	StylePart.call(this, styles, 'cellXfs', 'format');

	this.init();
	this.lastId = this.formats.length;
	this.exportEmpty = false;
}

util.inherits(Cells, StylePart);

Cells.prototype.init = function () {
	this.formats.push(
		{format: this.canon({})}
	);
};

Cells.prototype.canon = function (format, flags) {
	var result = {};

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
	result.alignment = alignment.canon(format);
	result.protection = protection.canon(format);
	return result;
};

Cells.prototype.fullGet = function (format) {
	var result = {};

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
};

Cells.prototype.merge = function (formatTo, formatFrom) {
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
};

Cells.prototype.exportFormat = function (format) {
	var styles = this.styles;
	var attributes = [];
	var children = [];

	if (format.alignment) {
		children.push(alignment.exportFormat(format.alignment));
		attributes.push(['applyAlignment', 'true']);
	}
	if (format.protection) {
		children.push(protection.exportFormat(format.protection));
		attributes.push(['applyProtection', 'true']);
	}

	_.forEach(format, function (value, key) {
		var xlsName;

		if (_.includes(ALLOWED_PARTS, key)) {
			xlsName = XLS_NAMES[_.indexOf(ALLOWED_PARTS, key)];

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
		attributes: attributes,
		children: children
	});
};

module.exports = Cells;
