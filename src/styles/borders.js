'use strict';

const _ = require('lodash');
const StylePart = require('./stylePart');
const formatUtils = require('./utils');
const toXMLString = require('../XMLString');

const BORDERS = ['left', 'right', 'top', 'bottom', 'diagonal'];

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.borders.aspx
class Borders extends StylePart {
	constructor(styles)	{
		super(styles, 'borders', 'border');

		this.init();
		this.lastId = this.formats.length;
	}
	init() {
		this.formats.push(
			{format: this.canon({})}
		);
	}
	static canon(format) {
		const result = {};

		_.forEach(BORDERS, name => {
			const border = format[name];

			if (border) {
				result[name] = {
					style: border.style,
					color: border.color
				};
			} else {
				result[name] = {};
			}
		});
		return result;
	}
	static saveFormat(format) {
		const children = _.map(BORDERS, function (name) {
			const border = format[name];
			let attributes;
			let children;

			if (border) {
				if (border.style) {
					attributes = [['style', border.style]];
				}
				if (border.color) {
					children = [formatUtils.saveColor(border.color)];
				}
			}
			return toXMLString({
				name,
				attributes,
				children
			});
		});

		return toXMLString({
			name: 'border',
			children
		});
	}
	canon(format) {
		return Borders.canon(format);
	}
	merge(formatTo, formatFrom) {
		formatTo = formatTo || {};

		if (formatFrom) {
			_.forEach(BORDERS, name => {
				const borderFrom = formatFrom[name];

				if (borderFrom && (borderFrom.style || borderFrom.color)) {
					formatTo[name] = {
						style: borderFrom.style,
						color: borderFrom.color
					};
				} else if (!formatTo[name]) {
					formatTo[name] = {};
				}
			});
		}
		return formatTo;
	}
	saveFormat(format, styleFormat) {
		return Borders.saveFormat(format, styleFormat);
	}
}

module.exports = Borders;
