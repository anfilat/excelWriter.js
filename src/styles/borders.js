'use strict';

const _ = require('lodash');
const StylePart = require('./stylePart');
const {saveColor} = require('./utils');
const toXMLString = require('../XMLString');

const MAIN_BORDERS = ['left', 'right', 'top', 'bottom'];
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

		if (_.has(format, 'style') || _.has(format, 'color')) {
			MAIN_BORDERS.forEach(name => {
				result[name] = {
					style: format.style,
					color: format.color
				};
			});
		} else {
			BORDERS.forEach(name => {
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
		}
		return result;
	}
	static saveFormat(format) {
		const children = BORDERS.map(name => {
			const border = format[name];
			let attributes;
			let children;

			if (border) {
				if (border.style) {
					attributes = [['style', border.style]];
				}
				if (border.color) {
					children = [saveColor(border.color)];
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
	merge(formatTo = Borders.canon({}), formatFrom) {
		if (formatFrom) {
			BORDERS.forEach(name => {
				const borderFrom = formatFrom[name];

				if (borderFrom && borderFrom.style) {
					formatTo[name].style = borderFrom.style;
				}
				if (borderFrom && borderFrom.color) {
					formatTo[name].color = borderFrom.color;
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
