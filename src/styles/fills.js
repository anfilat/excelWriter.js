'use strict';

const _ = require('lodash');
const StylePart = require('./stylePart');
const toXMLString = require('../XMLString');

const PATTERN_TYPES = ['none', 'solid', 'darkGray', 'mediumGray', 'lightGray', 'gray125', 'gray0625',
	'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis',
	'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp',	'lightGrid', 'lightTrellis'];

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.fills.aspx
class Fills extends StylePart {
	constructor(styles) {
		super(styles, 'fills', 'fill');

		this.init();
		this.lastId = this.formats.length;
	}
	init() {
		this.formats.push(
			{format: this.canon({type: 'none'}, {fillType: 'pattern'})},
			{format: this.canon({type: 'gray125'}, {fillType: 'pattern'})}
		);
	}
	static canon(format, flags) {
		const result = {
			fillType: flags.merge ? format.fillType : flags.fillType
		};

		if (result.fillType === 'pattern') {
			const fgColor = (flags.merge ? format.fgColor : format.color) || 'FFFFFFFF';
			const bgColor = (flags.merge ? format.bgColor : format.backColor) || 'FFFFFFFF';
			const patternType = flags.merge ? format.patternType : format.type;

			result.patternType = _.includes(PATTERN_TYPES, patternType) ? patternType : 'solid';
			if (flags.isTable && result.patternType === 'solid') {
				result.fgColor = bgColor;
				result.bgColor = fgColor;
			} else {
				result.fgColor = fgColor;
				result.bgColor = bgColor;
			}
		} else {
			if (_.has(format, 'left')) {
				result.left = format.left || 0;
				result.right = format.right || 0;
				result.top = format.top || 0;
				result.bottom = format.bottom || 0;
			} else {
				result.degree = format.degree || 0;
			}
			result.start = format.start || 'FFFFFFFF';
			result.end = format.end || 'FFFFFFFF';
		}
		return result;
	}
	static saveFormat(format) {
		let children;

		if (format.fillType === 'pattern') {
			children = [savePatternFill(format)];
		} else {
			children = [saveGradientFill(format)];
		}

		return toXMLString({
			name: 'fill',
			children
		});
	}
	canon(format, flags) {
		return Fills.canon(format, flags);
	}
	merge(formatTo, formatFrom) {
		return formatFrom || formatTo;
	}
	saveFormat(format) {
		return Fills.saveFormat(format);
	}
}

function savePatternFill(format) {
	const attributes = [
		['patternType', format.patternType]
	];
	const children = [
		toXMLString({
			name: 'fgColor',
			attributes: [
				['rgb', format.fgColor]
			]
		}),
		toXMLString({
			name: 'bgColor',
			attributes: [
				['rgb', format.bgColor]
			]
		})
	];

	return toXMLString({
		name: 'patternFill',
		attributes,
		children
	});
}

function saveGradientFill(format) {
	const attributes = [];
	const children = [];
	let attrs;

	if (format.degree) {
		attributes.push(['degree', format.degree]);
	} else if (format.left) {
		attributes.push(['type', 'path']);
		attributes.push(['left', format.left]);
		attributes.push(['right', format.right]);
		attributes.push(['top', format.top]);
		attributes.push(['bottom', format.bottom]);
	}

	attrs = [['rgb', format.start]];
	children.push(toXMLString({
		name: 'stop',
		attributes: [
			['position', 0]
		],
		children : [toXMLString({name: 'color',	attributes: attrs})]
	}));

	attrs = [['rgb', format.end]];
	children.push(toXMLString({
		name: 'stop',
		attributes: [
			['position', 1]
		],
		children : [toXMLString({name: 'color',	attributes: attrs})]
	}));

	return toXMLString({
		name: 'gradientFill',
		attributes,
		children
	});
}

module.exports = Fills;
