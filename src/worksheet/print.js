'use strict';

const _ = require('lodash');
const toXMLString = require('../XMLString');

class Print {
	constructor() {
		this._headers = [];
		this._footers = [];
	}
	/**
	 * Expects an array length of three.
	 * @param {Array} headers [left, center, right]
	 */
	setHeader(headers) {
		if (!_.isArray(headers)) {
			throw 'Invalid argument type - setHeader expects an array of three instructions';
		}
		this._headers = headers;
		return this;
	}
	/**
	 * Expects an array length of three.
	 * @param {Array} footers [left, center, right]
	 */
	setFooter(footers) {
		if (!_.isArray(footers)) {
			throw 'Invalid argument type - setFooter expects an array of three instructions';
		}
		this._footers = footers;
		return this;
	}
	/**
	 * Set page details in inches.
	 */
	setPageMargin(margin) {
		this._margin = _.defaults(margin, {
			left: 0.7,
			right: 0.7,
			top: 0.75,
			bottom: 0.75,
			header: 0.3,
			footer: 0.3
		});
		return this;
	}
	/**
	 * http://www.datypic.com/sc/ooxml/t-ssml_ST_Orientation.html
	 *
	 * Can be one of 'portrait' or 'landscape'.
	 *
	 * @param {String} orientation
	 */
	setPageOrientation(orientation) {
		this._orientation = orientation;
		return this;
	}
	/**
	 * Set rows to repeat for print
	 *
	 * @param {int|[int, int]} params - number of rows to repeat from the top | [first, last] repeat rows
	 */
	setPrintTitleTop(params) {
		this._printTitles = this._printTitles || {};

		if (_.isObject(params)) {
			this._printTitles.topFrom = params[0];
			this._printTitles.topTo = params[1];
		} else {
			this._printTitles.topFrom = 0;
			this._printTitles.topTo = params - 1;
		}
		return this;
	}
	/**
	 * Set columns to repeat for print
	 *
	 * @param {int|[int, int]} params - number of columns to repeat from the left | [first, last] repeat columns
	 */
	setPrintTitleLeft(params) {
		this._printTitles = this._printTitles || {};

		if (_.isObject(params)) {
			this._printTitles.leftFrom = params[0];
			this._printTitles.leftTo = params[1];
		} else {
			this._printTitles.leftFrom = 0;
			this._printTitles.leftTo = params - 1;
		}
		return this;
	}
	_savePrint() {
		return savePageMargins(this._margin) +
			savePageSetup(this._orientation) +
			saveHeaderFooter(this._headers, this._footers);
	}
}

function savePageMargins(margin) {
	if (margin) {
		return toXMLString({
			name: 'pageMargins',
			attributes: [
				['top', margin.top],
				['bottom', margin.bottom],
				['left', margin.left],
				['right', margin.right],
				['header', margin.header],
				['footer', margin.footer]
			]
		});
	}
	return '';
}

function savePageSetup(orientation) {
	if (orientation) {
		return toXMLString({
			name: 'pageSetup',
			attributes: [
				['orientation', orientation]
			]
		});
	}
	return '';
}

function saveHeaderFooter(headers, footers) {
	if (headers.length > 0 || footers.length > 0) {
		const children = [];

		if (headers.length > 0) {
			children.push(saveHeader(headers));
		}
		if (footers.length > 0) {
			children.push(saveFooter(footers));
		}

		return toXMLString({
			name: 'headerFooter',
			children
		});
	}
	return '';
}

function saveHeader(headers) {
	return toXMLString({
		name: 'oddHeader',
		value: compilePageDetailPackage(headers)
	});
}

function saveFooter(footers) {
	return toXMLString({
		name: 'oddFooter',
		value: compilePageDetailPackage(footers)
	});
}

function compilePageDetailPackage(data) {
	data = data || '';

	return [
		'&L', compilePageDetailPiece(data[0] || ''),
		'&C', compilePageDetailPiece(data[1] || ''),
		'&R', compilePageDetailPiece(data[2] || '')
	].join('');
}

function compilePageDetailPiece(data) {
	if (_.isString(data)) {
		return '&"-,Regular"'.concat(data);
	} else if (_.isObject(data) && !_.isArray(data)) {
		let string = '';

		if (data.font || data.bold) {
			const weighting = data.bold ? 'Bold' : 'Regular';

			string += '&"' + (data.font || '-') + ',' + weighting + '"';
		} else {
			string += '&"-,Regular"';
		}
		if (data.underline) {
			string += '&U';
		}
		if (data.fontSize) {
			string += '&' + data.fontSize;
		}
		string += data.text;

		return string;
	} else if (_.isArray(data)) {
		return _.reduce(data, function (result, value) {
			return result.concat(compilePageDetailPiece(value));
		}, '');
	}
}

module.exports = Print;
