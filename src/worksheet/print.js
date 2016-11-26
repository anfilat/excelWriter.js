'use strict';

var _ = require('lodash');
var toXMLString = require('../XMLString');

function Print(worksheet) {
	this.relations = worksheet.relations;

	this.headers = [];
	this.footers = [];
}

/**
 * Expects an array length of three.
 * @param {Array} headers [left, center, right]
 */
Print.prototype.setHeader = function (headers) {
	if (!_.isArray(headers)) {
		throw 'Invalid argument type - setHeader expects an array of three instructions';
	}
	this.headers = headers;
};

/**
 * Expects an array length of three.
 * @param {Array} footers [left, center, right]
 */
Print.prototype.setFooter = function (footers) {
	if (!_.isArray(footers)) {
		throw 'Invalid argument type - setFooter expects an array of three instructions';
	}
	this.footers = footers;
};

/**
 * Set page details in inches.
 */
Print.prototype.setPageMargin = function (margin) {
	this._margin = _.defaults(margin, {
		left: 0.7,
		right: 0.7,
		top: 0.75,
		bottom: 0.75,
		header: 0.3,
		footer: 0.3
	});
};

/**
 * http://www.datypic.com/sc/ooxml/t-ssml_ST_Orientation.html
 *
 * Can be one of 'portrait' or 'landscape'.
 *
 * @param {String} orientation
 */
Print.prototype.setPageOrientation = function (orientation) {
	this._orientation = orientation;
};

/**
 * Set rows to repeat for print
 *
 * @param {int|[int, int]} params - number of rows to repeat from the top | [first, last] repeat rows
 */
Print.prototype.setPrintTitleTop = function (params) {
	this._printTitles = this._printTitles || {};

	if (_.isObject(params)) {
		this._printTitles.topFrom = params[0];
		this._printTitles.topTo = params[1];
	} else {
		this._printTitles.topFrom = 0;
		this._printTitles.topTo = params - 1;
	}
};

/**
 * Set columns to repeat for print
 *
 * @param {int|[int, int]} params - number of columns to repeat from the left | [first, last] repeat columns
 */
Print.prototype.setPrintTitleLeft = function (params) {
	this._printTitles = this._printTitles || {};

	if (_.isObject(params)) {
		this._printTitles.leftFrom = params[0];
		this._printTitles.leftTo = params[1];
	} else {
		this._printTitles.leftFrom = 0;
		this._printTitles.leftTo = params - 1;
	}
};

Print.prototype.export = function () {
	return exportPageMargins(this._margin) +
		exportPageSetup(this._orientation) +
		exportHeaderFooter(this.headers, this.footers);
};

function exportPageMargins(margin) {
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

function exportPageSetup(orientation) {
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

function exportHeaderFooter(headers, footers) {
	if (headers.length > 0 || footers.length > 0) {
		var children = [];

		if (headers.length > 0) {
			children.push(exportHeader(headers));
		}
		if (footers.length > 0) {
			children.push(exportFooter(footers));
		}

		return toXMLString({
			name: 'headerFooter',
			children: children
		});
	}
	return '';
}

function exportHeader(headers) {
	return toXMLString({
		name: 'oddHeader',
		value: compilePageDetailPackage(headers)
	});
}

function exportFooter(footers) {
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
		var string = '';

		if (data.font || data.bold) {
			var weighting = data.bold ? 'Bold' : 'Regular';

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
