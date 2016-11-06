'use strict';

var _ = require('lodash');
var idSpaces = {};
var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
var LETTER_REFS = {};

function positionToLetter(x, y) {
	var result = LETTER_REFS[x];

	if (!result) {
		var string = '';
		var num = x;
		var index;

		do {
			index = (num - 1) % 26;
			string = alphabet[index] + string;
			num = (num - (index + 1)) / 26;
		} while (num > 0);

		LETTER_REFS[x] = string;
		result = string;
	}
	return result + (y || '');
}

function letterToPosition(cell) {
	var x = 0;
	var y = 0;
	var i;
	var len;
	var charCode;

	for (i = 0, len = cell.length; i < len; i++) {
		charCode = cell.charCodeAt(i);
		if (charCode >= 65) {
			x = x * 26 + charCode - 64;
		} else {
			y = parseInt(cell.slice(i), 10);
			break;
		}
	}
	return {
		x: x || 1,
		y: y || 1
	};
}

var util = {
	/**
	 * Returns a number based on a namespace. So, running with 'Picture' will return 1. Run again, you will get 2. Run with 'Foo', you'll get 1.
	 * @param {String} space
	 * @returns {Number}
	 */
	uniqueId: function (space) {
		if (!idSpaces[space]) {
			idSpaces[space] = 1;
		}
		return idSpaces[space]++;
	},

	inherits: function (ctor, superCtor) {
		var Obj = function () {};
		Obj.prototype = superCtor.prototype;
		ctor.prototype = new Obj();
	},

	pixelsToEMUs: function (pixels) {
		return Math.round(pixels * 914400 / 96);
	},

	canonCell: function (cell) {
		if (_.isObject(cell)) {
			return positionToLetter(cell.c || 1, cell.r || 1);
		}
		return cell;
	},

	positionToLetter: positionToLetter,
	letterToPosition: letterToPosition,

	xmlPrefix: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',

	schemas: {
		'worksheet': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
		'sharedStrings': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
		'stylesheet': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
		'relationships': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
		'relationshipPackage': 'http://schemas.openxmlformats.org/package/2006/relationships',
		'contentTypes': 'http://schemas.openxmlformats.org/package/2006/content-types',
		'spreadsheetml': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
		'markupCompat': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
		'x14ac': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
		'officeDocument': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
		'package': 'http://schemas.openxmlformats.org/package/2006/relationships',
		'table': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
		'spreadsheetDrawing': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
		'drawing': 'http://schemas.openxmlformats.org/drawingml/2006/main',
		'drawingRelationship': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
		'image': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
		'chart': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
		'hyperlink': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
	}
};

module.exports = util;
