'use strict';

var Readable = require('stream').Readable;
var _ = require('lodash');
var util = require('./util');

function SharedStrings() {
	this.objectId = _.uniqueId('SharedStrings');
	this._strings = {};
	this._stringArray = [];
}

/**
 * Adds a string to the shared string file, and returns the ID of the
 * string which can be used to reference it in worksheets.
 *
 * @param string {String}
 * @return int
 */
SharedStrings.prototype.addString = function (string) {
	var stringId = this._strings[string];

	if (stringId === undefined) {
		stringId = this._stringArray.length;

		this._strings[string] = stringId;
		this._stringArray[stringId] = string;
	}

	return stringId;
};

SharedStrings.prototype.isEmpty = function () {
	return this._stringArray.length === 0;
};

SharedStrings.prototype._export = function (canStream) {
	if (canStream) {
		return new SharedStringsStream({
			strings: this._stringArray
		});
	} else {
		this._strings = null;

		var len = this._stringArray.length;
		var children = _.map(this._stringArray, function (string) {
			return '<si><t>' + _.escape(string) + '</t></si>';
		});

		return getXMLBegin(len) + children.join('') + getXMLEnd();
	}
};

function SharedStringsStream(options) {
	Readable.call(this, options);

	this.strings = options.strings;
	this.status = 0;
}

util.inherits(SharedStringsStream, Readable || {});

SharedStringsStream.prototype._read = function (size) {
	var stop = false;

	if (this.status === 0) {
		this._strings = null;
		stop = !this.push(getXMLBegin(this.strings.length));

		this.status = 1;
		this.index = 0;
		this.len = this.strings.length;
	}

	if (this.status === 1) {
		var s = '';
		while (this.index < this.len && !stop) {
			while (this.index < this.len && s.length < size) {
				s += '<si><t>' + _.escape(this.strings[this.index]) + '</t></si>';
				this.strings[this.index] = null;
				this.index++;
			}
			stop = !this.push(s);
			s = '';
		}

		if (this.index === this.len) {
			this.status = 2;
		}
	}

	if (this.status === 2) {
		this.push(getXMLEnd());
		this.push(null);
	}
};

function getXMLBegin(length) {
	return util.xmlPrefix + '<sst xmlns="' + util.schemas.spreadsheetml +
		'" count="' + length + '" uniqueCount="' + length + '">';
}

function getXMLEnd() {
	return '</sst>';
}

module.exports = SharedStrings;
