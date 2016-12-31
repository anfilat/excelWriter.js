'use strict';

var Readable = require('stream').Readable;
var _ = require('lodash');
var util = require('../util');

var spaceRE = /^\s|\s$/;

function SharedStrings(common) {
	this.objectId = common.uniqueId('SharedStrings');
	this._strings = Object.create(null);
	this._stringArray = [];
}

SharedStrings.prototype.add = function (string) {
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

SharedStrings.prototype.export = function (canStream) {
	this._strings = null;

	if (canStream) {
		return new SharedStringsStream({
			strings: this._stringArray
		});
	} else {
		var len = this._stringArray.length;
		var children = _.map(this._stringArray, function (string) {
			var str = _.escape(string);

			if (spaceRE.test(str)) {
				return '<si><t xml:space="preserve">' + str + '</t></si>';
			}
			return '<si><t>' + str + '</t></si>';
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
		stop = !this.push(getXMLBegin(this.strings.length));

		this.status = 1;
		this.index = 0;
		this.len = this.strings.length;
	}

	if (this.status === 1) {
		var s = '';
		var str;

		while (this.index < this.len && !stop) {
			while (this.index < this.len && s.length < size) {
				str = _.escape(this.strings[this.index]);

				if (spaceRE.test(str)) {
					s += '<si><t xml:space="preserve">' + str + '</t></si>';
				} else {
					s += '<si><t>' + str + '</t></si>';
				}
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
