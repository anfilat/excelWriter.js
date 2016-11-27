'use strict';

var _ = require('lodash');
var toXMLString = require('../XMLString');

function StylePart(styles, exportName, formatName) {
	this.styles = styles;
	this.exportName = exportName;
	this.formatName = formatName;
	this.lastName = 1;
	this.lastId = 0;
	this.exportEmpty = true;
	this.formats = [];
	this.formatsByData = {};
	this.formatsByNames = {};
}

StylePart.prototype.add = function (format, type, name) {
	var canonFormat;
	var stringFormat;
	var styleFormat;

	if (name && this.formatsByNames[name]) {
		canonFormat = this.canon(format, type);
		if (_.isObject(canonFormat)) {
			stringFormat = JSON.stringify(canonFormat);
		} else {
			stringFormat = canonFormat;
		}

		if (stringFormat !== this.formatsByNames[name].format) {
			this._add(canonFormat, stringFormat, name);
		}
		return name;
	}

	//first argument is format name
	if (!name && _.isString(format) && this.formatsByNames[format]) {
		return format;
	}

	canonFormat = this.canon(format, type || format.fillType);
	if (_.isObject(canonFormat)) {
		stringFormat = JSON.stringify(canonFormat);
	} else {
		stringFormat = canonFormat;
	}
	styleFormat = this.formatsByData[stringFormat];

	if (!styleFormat) {
		styleFormat = this._add(canonFormat, stringFormat, name);
	} else if (name && !this.formatsByNames[name]) {
		styleFormat.name = name;
		this.formatsByNames[name] = styleFormat;
	}
	return styleFormat.name;
};

StylePart.prototype._add = function (canonFormat, stringFormat, name) {
	name = name || this.formatName + this.lastName++;

	var styleFormat = {
		name: name,
		formatId: this.lastId++,
		format: canonFormat
	};

	this.formats.push(styleFormat);
	this.formatsByData[stringFormat] = styleFormat;
	this.formatsByNames[name] = styleFormat;

	return styleFormat;
};

StylePart.prototype.get = function (name) {
	var styleFormat = this.formatsByNames[name];

	return styleFormat ? styleFormat.format : null;
};

StylePart.prototype.getId = function (name) {
	var styleFormat = this.formatsByNames[name];

	return styleFormat ? styleFormat.formatId : null;
};

StylePart.prototype.export = function () {
	if (this.exportEmpty !== false || this.formats.length) {
		var self = this;
		var attributes = [
			['count', this.formats.length]
		];
		var children = _.map(this.formats, function (format) {
			return self.exportFormat(format.format, format);
		});

		this.exportCollectionExt(attributes, children);

		return toXMLString({
			name: this.exportName,
			attributes: attributes,
			children: children
		});
	}
	return '';
};

StylePart.prototype.canon = function (format) {
	return format;
};

StylePart.prototype.merge = function (formatTo, formatFrom) {
	return this.add(this._merge(this.get(formatTo), this.get(formatFrom)));
};

StylePart.prototype.exportCollectionExt = function () {};

StylePart.prototype.exportFormat = function () {
	return '';
};

module.exports = StylePart;
