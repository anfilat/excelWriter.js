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
	this.formatsByData = Object.create(null);
	this.formatsByNames = Object.create(null);
}

StylePart.prototype.add = function (format, name, flags) {
	var canonFormat;
	var stringFormat;
	var styleFormat;

	if (name && this.formatsByNames[name]) {
		canonFormat = this.canon(format, flags);
		stringFormat = _.isObject(canonFormat) ? JSON.stringify(canonFormat) : canonFormat;

		if (stringFormat !== this.formatsByNames[name].stringFormat) {
			this._add(canonFormat, stringFormat, name);
		}
		return name;
	}

	//first argument is format name
	if (!name && _.isString(format) && this.formatsByNames[format]) {
		return format;
	}

	canonFormat = this.canon(format, flags);
	stringFormat = _.isObject(canonFormat) ? JSON.stringify(canonFormat) : canonFormat;
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
		format: canonFormat,
		stringFormat: stringFormat
	};

	this.formats.push(styleFormat);
	this.formatsByData[stringFormat] = styleFormat;
	this.formatsByNames[name] = styleFormat;

	return styleFormat;
};

StylePart.prototype.canon = function (format) {
	return format;
};

StylePart.prototype.get = function (format) {
	if (_.isString(format)) {
		var styleFormat = this.formatsByNames[format];

		return styleFormat ? styleFormat.format : format;
	}
	return format;
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

StylePart.prototype.exportCollectionExt = function () {};

StylePart.prototype.exportFormat = function () {
	return '';
};

module.exports = StylePart;
