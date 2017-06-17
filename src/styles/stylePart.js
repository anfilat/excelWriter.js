'use strict';

const _ = require('lodash');
const toXMLString = require('../XMLString');

function StylePart(styles, saveName, formatName) {
	this.styles = styles;
	this.saveName = saveName;
	this.formatName = formatName;
	this.lastName = 1;
	this.lastId = 0;
	this.saveEmpty = true;
	this.formats = [];
	this.formatsByData = Object.create(null);
	this.formatsByNames = Object.create(null);
	this.predefined = {};
}

StylePart.prototype = {
	init() {},
	add(format, name, flags) {
		if (name && this.formatsByNames[name]) {
			const canonFormat = this.canon(format, flags);
			const stringFormat = _.isObject(canonFormat) ? JSON.stringify(canonFormat) : canonFormat;

			if (stringFormat !== this.formatsByNames[name].stringFormat) {
				this.addNew(canonFormat, stringFormat, name);
			}
			return name;
		}

		//first argument is format name
		if (!name && _.isString(format)) {
			if (this.formatsByNames[format]) {
				return format;
			} else if (this.predefined[format]) {
				return this.add(this.predefined[format], format);
			}
		}

		const canonFormat = this.canon(format, flags);
		const stringFormat = _.isObject(canonFormat) ? JSON.stringify(canonFormat) : canonFormat;
		let styleFormat = this.formatsByData[stringFormat];

		if (!styleFormat) {
			styleFormat = this.addNew(canonFormat, stringFormat, name);
		} else if (name && !this.formatsByNames[name]) {
			styleFormat.name = name;
			this.formatsByNames[name] = styleFormat;
		}
		return styleFormat.name;
	},
	addNew(canonFormat, stringFormat, name) {
		name = name || this.formatName + this.lastName++;

		const styleFormat = {
			name,
			formatId: this.lastId++,
			format: canonFormat,
			stringFormat
		};

		this.formats.push(styleFormat);
		this.formatsByData[stringFormat] = styleFormat;
		this.formatsByNames[name] = styleFormat;

		return styleFormat;
	},
	canon(format) {
		return format;
	},
	get(format) {
		if (_.isString(format)) {
			const styleFormat = this.formatsByNames[format];

			return styleFormat ? styleFormat.format : format;
		}
		return format;
	},
	getId(name) {
		const styleFormat = this.formatsByNames[name];

		return styleFormat ? styleFormat.formatId : this.getPredefinedId(name);
	},
	getPredefinedId(name) {
		if (this.predefined[name]) {
			return this.getId(this.add(this.predefined[name], name));
		}
		return null;
	},
	save() {
		if (this.saveEmpty !== false || this.formats.length) {
			const attributes = [
				['count', this.formats.length]
			];
			const children = this.formats.map(format => this.saveFormat(format.format, format));

			this.saveCollectionExt(attributes, children);

			return toXMLString({
				name: this.saveName,
				attributes,
				children
			});
		}
		return '';
	},
	saveCollectionExt() {},
	saveFormat() {
		return '';
	}
};

module.exports = StylePart;
