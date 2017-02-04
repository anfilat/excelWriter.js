'use strict';

const _ = require('lodash');
const toXMLString = require('../XMLString');

class StylePart {
	constructor(styles, exportName, formatName) {
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
	add(format, name, flags) {
		if (name && this.formatsByNames[name]) {
			const canonFormat = this.canon(format, flags);
			const stringFormat = _.isObject(canonFormat) ? JSON.stringify(canonFormat) : canonFormat;

			if (stringFormat !== this.formatsByNames[name].stringFormat) {
				this._add(canonFormat, stringFormat, name);
			}
			return name;
		}

		//first argument is format name
		if (!name && _.isString(format) && this.formatsByNames[format]) {
			return format;
		}

		const canonFormat = this.canon(format, flags);
		const stringFormat = _.isObject(canonFormat) ? JSON.stringify(canonFormat) : canonFormat;
		let styleFormat = this.formatsByData[stringFormat];

		if (!styleFormat) {
			styleFormat = this._add(canonFormat, stringFormat, name);
		} else if (name && !this.formatsByNames[name]) {
			styleFormat.name = name;
			this.formatsByNames[name] = styleFormat;
		}
		return styleFormat.name;
	}
	_add(canonFormat, stringFormat, name) {
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
	}
	canon(format) {
		return format;
	}
	get(format) {
		if (_.isString(format)) {
			const styleFormat = this.formatsByNames[format];

			return styleFormat ? styleFormat.format : format;
		}
		return format;
	}
	getId(name) {
		const styleFormat = this.formatsByNames[name];

		return styleFormat ? styleFormat.formatId : null;
	}
	export() {
		if (this.exportEmpty !== false || this.formats.length) {
			const attributes = [
				['count', this.formats.length]
			];
			const children = _.map(this.formats, format => this.exportFormat(format.format, format));

			this.exportCollectionExt(attributes, children);

			return toXMLString({
				name: this.exportName,
				attributes,
				children
			});
		}
		return '';
	}
	exportCollectionExt() {}
	exportFormat() {
		return '';
	}
}

module.exports = StylePart;
