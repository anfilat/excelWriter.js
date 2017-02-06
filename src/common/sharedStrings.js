/*eslint global-require: "off"*/

'use strict';

const Readable = require('stream').Readable || null;
const _ = require('lodash');
const util = require('../util');

const spaceRE = /^\s|\s$/;

class SharedStrings {
	constructor(common) {
		this.objectId = common.uniqueId('SharedStrings');
		this._strings = new Map();
		this._stringArray = [];
	}
	add(string) {
		if (!this._strings.has(string)) {
			const stringId = this._stringArray.length;

			this._strings.set(string, stringId);
			this._stringArray[stringId] = string;
		}

		return this._strings.get(string);
	}
	isEmpty() {
		return this._stringArray.length === 0;
	}
	save(canStream) {
		this._strings = null;

		if (canStream) {
			return new SharedStringsStream({
				strings: this._stringArray
			});
		}
		return getXMLBegin(this._stringArray.length) +
			_.map(this._stringArray, string => {
				string = _.escape(string);

				if (spaceRE.test(string)) {
					return `<si><t xml:space="preserve">${string}</t></si>`;
				}
				return `<si><t>${string}</t></si>`;
			}).join('') +
			getXMLEnd();
	}
}

class SharedStringsStream extends Readable {
	constructor(options) {
		super(options);
		this.strings = options.strings;
		this.status = 0;
	}
	_read(size) {
		let stop = false;

		if (this.status === 0) {
			stop = !this.push(getXMLBegin(this.strings.length));

			this.status = 1;
			this.index = 0;
			this.len = this.strings.length;
		}

		if (this.status === 1) {
			let s = '';
			let str;

			while (this.index < this.len && !stop) {
				while (this.index < this.len && s.length < size) {
					str = _.escape(this.strings[this.index]);

					if (spaceRE.test(str)) {
						s += `<si><t xml:space="preserve">${str}</t></si>`;
					} else {
						s += `<si><t>${str}</t></si>`;
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
	}
}

function getXMLBegin(length) {
	return `${util.xmlPrefix}<sst xmlns="${util.schemas.spreadsheetml}" count="${length}" uniqueCount="${length}">`;
}

function getXMLEnd() {
	return '</sst>';
}

module.exports = SharedStrings;
