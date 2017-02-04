'use strict';

const Readable = require('stream').Readable;
const _ = require('lodash');
const util = require('../util');

const spaceRE = /^\s|\s$/;

class SharedStrings extends Readable {
	constructor(options) {
		super(options);
		this.objectId = options.common.uniqueId('SharedStrings');
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
	export(canStream) {
		this._strings = null;

		if (canStream) {
			this.status = 0;
			return this;
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
	_read(size) {
		let stop = false;

		if (this.status === 0) {
			stop = !this.push(getXMLBegin(this._stringArray.length));

			this.status = 1;
			this.index = 0;
			this.len = this._stringArray.length;
		}

		if (this.status === 1) {
			let s = '';
			let str;

			while (this.index < this.len && !stop) {
				while (this.index < this.len && s.length < size) {
					str = _.escape(this._stringArray[this.index]);

					if (spaceRE.test(str)) {
						s += `<si><t xml:space="preserve">${str}</t></si>`;
					} else {
						s += `<si><t>${str}</t></si>`;
					}
					this._stringArray[this.index] = null;
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
