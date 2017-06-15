'use strict';

const Readable = require('stream').Readable;
const _ = require('lodash');
const util = require('../util');

const spaceRE = /^\s|\s$/;

function SharedStrings(common) {
	this.objectId = common.uniqueId('SharedStrings');
	this.strings = Object.create(null);
	this.stringArray = [];
	this.count = 0;
}

SharedStrings.prototype = {
	add(string) {
		let stringId = this.strings[string];

		if (stringId === undefined) {
			stringId = this.count++;

			this.strings[string] = stringId;
			this.stringArray[stringId] = string;
		}

		return stringId;
	},
	isStrings() {
		return this.count > 0;
	},
	save(canStream) {
		this.strings = null;

		if (canStream) {
			return new SharedStringsStream({
				strings: this.stringArray
			});
		}
		const result = getXMLBegin(this.count) +
			this.stringArray.map(string => {
				string = _.escape(string);

				if (spaceRE.test(string)) {
					return `<si><t xml:space="preserve">${string}</t></si>`;
				}
				return `<si><t>${string}</t></si>`;
			}).join('') +
			getXMLEnd();
		this.stringArray = null;
		return result;
	}
};

class SharedStringsStream extends (Readable || null) {
	constructor(options) {
		super(options);
		this.strings = options.strings;
		this.status = 0;
		this.index = 0;
		this.len = this.strings.length;
	}
	_read(size) {
		let stop = false;

		if (this.status === 0) {
			stop = !this.push(getXMLBegin(this.len));

			this.status = 1;
		}

		if (this.status === 1) {
			while (this.index < this.len && !stop) {
				stop = !this.push(this.packChunk(size));
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
	packChunk(size) {
		let s = '';

		while (this.index < this.len && s.length < size) {
			const str = _.escape(this.strings[this.index]);

			if (spaceRE.test(str)) {
				s += `<si><t xml:space="preserve">${str}</t></si>`;
			} else {
				s += `<si><t>${str}</t></si>`;
			}
			this.strings[this.index] = null;
			this.index++;
		}
		return s;
	}
}

function getXMLBegin(length) {
	return `${util.xmlPrefix}<sst xmlns="${util.schemas.spreadsheetml}" count="${length}" uniqueCount="${length}">`;
}

function getXMLEnd() {
	return '</sst>';
}

module.exports = SharedStrings;
