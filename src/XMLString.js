'use strict';

const _ = require('lodash');
const util = require('./util');

/**
 * @param {{
 *  name: string,
 *  value?: *,
 *  children?: Array.<string>,
 *  attributes?: Array.<string, *>,
 *  ns?: string
 * }} config
 */
function toXMLString(config) {
	let string = '<' + config.name;
	let content = '';

	if (config.ns) {
		string = util.xmlPrefix + string + ' xmlns="' + util.schemas[config.ns] + '"';
	}
	if (config.attributes) {
		config.attributes.forEach(([key, value]) => {
			string += ` ${key}="${_.escape(value)}"`;
		});
	}
	if (config.value !== undefined) {
		content += _.escape(config.value);
	}
	if (config.children) {
		content += config.children.join('');
	}
	if (content) {
		string += '>' + content + '</' + config.name + '>';
	} else {
		string += '/>';
	}

	return string;
}

module.exports = toXMLString;
