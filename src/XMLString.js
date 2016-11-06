'use strict';

var _ = require('lodash');
var util = require('./util');

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
	var name = config.name;
	var string = '<' + name;
	var content = '';
	var attr;
	var i, l;

	if (config.ns) {
		string = util.xmlPrefix + string +	' xmlns="' + util.schemas[config.ns] + '"';
	}
	if (config.attributes) {
		for (i = 0, l = config.attributes.length; i < l; i++) {
			attr = config.attributes[i];

			string += ' ' + attr[0] + '="' + _.escape(attr[1]) + '"';
		}
	}
	if (!_.isUndefined(config.value)) {
		content += _.escape(config.value);
	}
	if (config.children) {
		content += config.children.join('');
	}

	if (content) {
		string += '>' + content + '</' + name + '>';
	} else {
		string += '/>';
	}

	return string;
}

module.exports = toXMLString;
