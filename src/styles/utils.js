'use strict';

const _ = require('lodash');
const toXMLString = require('../XMLString');

function canonColor(color) {
	return color[0] === '#' ? 'FF' + color.substr(1) : color;
}

function saveColor(color) {
	if (_.isString(color)) {
		return toXMLString({
			name: 'color',
			attributes: [
				['rgb', canonColor(color)]
			]
		});
	} else {
		const attributes = [];
		if (!_.isUndefined(color.tint)) {
			attributes.push(['tint', color.tint]);
		}
		if (!_.isUndefined(color.auto)) {
			attributes.push(['auto', !!color.auto]);
		}
		if (!_.isUndefined(color.theme)) {
			attributes.push(['theme', color.theme]);
		}

		return toXMLString({
			name: 'color',
			attributes
		});
	}
}

module.exports = {
	canonColor,
	saveColor
};
