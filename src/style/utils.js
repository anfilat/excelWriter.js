'use strict';

var _ = require('lodash');
var toXMLString = require('../XMLString');

function exportColor(color) {
	var attributes;

	if (_.isString(color)) {
		return toXMLString({
			name: 'color',
			attributes: [
				['rgb', color]
			]
		});
	} else {
		attributes = [];
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
			attributes: attributes
		});
	}
}

module.exports = {
	exportColor: exportColor
};
