'use strict';

const _ = require('lodash');
const toXMLString = require('../XMLString');

function exportColor(color) {
	if (_.isString(color)) {
		return toXMLString({
			name: 'color',
			attributes: [
				['rgb', color]
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
	exportColor: exportColor
};
