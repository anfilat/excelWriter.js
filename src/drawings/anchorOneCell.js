'use strict';

const _ = require('lodash');
const util = require('../util');
const toXMLString = require('../XMLString');

class AnchorOneCell {
	constructor(config) {
		let coord;

		if (_.isObject(config)) {
			if (_.has(config, 'cell')) {
				if (_.isObject(config.cell)) {
					coord = {x: config.cell.c || 1, y: config.cell.r || 1};
				} else {
					coord = util.letterToPosition(config.cell || '');
				}
			} else {
				coord = {x: config.c || 1, y: config.r || 1};
			}
		} else {
			coord = util.letterToPosition(config || '');
			config = {};
		}

		this.x = coord.x - 1;
		this.y = coord.y - 1;
		this.xOff = util.pixelsToEMUs(config.left || 0);
		this.yOff = util.pixelsToEMUs(config.top || 0);
		this.width = util.pixelsToEMUs(config.width || 0);
		this.height = util.pixelsToEMUs(config.height || 0);
	}
	saveWithContent(content) {
		return toXMLString({
			name: 'xdr:oneCellAnchor',
			children: [
				toXMLString({
					name: 'xdr:from',
					children: [
						toXMLString({
							name: 'xdr:col',
							value: this.x
						}),
						toXMLString({
							name: 'xdr:colOff',
							value: this.xOff
						}),
						toXMLString({
							name: 'xdr:row',
							value: this.y
						}),
						toXMLString({
							name: 'xdr:rowOff',
							value: this.yOff
						})
					]
				}),
				toXMLString({
					name: 'xdr:ext',
					attributes: [
						['cx', this.width],
						['cy', this.height]
					]
				}),
				content,
				toXMLString({
					name: 'xdr:clientData'
				})
			]
		});
	}
}

module.exports = AnchorOneCell;
