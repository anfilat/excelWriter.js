'use strict';

// https://msdn.microsoft.com/en-us/library/documentformat.openxml.drawing.spreadsheet.twocellanchor.aspx

const _ = require('lodash');
const util = require('../util');
const toXMLString = require('../XMLString');

function Anchor(config) {
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
	const x = coord.x - 1;
	const y = coord.y - 1;

	this.from = {
		x,
		y,
		xOff: util.pixelsToEMUs(config.left || 0),
		yOff: util.pixelsToEMUs(config.top || 0)
	};
	this.to = {
		x: x + (config.cols || 1),
		y: y + (config.rows || 1),
		xOff: util.pixelsToEMUs(-config.right || 0),
		yOff: util.pixelsToEMUs(-config.bottom || 0)
	};
}

Anchor.prototype = {
	saveWithContent(content) {
		return toXMLString({
			name: 'xdr:twoCellAnchor',
			children: [
				toXMLString({
					name: 'xdr:from',
					children: [
						toXMLString({
							name: 'xdr:col',
							value: this.from.x
						}),
						toXMLString({
							name: 'xdr:colOff',
							value: this.from.xOff
						}),
						toXMLString({
							name: 'xdr:row',
							value: this.from.y
						}),
						toXMLString({
							name: 'xdr:rowOff',
							value: this.from.yOff
						})
					]
				}),
				toXMLString({
					name: 'xdr:to',
					children: [
						toXMLString({
							name: 'xdr:col',
							value: this.to.x
						}),
						toXMLString({
							name: 'xdr:colOff',
							value: this.to.xOff
						}),
						toXMLString({
							name: 'xdr:row',
							value: this.to.y
						}),
						toXMLString({
							name: 'xdr:rowOff',
							value: this.to.yOff
						})
					]
				}),
				content,
				toXMLString({
					name: 'xdr:clientData'
				})
			]
		});
	}
};

module.exports = Anchor;
