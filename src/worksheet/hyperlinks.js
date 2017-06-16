'use strict';

const util = require('../util');
const toXMLString = require('../XMLString');

function Hyperlinks(common, relations) {
	this.common = common;
	this.relations = relations;
	this.hyperlinks = [];
}

Hyperlinks.prototype = {
	setHyperlink(hyperlink) {
		hyperlink.objectId = this.common.uniqueId('hyperlink');
		this.relations.add({
			objectId: hyperlink.objectId,
			target: hyperlink.location,
			targetMode: hyperlink.targetMode || 'External'
		}, 'hyperlink');
		this.hyperlinks.push(hyperlink);
	},
	insert(colIndex, rowIndex, hyperlink) {
		if (hyperlink) {
			const cell = {c: colIndex + 1, r: rowIndex + 1};

			if (typeof hyperlink === 'string') {
				this.setHyperlink({
					cell,
					location: hyperlink
				});
			} else {
				this.setHyperlink({
					cell,
					location: hyperlink.location,
					targetMode: hyperlink.targetMode,
					tooltip: hyperlink.tooltip
				});
			}
		}
	},
	save() {
		if (this.hyperlinks.length > 0) {
			const children = this.hyperlinks.map(hyperlink => {
				const attributes = [
					['ref', util.canonCell(hyperlink.cell)],
					['r:id', this.relations.getId(hyperlink)]
				];

				if (hyperlink.tooltip) {
					attributes.push(['tooltip', hyperlink.tooltip]);
				}
				return toXMLString({
					name: 'hyperlink',
					attributes
				});
			});

			return toXMLString({
				name: 'hyperlinks',
				children
			});
		}
		return '';
	}
};

module.exports = Hyperlinks;
