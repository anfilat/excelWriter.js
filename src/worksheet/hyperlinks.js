'use strict';

const _ = require('lodash');
const util = require('../util');
const toXMLString = require('../XMLString');

module.exports = SuperClass => class Hyperlinks extends SuperClass {
	constructor(workbook, config) {
		super(workbook, config);
		this._hyperlinks = [];
	}
	setHyperlink(hyperlink) {
		hyperlink.objectId = this.common.uniqueId('hyperlink');
		this.relations.addRelation({
			objectId: hyperlink.objectId,
			target: hyperlink.location,
			targetMode: hyperlink.targetMode || 'External'
		}, 'hyperlink');
		this._hyperlinks.push(hyperlink);
		return this;
	}
	_insertHyperlink(colIndex, rowIndex, hyperlink) {
		let location;
		let targetMode;
		let tooltip;

		if (typeof hyperlink === 'string') {
			location = hyperlink;
		} else {
			location = hyperlink.location;
			targetMode = hyperlink.targetMode;
			tooltip = hyperlink.tooltip;
		}
		this.setHyperlink({
			cell: {c: colIndex + 1, r: rowIndex + 1},
			location,
			targetMode,
			tooltip
		});
	}
	_saveHyperlinks() {
		const relations = this.relations;

		if (this._hyperlinks.length > 0) {
			const children = _.map(this._hyperlinks, function (hyperlink) {
				const attributes = [
					['ref', util.canonCell(hyperlink.cell)],
					['r:id', relations.getRelationshipId(hyperlink)]
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
