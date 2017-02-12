'use strict';

const _ = require('lodash');
const util = require('../util');
const MergedCells = require('./mergedCells');
const toXMLString = require('../XMLString');

class Hyperlinks extends MergedCells {
	constructor() {
		super();
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
	}
	_saveHyperlinks() {
		if (this._hyperlinks.length > 0) {
			const children = _.map(this._hyperlinks, hyperlink => {
				const attributes = [
					['ref', util.canonCell(hyperlink.cell)],
					['r:id', this.relations.getRelationshipId(hyperlink)]
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
}

module.exports = Hyperlinks;
