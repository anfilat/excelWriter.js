'use strict';

var _ = require('lodash');
var util = require('../util');
var toXMLString = require('../XMLString');

module.exports = {
	init: function () {
		this._hyperlinks = [];
	},
	methods: {
		setHyperlink: function (hyperlink) {
			hyperlink.objectId = this.common.uniqueId('hyperlink');
			this.relations.addRelation({
				objectId: hyperlink.objectId,
				target: hyperlink.location,
				targetMode: hyperlink.targetMode || 'External'
			}, 'hyperlink');
			this._hyperlinks.push(hyperlink);
			return this;
		},
		_insertHyperlink: function (colIndex, rowIndex, hyperlink) {
			var location;
			var targetMode;
			var tooltip;

			if (typeof hyperlink === 'string') {
				location = hyperlink;
			} else {
				location = hyperlink.location;
				targetMode = hyperlink.targetMode;
				tooltip = hyperlink.tooltip;
			}
			this.setHyperlink({
				cell: {c: colIndex + 1, r: rowIndex + 1},
				location: location,
				targetMode: targetMode,
				tooltip: tooltip
			});
		},
		_exportHyperlinks: function () {
			var relations = this.relations;

			if (this._hyperlinks.length > 0) {
				var children = _.map(this._hyperlinks, function (hyperlink) {
					var attributes = [
						['ref', util.canonCell(hyperlink.cell)],
						['r:id', relations.getRelationshipId(hyperlink)]
					];

					if (hyperlink.tooltip) {
						attributes.push(['tooltip', hyperlink.tooltip]);
					}
					return toXMLString({
						name: 'hyperlink',
						attributes: attributes
					});
				});

				return toXMLString({
					name: 'hyperlinks',
					children: children
				});
			}
			return '';
		}
	}
};
