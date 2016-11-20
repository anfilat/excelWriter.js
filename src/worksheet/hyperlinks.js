'use strict';

var _ = require('lodash');
var util = require('../util');
var toXMLString = require('../XMLString');

function Hyperlinks(worksheet) {
	this.relations = worksheet.relations;

	this.hyperlinks = [];
}

Hyperlinks.prototype.set = function (hyperlink) {
	hyperlink.objectId = _.uniqueId('hyperlink');
	this.relations.addRelation({
		objectId: hyperlink.objectId,
		target: hyperlink.location,
		targetMode: hyperlink.targetMode || 'External'
	}, 'hyperlink');
	this.hyperlinks.push(hyperlink);
};

Hyperlinks.prototype.insert = function (colIndex, rowIndex, hyperlink) {
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
	this.set({
		cell: {c: colIndex + 1, r: rowIndex + 1},
		location: location,
		targetMode: targetMode,
		tooltip: tooltip
	});
};

Hyperlinks.prototype._export = function () {
	var relations = this.relations;

	if (this.hyperlinks.length > 0) {
		var children = _.map(this.hyperlinks, function (hyperlink) {
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
};

module.exports = Hyperlinks;
