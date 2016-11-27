'use strict';

var _ = require('lodash');
var util = require('./util');
var toXMLString = require('./XMLString');

function RelationshipManager(common) {
	this.common = common;

	this.relations = Object.create(null);
	this.lastId = 1;
}

RelationshipManager.prototype.addRelation = function (object, type) {
	var relation = this.relations[object.objectId];
	var relationId;

	if (relation) {
		relationId = relation.relationId;
	} else {
		relationId = 'rId' + this.lastId++;

		this.relations[object.objectId] = {
			relationId: relationId,
			schema: util.schemas[type],
			object: object
		};
	}
	return relationId;
};

RelationshipManager.prototype.getRelationshipId = function (object) {
	var relation = this.relations[object.objectId];

	return relation ? relation.relationId : null;
};

RelationshipManager.prototype.export = function () {
	var common = this.common;
	var children = _.map(this.relations, function (relation) {
		var attributes = [
			['Id', relation.relationId],
			['Type', relation.schema],
			['Target', relation.object.target || common.getPath(relation.object)]
		];

		if (relation.object.targetMode) {
			attributes.push(['TargetMode', relation.object.targetMode]);
		}

		return toXMLString({
			name: 'Relationship',
			attributes: attributes
		});
	});

	return toXMLString({
		name: 'Relationships',
		ns: 'relationshipPackage',
		children: children
	});
};

module.exports = RelationshipManager;
