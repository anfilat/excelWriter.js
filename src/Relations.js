'use strict';

const _ = require('lodash');
const util = require('./util');
const toXMLString = require('./XMLString');

class RelationshipManager {
	constructor(common) {
		this.common = common;

		this.relations = Object.create(null);
		this.lastId = 1;
	}
	addRelation(object, type) {
		const relation = this.relations[object.objectId];

		if (relation) {
			return relation.relationId;
		}

		const relationId = 'rId' + this.lastId++;
		this.relations[object.objectId] = {
			relationId,
			schema: util.schemas[type],
			object
		};
		return relationId;
	}
	getRelationshipId(object) {
		const relation = this.relations[object.objectId];

		return relation ? relation.relationId : null;
	}
	save() {
		const common = this.common;
		const children = _.map(this.relations, relation => {
			const attributes = [
				['Id', relation.relationId],
				['Type', relation.schema],
				['Target', relation.object.target || common.getPath(relation.object)]
			];

			if (relation.object.targetMode) {
				attributes.push(['TargetMode', relation.object.targetMode]);
			}

			return toXMLString({
				name: 'Relationship',
				attributes
			});
		});

		return toXMLString({
			name: 'Relationships',
			ns: 'relationshipPackage',
			children
		});
	}
}

module.exports = RelationshipManager;
