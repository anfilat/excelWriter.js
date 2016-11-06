'use strict';

var _ = require('lodash');
var util = require('./util');
var RelationshipManager = require('./relationshipManager');
var Picture = require('./picture');
var toXMLString = require('./XMLString');

function Drawings(common) {
	this.objectId = _.uniqueId('Drawings');
	this.drawings = [];
	this.relations = new RelationshipManager(common);
	this.common = common;
}

Drawings.prototype.addImage = function (name, config, anchorType) {
	var image = this.common.getImage(name);
	var imageRelationId = this.relations.addRelation(image, 'image');
	var picture = new Picture({
		image: image,
		imageRelationId: imageRelationId,
		config: config,
		anchorType: anchorType
	});

	this.drawings.push(picture);
};

Drawings.prototype._export = function () {
	var attributes = [
		['xmlns:a', util.schemas.drawing],
		['xmlns:r', util.schemas.relationships],
		['xmlns:xdr', util.schemas.spreadsheetDrawing]
	];
	var children = _.map(this.drawings, function (picture) {
		return picture._export();
	});

	return toXMLString({
		name: 'xdr:wsDr',
		ns: 'spreadsheetDrawing',
		attributes: attributes,
		children: children
	});
};

module.exports = Drawings;
