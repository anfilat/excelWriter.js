'use strict';

var _ = require('lodash');
var Relations = require('../relations');
var util = require('../util');
var toXMLString = require('../XMLString');
var Picture = require('./picture');

function Drawings(common) {
	this.common = common;

	this.objectId = this.common.uniqueId('Drawings');
	this.drawings = [];
	this.relations = new Relations(common);
}

Drawings.prototype.addImage = function (name, config, anchorType) {
	var image = this.common.getImage(name);
	var imageRelationId = this.relations.addRelation(image, 'image');
	var picture = new Picture(this.common, {
		image: image,
		imageRelationId: imageRelationId,
		config: config,
		anchorType: anchorType
	});

	this.drawings.push(picture);
};

Drawings.prototype.export = function () {
	var attributes = [
		['xmlns:a', util.schemas.drawing],
		['xmlns:r', util.schemas.relationships],
		['xmlns:xdr', util.schemas.spreadsheetDrawing]
	];
	var children = _.map(this.drawings, function (picture) {
		return picture.export();
	});

	return toXMLString({
		name: 'xdr:wsDr',
		ns: 'spreadsheetDrawing',
		attributes: attributes,
		children: children
	});
};

module.exports = Drawings;
