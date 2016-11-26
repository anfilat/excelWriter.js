'use strict';

var _ = require('lodash');
var Drawings = require('../drawings');
var toXMLString = require('../XMLString');

function Drawing(worksheet) {
	this.common = worksheet.common;
	this.relations = worksheet.relations;

	this.drawings = null;
}

Drawing.prototype.set = function (image, config, anchorType) {
	var name;

	if (!this.drawings) {
		this.drawings = new Drawings(this.common);

		this.common.addDrawings(this.drawings);
		this.relations.addRelation(this.drawings, 'drawingRelationship');
	}

	if (_.isObject(image)) {
		name = this.common.addImage(image.data, image.type);
	} else {
		name = image;
	}

	this.drawings.addImage(name, config, anchorType);
};

Drawing.prototype.insert = function (colIndex, rowIndex, image) {
	var config;

	if (typeof image === 'string' || image.data) {
		this.set(image, {c: colIndex + 1, r: rowIndex + 1}, 'anchor');
	} else {
		config = image.config || {};
		config.cell = {c: colIndex + 1, r: rowIndex + 1};

		this.set(image.image, config, 'anchor');
	}
};

Drawing.prototype.export = function () {
	if (this.drawings) {
		return toXMLString({
			name: 'drawing',
			attributes: [
				['r:id', this.relations.getRelationshipId(this.drawings)]
			]
		});
	}
	return '';
};

module.exports = Drawing;
