'use strict';

const Relations = require('../relations');
const util = require('../util');
const toXMLString = require('../XMLString');
const Picture = require('./picture');

function Drawings(common) {
	this.common = common;

	this.objectId = this.common.uniqueId('Drawings');
	this.drawings = [];
	this.relations = new Relations(common);
}

Drawings.prototype = {
	addImage(name, config, anchorType) {
		const image = this.common.images.getImage(name);
		const imageRelationId = this.relations.add(image, 'image');
		const picture = new Picture(this.common, {
			image,
			imageRelationId,
			config,
			anchorType
		});

		this.drawings.push(picture);
	},
	save() {
		const attributes = [
			['xmlns:a', util.schemas.drawing],
			['xmlns:r', util.schemas.relationships],
			['xmlns:xdr', util.schemas.spreadsheetDrawing]
		];
		const children = this.drawings.map(picture => picture.save());

		return toXMLString({
			name: 'xdr:wsDr',
			ns: 'spreadsheetDrawing',
			attributes,
			children
		});
	}
};

module.exports = Drawings;
