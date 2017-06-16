'use strict';

const _ = require('lodash');
const Drawings = require('../drawings');
const toXMLString = require('../XMLString');

function WorksheetDrawings(common, relations) {
	this.common = common;
	this.relations = relations;
	this.drawings = null;
}

WorksheetDrawings.prototype = {
	setImage(image, config) {
		this.setDrawing(image, config, 'anchor');
	},
	setImageOneCell(image, config) {
		this.setDrawing(image, config, 'oneCell');
	},
	setImageAbsolute(image, config) {
		this.setDrawing(image, config, 'absolute');
	},
	insert(colIndex, rowIndex, image) {
		if (image) {
			const cell = {c: colIndex + 1, r: rowIndex + 1};

			if (typeof image === 'string' || image.data) {
				this.setDrawing(image, cell, 'anchor');
			} else {
				const config = image.config || {};
				config.cell = cell;

				this.setDrawing(image.image, config, 'anchor');
			}
		}
	},

	setDrawing(image, config, anchorType) {
		if (!this.drawings) {
			this.drawings = new Drawings(this.common);

			this.common.addDrawings(this.drawings);
			this.relations.add(this.drawings, 'drawingRelationship');
		}

		const name = _.isObject(image)
			? this.common.images.add(image.data, image.type)
			: image;

		this.drawings.add(name, config, anchorType);
	},

	save() {
		if (this.drawings) {
			return toXMLString({
				name: 'drawing',
				attributes: [
					['r:id', this.relations.getId(this.drawings)]
				]
			});
		}
		return '';
	}
};

module.exports = WorksheetDrawings;
