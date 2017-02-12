'use strict';

const _ = require('lodash');
const Tables = require('./tables');
const Drawings = require('../drawings');
const toXMLString = require('../XMLString');

class DrawingsExt extends Tables {
	constructor() {
		super();
		this._drawings = null;
	}
	setImage(image, config) {
		return this._setDrawing(image, config, 'anchor');
	}
	setImageOneCell(image, config) {
		return this._setDrawing(image, config, 'oneCell');
	}
	setImageAbsolute(image, config) {
		return this._setDrawing(image, config, 'absolute');
	}
	_setDrawing(image, config, anchorType) {
		let name;

		if (!this._drawings) {
			this._drawings = new Drawings(this.common);

			this.common.addDrawings(this._drawings);
			this.relations.addRelation(this._drawings, 'drawingRelationship');
		}

		if (_.isObject(image)) {
			name = this.common.images.addImage(image.data, image.type);
		} else {
			name = image;
		}

		this._drawings.addImage(name, config, anchorType);
		return this;
	}
	_insertDrawing(colIndex, rowIndex, image) {
		if (image) {
			const cell = {c: colIndex + 1, r: rowIndex + 1};

			if (typeof image === 'string' || image.data) {
				this._setDrawing(image, cell, 'anchor');
			} else {
				const config = image.config || {};
				config.cell = cell;

				this._setDrawing(image.image, config, 'anchor');
			}
		}
	}
	_saveDrawing() {
		if (this._drawings) {
			return toXMLString({
				name: 'drawing',
				attributes: [
					['r:id', this.relations.getRelationshipId(this._drawings)]
				]
			});
		}
		return '';
	}
}

module.exports = DrawingsExt;
