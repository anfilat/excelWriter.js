'use strict';

var _ = require('lodash');
var Drawings = require('../drawings');
var toXMLString = require('../XMLString');

module.exports = {
	init: function () {
		this._drawings = null;
	},
	methods: {
		setImage: function (image, config) {
			return this._setDrawing(image, config, 'anchor');
		},
		setImageOneCell: function (image, config) {
			return this._setDrawing(image, config, 'oneCell');
		},
		setImageAbsolute: function (image, config) {
			return this._setDrawing(image, config, 'absolute');
		},
		_setDrawing: function (image, config, anchorType) {
			var name;

			if (!this._drawings) {
				this._drawings = new Drawings(this.common);

				this.common.addDrawings(this._drawings);
				this.relations.addRelation(this._drawings, 'drawingRelationship');
			}

			if (_.isObject(image)) {
				name = this.common.addImage(image.data, image.type);
			} else {
				name = image;
			}

			this._drawings.addImage(name, config, anchorType);
			return this;
		},
		_insertDrawing: function (colIndex, rowIndex, image) {
			var config;

			if (typeof image === 'string' || image.data) {
				this._setDrawing(image, {c: colIndex + 1, r: rowIndex + 1}, 'anchor');
			} else {
				config = image.config || {};
				config.cell = {c: colIndex + 1, r: rowIndex + 1};

				this._setDrawing(image.image, config, 'anchor');
			}
		},
		_exportDrawing: function () {
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
};
