'use strict';

var util = require('../util');
var toXMLString = require('../XMLString');
var Anchor = require('./anchor');
var AnchorOneCell = require('./anchorOneCell');
var AnchorAbsolute = require('./anchorAbsolute');

function Picture(common, config) {
	this.pictureId = common.uniqueIdSeparated('Picture').id;
	this.image = config.image;
	this.imageRelationId = config.imageRelationId;
	this.createAnchor(config.anchorType, config.config);
}

/**
 *
 * @param {String} type Can be 'anchor', 'oneCell' or 'absolute'.
 * @param {Object} config Shorthand - pass the created anchor coords that can normally be used to construct it.
 */
Picture.prototype.createAnchor = function (type, config) {
	switch (type) {
		case 'anchor':
			this.anchor = new Anchor(config);
			break;
		case 'oneCell':
			this.anchor = new AnchorOneCell(config);
			break;
		case 'absolute':
			this.anchor = new AnchorAbsolute(config);
			break;
	}
};

Picture.prototype.export = function () {
	var picture = toXMLString({
		name: 'xdr:pic',
		children: [
			toXMLString({
				name: 'xdr:nvPicPr',
				children: [
					toXMLString({
						name: 'xdr:cNvPr',
						attributes: [
							['id', this.pictureId],
							['name', this.image.name]
						]
					}),
					toXMLString({
						name: 'xdr:cNvPicPr',
						children: [
							toXMLString({
								name: 'a:picLocks',
								attributes: [
									['noChangeAspect', '1'],
									['noChangeArrowheads', '1']
								]
							})
						]
					})
				]
			}),
			toXMLString({
				name: 'xdr:blipFill',
				children: [
					toXMLString({
						name: 'a:blip',
						attributes: [
							['xmlns:r', util.schemas.relationships],
							['r:embed', this.imageRelationId]
						]
					}),
					toXMLString({
						name: 'a:srcRect'
					}),
					toXMLString({
						name: 'a:stretch',
						children: [
							toXMLString({
								name: 'a:fillRect'
							})
						]
					})
				]
			}),
			toXMLString({
				name: 'xdr:spPr',
				attributes: [
					['bwMode', 'auto']
				],
				children: [
					toXMLString({
						name: 'a:xfrm'
					}),
					toXMLString({
						name: 'a:prstGeom',
						attributes: [
							['prst', 'rect']
						]
					})
				]
			})
		]
	});

	return this.anchor.exportWithContent(picture);
};

module.exports = Picture;
