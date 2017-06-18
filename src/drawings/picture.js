'use strict';

// https://msdn.microsoft.com/en-us/library/documentformat.openxml.drawing.spreadsheet.picture.aspx

const util = require('../util');
const toXMLString = require('../XMLString');
const Anchor = require('./anchor');
const AnchorOneCell = require('./anchorOneCell');
const AnchorAbsolute = require('./anchorAbsolute');

function Picture(common, config) {
	this.pictureId = common.uniqueIdForSpace('Picture');
	this.image = config.image;
	this.imageRelationId = config.imageRelationId;
	this.createAnchor(config.anchorType, config.config);
}

Picture.prototype = {
	/**
	 *
	 * @param {String} type Can be 'anchor', 'oneCell' or 'absolute'.
	 * @param {Object} config Shorthand - pass the created anchor coords that can normally be used to construct it.
	 */
	createAnchor(type, config) {
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
	},
	save() {
		return this.anchor.saveWithContent(toXMLString({
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
		}));
	}
};

module.exports = Picture;
