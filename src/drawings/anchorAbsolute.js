'use strict';

const util = require('../util');
const toXMLString = require('../XMLString');

class AnchorAbsolute {
	constructor(config)	{
		config = config || {};

		this.x = util.pixelsToEMUs(config.left || 0);
		this.y = util.pixelsToEMUs(config.top || 0);
		this.width = util.pixelsToEMUs(config.width || 0);
		this.height = util.pixelsToEMUs(config.height || 0);
	}
	saveWithContent(content) {
		return toXMLString({
			name: 'xdr:absoluteAnchor',
			children: [
				toXMLString({
					name: 'xdr:pos',
					attributes: [
						['x', this.x],
						['y', this.y]
					]
				}),
				toXMLString({
					name: 'xdr:ext',
					attributes: [
						['cx', this.width],
						['cy', this.height]
					]
				}),
				content,
				toXMLString({
					name: 'xdr:clientData'
				})
			]
		});
	}
}

module.exports = AnchorAbsolute;
