'use strict';

// https://msdn.microsoft.com/en-us/library/documentformat.openxml.drawing.spreadsheet.absoluteanchor.aspx

const util = require('../util');
const toXMLString = require('../XMLString');

function AnchorAbsolute({left = 0, top = 0, width = 0, height = 0} = {}) {
	this.x = util.pixelsToEMUs(left);
	this.y = util.pixelsToEMUs(top);
	this.width = util.pixelsToEMUs(width);
	this.height = util.pixelsToEMUs(height);
}

AnchorAbsolute.prototype = {
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
};

module.exports = AnchorAbsolute;
